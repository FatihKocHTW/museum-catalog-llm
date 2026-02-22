import os
import re
import ast
import time
import random
import base64
import glob
from typing import List, Dict, Any, Optional

import pandas as pd
import requests
from dotenv import load_dotenv


# ================== Pfade/Dateien ==================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LINKED_CSV = os.path.join(BASE_DIR, "linked.csv")
IMAGE_DIR = os.path.join(BASE_DIR, "Images", "Objektbilder")
EXCEL_DIR = os.path.join(BASE_DIR, "Excel-Files", "Objektdaten")
OUT_XLSX = os.path.join(BASE_DIR, "catalog_ablation_results.xlsx")

# Welche Excel-Dateien + Zeilenbereiche (wie beim Linking)
EXCEL_SPECS = [
    ("Liste_AK Kommunikation.xls", (300, 600)),      # 301–600
    ("Liste_AEG Produktsammlung.xls", (2500, 3000)), # 2501–3000
]

# ===== Mapping: Excel-Header -> sprechende Labels (nur für Prompt-Ausgabe) =====
LABELS: Dict[str, str] = {
    "t1":  "Inventar-Nr",
    "t2":  "Hersteller",
    "t3":  "Materialien",
    "t4":  "Maße",
    "t5":  "Beschreibung (Excel)",
    "t6":  "Kategorie",
    "t7":  "Bezeichnung",
    "t8":  "Standort / Depot",
    "t9":  "Provenienz",
    "t10": "Zustand",
    "t11": "Signatur",
    "t12": "Inventarnr. (alt)",
    "t13": "Bildangabe (Excel)",
    "t14": "Datierung / Jahr",
}

# ===== Rate Limit / Retries =====
MAX_RETRIES = 3
BASE_BACKOFF = 2.0
PAUSE_BETWEEN_OBJECTS = 2.0
MAX_IMAGES_PER_CALL = 5  # kontrolliert Kosten/429

# ================== OpenAI Setup ==================
load_dotenv()
API_KEY = os.getenv("OPENAI_API_KEY")
if not API_KEY:
    raise ValueError("❌ Kein OPENAI_API_KEY gefunden! Bitte .env prüfen.")

OPENAI_URL = "https://api.openai.com/v1/chat/completions"
OPENAI_MODEL = "gpt-4o-mini"

# Für Vergleichbarkeit (Text-only vs. Text+Bild) sinnvoll:
TEMPERATURE = 0.0  # möglichst deterministisch
MAX_TOKENS = 700


# ================== Utils ==================
PARENS_RE = re.compile(r"\s*\([^)]*\)")

def strip_parens(val: Any) -> str:
    """Entfernt (...) und wandelt NaN/None in leere Strings."""
    if val is None:
        return ""
    s = str(val).strip()
    if not s or s.lower() == "nan":
        return ""
    return PARENS_RE.sub("", s).strip()

def read_excel_any(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    try:
        if ext == ".xlsx":
            return pd.read_excel(path, engine="openpyxl")
        elif ext == ".xls":
            try:
                return pd.read_excel(path)  # Auto
            except Exception:
                return pd.read_excel(path, engine="xlrd")  # Fallback
        else:
            raise RuntimeError(f"Unsupported extension: {ext}")
    except Exception as e:
        print(f"WARN: {os.path.basename(path)} nicht lesbar ({e}).")
        return pd.DataFrame()

def _norm_year(y: Any) -> str:
    s = "" if y is None else str(y)
    m = re.search(r"\b(\d{4})\b", s)
    return m.group(1) if m else ""

def find_file_anywhere(root: str, year: str, fname: str) -> Optional[str]:
    """Suche primär in <root>/<year>/fname, dann <root>/fname, dann rekursiv."""
    year = _norm_year(year)
    if year:
        p = os.path.join(root, year, fname)
        if os.path.exists(p):
            return p
    p2 = os.path.join(root, fname)
    if os.path.exists(p2):
        return p2
    hits = glob.glob(os.path.join(root, "**", fname), recursive=True)
    return hits[0] if hits else None

def encode_image(image_path: str) -> str:
    with open(image_path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")

def _parse_images_field(field: Any) -> List[str]:
    """Akzeptiert echte Listen, String-Listen '[]', oder Komma-Strings."""
    if field is None or (isinstance(field, float) and pd.isna(field)):
        return []
    if isinstance(field, list):
        return [str(x).strip() for x in field if str(x).strip()]
    s = str(field).strip()
    if not s:
        return []
    if s.startswith("[") and s.endswith("]"):
        try:
            val = ast.literal_eval(s)
            if isinstance(val, list):
                return [str(x).strip() for x in val if str(x).strip()]
        except Exception:
            pass
    parts = [t.strip().strip("'\"[]") for t in s.split(",")]
    return [p for p in parts if p]

# ================== Prompt ==================
def build_prompt_from_tcols(meta_row: Dict[str, Any]) -> str:
    """
    Prompt nutzt ausschließlich Excel-Spalten (t1..t14):
    Für jede vorhandene Spalte:  'LABEL [Tn]: Wert'
    """
    lines = []
    for tcol in [f"t{i}" for i in range(1, 15)]:
        if tcol in meta_row:
            label = LABELS.get(tcol, tcol.upper())
            val = strip_parens(meta_row.get(tcol, ""))
            if val:
                lines.append(f"- {label} [{tcol.upper()}]: {val}")

    meta_block = "\n".join(lines) if lines else "- (keine Metadaten gefunden)"

    return f"""Du bist Museums-Kurator:in.
Erstelle eine präzise, wissenschaftlich klingende technische Objektbeschreibung auf Deutsch.
Schreibe sachlich, ohne Marketing-Sprache.
Vermeide Phrasen wie „Dieses Bild zeigt…“.
Beziehe dich auf überprüfbare Merkmale (Materialien, Konstruktion, Maße falls erkennbar, Funktion/Nutzung, ggf. vorsichtiger Kontext).

Regeln:
- Verwende die Metadaten als Kontext, ohne sie 1:1 aufzuzählen.
- Wenn etwas nicht sicher ist, formuliere vorsichtig (z.B. „vermutlich“, „wirkt wie“).
- Keine erfundenen Herstellernamen/Datierungen.

Excel-Metadaten (Label + Originalspalte):
{meta_block}
"""

# ================== API mit Retries ==================
def _post_with_retries(payload: dict) -> dict:
    headers = {"Authorization": f"Bearer {API_KEY}", "Content-Type": "application/json"}
    for attempt in range(MAX_RETRIES):
        try:
            resp = requests.post(OPENAI_URL, headers=headers, json=payload, timeout=180)
            if resp.status_code == 429:
                retry_after = resp.headers.get("Retry-After")
                wait = float(retry_after) if retry_after else BASE_BACKOFF * (2 ** attempt) + random.uniform(0, 0.5)
                print(f"429 Too Many Requests – warte {wait:.1f}s (Versuch {attempt+1}/{MAX_RETRIES})")
                time.sleep(wait)
                continue
            resp.raise_for_status()
            return resp.json()
        except requests.exceptions.RequestException as e:
            wait = BASE_BACKOFF * (2 ** attempt) + random.uniform(0, 0.5)
            print(f"WARN: API-Fehler ({e}) – backoff {wait:.1f}s (Versuch {attempt+1}/{MAX_RETRIES})")
            time.sleep(wait)
    raise RuntimeError("Abbruch nach MAX_RETRIES wegen wiederholter 429/Netzwerkfehler.")

def call_openai_text_only(prompt_text: str) -> str:
    payload = {
        "model": OPENAI_MODEL,
        "temperature": TEMPERATURE,
        "messages": [{
            "role": "user",
            "content": [{"type": "text", "text": prompt_text}]
        }],
        "max_tokens": MAX_TOKENS
    }
    data = _post_with_retries(payload)
    return data["choices"][0]["message"]["content"]

def call_openai_with_images(image_paths: List[str], prompt_text: str) -> str:
    to_send = image_paths[:MAX_IMAGES_PER_CALL]
    image_contents = []
    for p in to_send:
        try:
            b64 = encode_image(p)
            image_contents.append({
                "type": "image_url",
                "image_url": {"url": f"data:image/jpeg;base64,{b64}"}
            })
        except Exception as e:
            print(f"WARN: Konnte Bild nicht lesen ({p}): {e}")

    payload = {
        "model": OPENAI_MODEL,
        "temperature": TEMPERATURE,
        "messages": [{
            "role": "user",
            "content": [{"type": "text", "text": prompt_text}, *image_contents]
        }],
        "max_tokens": MAX_TOKENS
    }
    data = _post_with_retries(payload)
    return data["choices"][0]["message"]["content"]

# ================== Excel-Metadaten laden ==================
def load_meta_rows_from_excels(excel_dir: str) -> pd.DataFrame:
    """
    Lädt ausschließlich definierte Zeilenbereiche.
    Behält Header als t1..t14 (lowercase) und säubert Werte.
    """
    parts = []
    want_lower = {f"t{i}" for i in range(1, 15)}  # t1..t14

    for fname, (start, stop) in EXCEL_SPECS:
        path = os.path.join(excel_dir, fname)
        if not os.path.exists(path):
            alt = os.path.splitext(path)[0] + ".xlsx"
            path = alt if os.path.exists(alt) else path

        df = read_excel_any(path)
        if df.empty:
            continue

        lower_map = {c.lower(): c for c in df.columns}
        have_keys = [k for k in want_lower if k in lower_map]
        if not have_keys:
            continue

        cols_real = [lower_map[k] for k in have_keys]
        sub = df.iloc[start:stop][cols_real].copy()

        rename_dict = {lower_map[k]: k for k in have_keys}
        sub = sub.rename(columns=rename_dict)

        for col in sub.columns:
            sub[col] = sub[col].map(strip_parens)

        parts.append(sub)

    if not parts:
        return pd.DataFrame(columns=["t1"])

    full_cols = sorted(set().union(*[set(p.columns) for p in parts]))
    parts = [p.reindex(columns=full_cols) for p in parts]
    return pd.concat(parts, ignore_index=True)

# ================== Runner ==================
class AblationRunner:
    def __init__(self, linked_csv: str, excel_dir: str, image_dir: str, out_xlsx: str):
        if not os.path.exists(linked_csv):
            raise FileNotFoundError(f"linked.csv nicht gefunden: {linked_csv}")
        self.linked = pd.read_csv(linked_csv)
        self.meta_df = load_meta_rows_from_excels(excel_dir)
        self.image_dir = image_dir
        self.out_xlsx = out_xlsx

    def _meta_for_obj(self, obj_id_raw: str) -> Dict[str, Any]:
        if self.meta_df.empty or "t1" not in self.meta_df.columns:
            return {"t1": obj_id_raw}
        left = str(obj_id_raw).strip()
        hit = self.meta_df[self.meta_df["t1"].astype(str).str.strip() == left]
        if hit.empty:
            return {"t1": obj_id_raw}
        return hit.iloc[0].to_dict()

    def _resolve_images(self, year: Any, images_field: Any) -> List[str]:
        image_names = _parse_images_field(images_field)
        paths: List[str] = []
        for img in image_names:
            if not img:
                continue
            # wenn schon ein Pfad drin ist:
            if os.path.isabs(img) and os.path.exists(img):
                paths.append(img)
                continue
            found = find_file_anywhere(self.image_dir, str(year), img)
            if found:
                paths.append(found)
        return paths

    def run(self, limit: int = 10, start_at: int = 0, show_prompts: bool = False):
        if self.linked.empty:
            print("WARN: linked.csv ist leer — nichts zu tun.")
            pd.DataFrame(columns=[
                "objekt_id","year","image_count_total","image_count_sent","prompt_text",
                "description_text_only","description_text_plus_images","errors"
            ]).to_excel(self.out_xlsx, index=False, engine="openpyxl")
            print("📘 gespeichert:", self.out_xlsx)
            return

        results: List[Dict[str, Any]] = []

        rows = self.linked.iloc[start_at:start_at+limit].copy()
        for idx, row in rows.iterrows():
            obj_id = str(row.get("obj_id_raw", "")).strip()
            year = row.get("year", "")

            meta_row = self._meta_for_obj(obj_id)
            prompt = build_prompt_from_tcols(meta_row)

            image_paths = self._resolve_images(year, row.get("images", ""))
            sent = image_paths[:MAX_IMAGES_PER_CALL]

            if show_prompts:
                print(f"\n--- PROMPT für {obj_id} ---\n{prompt}\n--- ENDE PROMPT ---\n")

            print(f"🔍 {obj_id}: images total={len(image_paths)} (send={len(sent)}) → text-only …")
            err = ""
            text_only = ""
            with_images = ""

            try:
                text_only = call_openai_text_only(prompt)
            except Exception as e:
                err += f"text_only_error: {e}; "
                text_only = f"ERROR: {e}"

            time.sleep(PAUSE_BETWEEN_OBJECTS)

            print(f"🔍 {obj_id}: images total={len(image_paths)} (send={len(sent)}) → text+images …")
            try:
                with_images = call_openai_with_images(image_paths, prompt)
            except Exception as e:
                err += f"with_images_error: {e}; "
                with_images = f"ERROR: {e}"

            results.append({
                "objekt_id": obj_id,
                "year": _norm_year(year),
                "image_count_total": len(image_paths),
                "image_count_sent": len(sent),
                "prompt_text": prompt,
                "description_text_only": text_only,
                "description_text_plus_images": with_images,
                "errors": err.strip()
            })

            time.sleep(PAUSE_BETWEEN_OBJECTS)

        df_out = pd.DataFrame(results)
        df_out.to_excel(self.out_xlsx, index=False, engine="openpyxl")
        print(f"\n📘 Ergebnisse gespeichert in: {self.out_xlsx}")


if __name__ == "__main__":
    runner = AblationRunner(
        linked_csv=LINKED_CSV,
        excel_dir=EXCEL_DIR,
        image_dir=IMAGE_DIR,
        out_xlsx=OUT_XLSX
    )

    # Für deinen Forschungsaufbau: erst klein testen (z.B. 10 Objekte)
    # Danach limit erhöhen.
    runner.run(limit=10, start_at=0, show_prompts=False)