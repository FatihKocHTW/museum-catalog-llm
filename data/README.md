# Data

The inventory lists and object photos come from a museum collection and are **not part of this repository**.
To run the pipeline, place them here:

```
data/
├── Objektdaten/
│   ├── Liste_AK Kommunikation.xls
│   └── Liste_AEG Produktsammlung.xls
└── Objektbilder/
    ├── 1997/
    │   ├── 1-1997-1063-000-000.JPG
    │   └── ...
    └── 2024/
        └── ...
```

The Excel exports use the raw column names `t1` … `t14`; `t1` holds the inventory number
in the format `1/1997/1063 0`. The row ranges that are processed are set in
`src/museum_catalog/config.py` (`EXCEL_SPECS`).
