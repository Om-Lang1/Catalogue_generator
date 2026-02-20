# Colour Stone Name Mapping

## 📊 Data Source Explanation

### Where the Data Comes From:

**Col Qty Format:** `{Stone Name} - {Pieces}`

Example: `Emerald - 1`, `Ruby - 2`

### Database Columns Used:

1. **`QltyCode_CS`** (Quality Code for Colour Stones) 
   - Source: `SPM_ItemView.QlyCode`
   - Contains short codes like: `EM`, `RU`, `SP`, `AM`, etc.

2. **`Pieces_CS`** (Pieces Count)
   - Source: `StyleMstDetail.Pieces`
   - Contains the number of colour stone pieces: `1`, `2`, `3`, etc.

## 🗺️ Quality Code to Stone Name Mapping

The system automatically converts quality codes to full stone names:

| Quality Code | Stone Name   | Example Items                          |
|--------------|--------------|----------------------------------------|
| **EM**       | Emerald      | SYN-RND-GR-EM (Synthetic Green Emerald)|
| **RU**       | Ruby         | SYN-OV-RED-RU (Synthetic Red Ruby)     |
| **SP**       | Sapphire     | PS-EMR-BL-SP (Blue Sapphire)           |
| **AM**       | Amethyst     | Semi Precious Amethyst                 |
| **AQM**      | Aquamarine   | Semi Precious Aquamarine               |
| **PD**       | Peridot      | CG-PER-GR-PD (Green Peridot)           |
| **TZ**       | Tanzanite    | CG-EMR-BL-TZ (Blue Tanzanite)          |
| **TL**       | Tourmaline   | CG-EMR-PNK-TL (Pink Tourmaline)        |
| **TP**       | Topaz        | CG-EMR-SKBL-TP (Sky Blue Topaz)        |
| **GR**       | Garnet       | CG-OVL-RED-GR (Red Garnet)             |
| **OPL**      | Opal         | Opal Stones                            |
| **MC**       | Malachite    | CG-RCT-GR-MC (Green Malachite)         |
| **CIT**      | Citrine      | Citrine Stones                         |
| **QTZ**      | Quartz       | Quartz Stones                          |
| **OX**       | Onyx         | CG-CUS-BLK-OX (Black Onyx)             |
| **MG**       | Morganite    | Semi Precious Morganite                |

## 📋 How It Works

### 1. SQL Query Level
```sql
ColourStoneData AS (
    SELECT
        d.QlyCode AS QltyCode_CS,  -- Quality code (EM, RU, SP, etc.)
        c.Pieces AS Pieces_CS,      -- Number of pieces
        ...
    FROM StyleMst a
    INNER JOIN StyleMstDetail c ON a.StyleId = c.StyleId
    LEFT JOIN SPM_ItemView d ON c.ItemId = d.ItemId
    WHERE d.RawMitName LIKE '%Stone%'
)
```

### 2. Python Processing Level
```python
# Extract quality code
qlty_code_cs = 'EM'  # or 'EM-AA', 'RU-AAA', etc.

# Map to stone name
stone_name_map = {
    'EM': 'Emerald',
    'RU': 'Ruby',
    'SP': 'Sapphire',
    ...
}

# Handle compound codes (e.g., "EM-AA" -> "EM")
base_code = qlty_code_cs.split('-')[0]  # 'EM'
stone_name = stone_name_map.get(base_code, qlty_code_cs)  # 'Emerald'
```

### 3. Display Level
```
Col Qty: Emerald - 1
         └─┬─┘    └─┬─┘
           │         │
    Stone Name   Pieces Count
   (from QltyCode) (from Pieces_CS)
```

## 🔍 Quality Code Variations

Some quality codes have suffixes indicating quality grades:

- **EM-A** → Emerald (Grade A)
- **EM-AA** → Emerald (Grade AA)
- **EM-AAA** → Emerald (Grade AAA)
- **RU-AA** → Ruby (Grade AA)
- **SP-AAA** → Sapphire (Grade AAA)

The system extracts the base code (e.g., `EM` from `EM-AA`) and maps it to the stone name.

## 📊 Real Data Examples

### From Your StyleCodes:

**ER0000589A:**
- QltyCode: `EM`
- Pieces: 2 + 12 = 14
- Display: **"Emerald - 14"**

**ER0000496B:**
- QltyCode: `RU`
- Pieces: 2
- Display: **"Ruby - 2"**

**RG0002937C:**
- QltyCode: `EM`
- Pieces: 1 + 6 = 7
- Display: **"Emerald - 7"**

**PD0000372B:**
- QltyCode: `RU`
- Pieces: 1
- Display: **"Ruby - 1"**

## 🔄 Data Flow Summary

```
Database (SPM_ItemView)
    ↓
QlyCode = "EM"
    ↓
Split by '-' if compound code
    ↓
Base Code = "EM"
    ↓
Map to stone name
    ↓
Stone Name = "Emerald"
    ↓
Combine with Pieces_CS
    ↓
Output: "Emerald - 1"
```

## ✨ Benefits

1. **User-Friendly**: Displays full stone names instead of cryptic codes
2. **Flexible**: Handles quality grade suffixes automatically
3. **Consistent**: Same mapping across web UI, PDF, and Excel
4. **Extensible**: Easy to add new stone types to the mapping

## 🔧 Adding New Stone Types

To add a new stone type, simply update the `stone_name_map` dictionary in `app.py`:

```python
stone_name_map = {
    'EM': 'Emerald',
    'RU': 'Ruby',
    'SP': 'Sapphire',
    'NEW': 'New Stone Name',  # Add your new mapping here
    ...
}
```

---

**Updated:** January 30, 2026  
**Version:** 1.1.0
