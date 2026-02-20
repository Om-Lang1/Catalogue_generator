# Excel Catalogue Visual Layout Guide

## Complete Excel Structure

### 2x2 Layout Example (4 items per page)

```
┌────────────────────────┬────┬────────────────────────┐
│    ITEM 1 (2 cols)     │    │    ITEM 2 (2 cols)     │
├───────────┬────────────┼────┼───────────┬────────────┤
│   Image (merged)       │    │   Image (merged)       │
├───────────┼────────────┼────┼───────────┼────────────┤
│Design No: │ RG0002937C │    │Design No: │ ER0000589A │
├───────────┼────────────┼────┼───────────┼────────────┤
│Grams :    │ 1.400      │    │Grams :    │ 1.596      │
├───────────┼────────────┼────┼───────────┼────────────┤
│Dia Det.:  │26 Qty/0.2C │    │Dia Det.:  │24 Qty/0.12 │
├───────────┼────────────┼────┼───────────┼────────────┤
│Col Qty :  │Emerald - 1 │    │Col Qty :  │Emerald - 2 │
├───────────┼────────────┼────┼───────────┼────────────┤
│Tray :     │TCS/1 - 1/1 │    │Tray :     │TCS/1 - 1/2 │
└───────────┴────────────┴────┴───────────┴────────────┘
       ↑ Blank column (2px) ↑
                                                         
← Blank row (15px) for vertical spacing between item rows
                                                         
┌────────────────────────┬────┬────────────────────────┐
│    ITEM 3 (2 cols)     │    │    ITEM 4 (2 cols)     │
├───────────┬────────────┼────┼───────────┬────────────┤
│   Image (merged)       │    │   Image (merged)       │
├───────────┼────────────┼────┼───────────┼────────────┤
│Design No: │ PD0000466A │    │Design No: │ RG0002562B │
├───────────┼────────────┼────┼───────────┼────────────┤
│Grams :    │ 0.912      │    │Grams :    │ 1.420      │
├───────────┼────────────┼────┼───────────┼────────────┤
│Dia Det.:  │12 Qty/0.09 │    │Dia Det.:  │28 Qty/0.22 │
├───────────┼────────────┼────┼───────────┼────────────┤
│Col Qty :  │Emerald - 1 │    │Col Qty :  │Ruby - 1    │
├───────────┼────────────┼────┼───────────┼────────────┤
│Tray :     │TCS/1 - 1/3 │    │Tray :     │TCS/2 - 2/1 │
└───────────┴────────────┴────┴───────────┴────────────┘
```

**Notes:** 
- Blank columns (2px wide) separate items horizontally
- Blank rows (15px high) separate item rows vertically

## Column & Row Specifications

### 2x2 Layout
- **Total Columns**: 5 (2 items × 2 columns each + 1 blank column between)
- **Column Pattern**: Label, Value, Blank, Label, Value
- **Rows per item**: 6 data rows + 1 blank row = 7 rows total
- **Total rows for page**: 14 rows (2 item rows × 7 rows each)

### 3x3 Layout
- **Total Columns**: 8 (3 items × 2 columns each + 2 blank columns between)
- **Column Pattern**: Label, Value, Blank, Label, Value, Blank, Label, Value
- **Rows per item**: 6 data rows + 1 blank row = 7 rows total
- **Total rows for page**: 21 rows (3 item rows × 7 rows each)

### 4x4 Layout
- **Total Columns**: 11 (4 items × 2 columns each + 3 blank columns between)
- **Column Pattern**: Label, Value, Blank, Label, Value, Blank, Label, Value, Blank, Label, Value
- **Rows per item**: 6 data rows + 1 blank row = 7 rows total
- **Total rows for page**: 28 rows (4 item rows × 7 rows each)

## Cell Specifications

### Column A, D, G, J... (Label Columns)
- **Width**: 15 characters
- **Font**: Bold, Size 10
- **Alignment**: Left, Vertical Center
- **Border**: Thin on all sides

### Column B, E, H, K... (Value Columns)
- **Width**: 20 characters
- **Font**: Regular, Size 10 (Blue for Design No)
- **Alignment**: Center, Vertical Center
- **Border**: Thin on all sides

### Column C, F, I, L... (Blank Separator Columns)
- **Width**: 2 characters
- **Purpose**: Visual separation between items
- **No content or borders**

### Row 1, 8, 15, 22... (Image Rows)
- **Height**: 120 pixels
- **Background**: Light blue (#E8F4F8)
- **Content**: "Image" text or actual product image
- **Merge**: Across 2 columns (label + value)

### Rows 2-6, 9-13, 16-20... (Data Rows)
- **Height**: 25 pixels
- **Background**: White
- **Content**: Label | Value pairs

### Rows 7, 14, 21, 28... (Blank Separator Rows)
- **Height**: 15 pixels
- **Purpose**: Vertical spacing between item rows
- **No content or borders**

## Data Row Details

### Row Pattern for Each Item (7 rows total)
1. **Row 1**: Image (merged, 120px height, light blue bg)
2. **Row 2**: Design No | StyleCode (blue value)
3. **Row 3**: Grams | Weight_M (3 decimal places)
4. **Row 4**: Dia Details | "X Qty / Y.YY Cts"
5. **Row 5**: Col Qty | "StoneName - X"
6. **Row 6**: Tray | "SetCode - SetName"
7. **Row 7**: Blank row (15px height, for spacing)

## Example Data Mapping

### From Database to Excel

**Database Row:**
```
StyleCode: RG0002937C
Weight_M: 1.4
Pieces_D: 26
Weight_D: 0.2
QltyCode_CS: EM-AA
Pieces_CS: 1
SetCode: TCS/1
SetName: 1/1
ImagePath: \\sjserver\...\RG\DM 3D RG0002937C.jpg
```

**Excel Output:**
```
┌───────────┬────────────┐
│  [Image]  │  [Image]   │  ← Row 1 (merged, 120px)
├───────────┼────────────┤
│Design No: │ RG0002937C │  ← Row 2
├───────────┼────────────┤
│Grams :    │ 1.400      │  ← Row 3
├───────────┼────────────┤
│Dia Det.:  │26 Qty/0.2C │  ← Row 4
├───────────┼────────────┤
│Col Qty :  │Emerald - 1 │  ← Row 5 (EM-AA → Emerald)
├───────────┼────────────┤
│Tray :     │TCS/1 - 1/1 │  ← Row 6
└───────────┴────────────┘
```

## Quality Code Mappings

The system automatically converts quality codes to friendly stone names:

| Code | Stone Name  | Code | Stone Name  |
|------|------------|------|-------------|
| EM   | Emerald    | TZ   | Tanzanite   |
| RU   | Ruby       | TL   | Tourmaline  |
| SP   | Sapphire   | TP   | Topaz       |
| AM   | Amethyst   | GR   | Garnet      |
| AQM  | Aquamarine | OPL  | Opal        |
| PD   | Peridot    | MC   | Malachite   |
| CIT  | Citrine    | OX   | Onyx        |
| QTZ  | Quartz     | MG   | Morganite   |

## Grid Positioning Logic

### For 2x2 Layout:
- **Item 1**: Cols 1-2, Rows 1-6, [Blank Col 3], [Blank Row 7]
- **Item 2**: Cols 4-5, Rows 1-6, [Blank Row 7]
- **Item 3**: Cols 1-2, Rows 8-13, [Blank Col 3], [Blank Row 14]
- **Item 4**: Cols 4-5, Rows 8-13, [Blank Row 14]

### For 3x3 Layout:
- **Row 1**: 
  - Item 1: Cols 1-2, Rows 1-6, [Blank Col 3]
  - Item 2: Cols 4-5, Rows 1-6, [Blank Col 6]
  - Item 3: Cols 7-8, Rows 1-6
  - [Blank Row 7]
- **Row 2**: Same column pattern, Rows 8-13, [Blank Row 14]
- **Row 3**: Same column pattern, Rows 15-20, [Blank Row 21]

### For 4x4 Layout:
- **Row 1**: 
  - Item 1: Cols 1-2, [Blank Col 3]
  - Item 2: Cols 4-5, [Blank Col 6]
  - Item 3: Cols 7-8, [Blank Col 9]
  - Item 4: Cols 10-11
  - [Blank Row 7]
- **Rows 2-4**: Same column pattern for subsequent rows with blank rows after each

## Implementation Code Reference

The Excel generation follows this algorithm:

```python
# Each item takes 2 columns + 1 blank column (except last)
# Each item takes 6 data rows + 1 blank row
item_width_cols = 2
blank_col_width = 1
item_height_rows = 6
blank_row_height = 1

# For each item in the grid
for grid_row_idx in range(grid_size):
    for grid_col_idx in range(grid_size):
        item_idx = grid_row_idx * grid_size + grid_col_idx
        
        # Calculate position with blank columns
        col_start = grid_col_idx * (item_width_cols + blank_col_width) + 1
        row_start = current_row
        
        # Create 6 data rows + 1 blank row
        # Rows 1-6: Item data
        # Row 7: Blank for spacing
        
    # Move to next row including blank row
    current_row += item_height_rows + blank_row_height
```

## Testing the Output

1. Navigate to http://127.0.0.1:5000/catalogue
2. Enter test style codes:
   ```
   RG0002937C, ER0000589A, PD0000466A
   ```
3. Click "Load Preview" to verify data
4. Select 2x2 layout
5. Click "Excel" button
6. Open downloaded file to verify:
   - ✅ 5 columns for 2x2 (2 items × 2 cols + 1 blank)
   - ✅ 14 rows (2 item rows × 7 rows each, includes blank rows)
   - ✅ Blank columns between items (2px width)
   - ✅ Blank rows after each item row (15px height)
   - ✅ Images in rows 1 and 8
   - ✅ All 5 data fields present
   - ✅ Proper borders and styling
   - ✅ Stone names (not codes)

## File Output Example

**Filename Format:**
```
catalogue_2x2_20260131_123456.xlsx
```

Where:
- `2x2`: Layout selected
- `20260131`: Date (YYYYMMDD)
- `123456`: Time (HHMMSS)

This ensures unique filenames for each generation.
