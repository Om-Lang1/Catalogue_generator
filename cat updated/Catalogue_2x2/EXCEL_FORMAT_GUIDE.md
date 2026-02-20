# Excel Catalogue Format Guide

## Overview
The Excel catalogue now matches the exact format shown in the PDF layout with a structured table design.

## Layout Structure

### Grid Layout
- **2x2 Layout**: 4 items per page (2 rows × 2 columns)
- **3x3 Layout**: 9 items per page (3 rows × 3 columns)
- **4x4 Layout**: 16 items per page (4 rows × 4 columns)

### Item Structure
Each item is displayed as a 2-column table with 6 rows:

| Label Column | Value Column |
|--------------|--------------|
| **Image** (merged) | Product image (120px height) |
| **Design No :** | Style Code (e.g., RG0002937C) |
| **Grams :** | Metal Weight (e.g., 1.400) |
| **Dia Details :** | Diamond Qty/Carats (e.g., 26 Qty / 0.2 Cts) |
| **Col Qty :** | Stone Name - Quantity (e.g., Emerald - 1) |
| **Tray :** | Set Code - Set Name (e.g., TCS/1 - 1/1) |

## Field Descriptions

1. **Design No**: The unique style code from the database
2. **Grams**: Metal weight (from Weight_M field)
3. **Dia Details**: Diamond pieces and total carat weight
4. **Col Qty**: Color stone name and piece count
5. **Tray**: Tray/Set code and name for organization

## Styling Features

- **Label Column**: Bold text, left-aligned, 15px width
- **Value Column**: Regular text, center-aligned, 20px width, blue color for Design No
- **Image Row**: Light blue background (#E8F4F8), centered, merged across both columns
- **Borders**: Thin borders around all cells for clear separation
- **Row Heights**: 
  - Image row: 120px
  - Data rows: 25px each

## Data Mapping

The Excel generation pulls data from the database query that includes:
- Diamond data (Pieces_D, Weight_D)
- Metal data (Weight_M for Grams)
- Color Stone data (Pieces_CS, QltyCode_CS converted to stone names)
- Set data (SetCode, SetName for Tray)
- Images from network path (\\\\sjserver\\GatiSoftTech\\Images\\...)

## Stone Name Mapping

Quality codes are automatically converted to friendly names:
- EM → Emerald
- RU → Ruby
- SP → Sapphire
- AM → Amethyst
- And more...

## Usage

1. Navigate to `/catalogue` page
2. Enter style codes (comma or line-separated)
3. Click "Load Preview" to verify data
4. Select layout (2x2, 3x3, or 4x4)
5. Click "Excel" button to download

The generated Excel file will have the exact format matching the PDF catalogue layout.
