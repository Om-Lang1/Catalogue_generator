# Quick Start Guide - Catalogue Generator

## 🚀 Quick Setup

The application is now running at: **http://127.0.0.1:5000**

## 📋 How to Generate Catalogues

### Step 1: Access the Catalogue Generator
1. Open your browser and go to: `http://127.0.0.1:5000`
2. Click the **"📋 Go to Catalogue Generator"** button

### Step 2: Enter Style Codes
Enter your product style codes in the text area. You can:
- Enter one code per line
- Separate codes with commas
- Mix both methods

**Example:**
```
RG0002937C
ER0000589A
PD0000466A
```
or
```
RG0002937C, ER0000589A, PD0000466A, RG0002562B, ER0000496B, PD0000372B
```

### Step 3: Preview Data
Click the **"🔍 Load Preview"** button to:
- Verify your style codes are correct
- See the data that will be included
- Check Diamond and Colour Stone details

### Step 4: Select Layout & Generate
Choose your preferred layout and format:

#### **2x2 Layout** (4 items per page)
- Best for: Detailed view with large product display
- Click **"📄 PDF"** for PDF format
- Click **"📊 Excel"** for Excel format

#### **3x3 Layout** (9 items per page)
- Best for: Balanced layout with good visibility
- Click **"📄 PDF"** for PDF format
- Click **"📊 Excel"** for Excel format

#### **4x4 Layout** (16 items per page)
- Best for: Compact view with many products
- Click **"📄 PDF"** for PDF format
- Click **"📊 Excel"** for Excel format

## 📊 What Data is Displayed?

Each product card shows:

| Field | Description | Example |
|-------|-------------|---------|
| **Design No** | Style Code | ER0000589A |
| **Dia Details** | Diamond quantity and weight | 24 Qty / 0.12 Cts |
| **Col Qty** | Colour stone type and quantity | Emerald - 2 |

## 🎯 Tips & Best Practices

1. **Preview First**: Always click "Load Preview" to verify data before generating
2. **Multiple Codes**: You can process as many style codes as needed
3. **Layout Selection**: 
   - Use 2x2 for catalogs with fewer items (more detail)
   - Use 3x3 for medium-sized catalogs (balanced)
   - Use 4x4 for large catalogs (more items per page)
4. **File Downloads**: Files are automatically named with timestamp for easy organization

## 📁 Generated File Naming

Files are automatically named:
- **PDF**: `catalogue_2x2_20260130_123045.pdf`
- **Excel**: `catalogue_3x3_20260130_123045.xlsx`

Format: `catalogue_<layout>_<YYYYMMDD>_<HHMMSS>.<extension>`

## ⚠️ Troubleshooting

### "No data found" error
- Check if the style codes exist in the database
- Verify the style codes are spelled correctly
- Ensure the database connection is working

### "Database connection failed" error
- Verify SQL Server is running
- Check database credentials in `app.py`
- Ensure Windows authentication is enabled

### PDF/Excel not downloading
- Check browser's download settings
- Ensure pop-up blocker is disabled
- Try a different browser

## 🔧 Advanced Usage

### API Endpoint
You can also use the API directly:

```bash
POST /api/catalogue-data
Content-Type: application/json

{
  "style_codes": ["RG0002937C", "ER0000589A", "PD0000466A"]
}
```

### Direct PDF Generation
```
GET /generate-pdf/2x2?style_codes=RG0002937C,ER0000589A,PD0000466A
```

### Direct Excel Generation
```
GET /generate-excel/3x3?style_codes=RG0002937C,ER0000589A,PD0000466A
```

## 📞 Need Help?

- Check the main `README.md` for detailed documentation
- Review the sample data to understand the format
- Verify your database connection settings

---

**Happy Cataloguing! 📋✨**
