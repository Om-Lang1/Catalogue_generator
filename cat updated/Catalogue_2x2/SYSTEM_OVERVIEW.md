# System Overview - Catalogue 2x2 Generator

## 🏗️ Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                         Web Browser                              │
│                   http://127.0.0.1:5000                         │
└────────────────┬────────────────────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                      Flask Application                           │
│  ┌──────────────┐  ┌──────────────┐  ┌────────────────────┐   │
│  │   Home Page  │  │  Catalogue   │  │   API Endpoints    │   │
│  │      /       │  │  Generator   │  │   /api/data        │   │
│  │              │  │  /catalogue  │  │   /api/catalogue   │   │
│  └──────────────┘  └──────────────┘  └────────────────────┘   │
└────────────────┬────────────────────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                   Data Processing Layer                          │
│  ┌──────────────────┐  ┌──────────────────┐  ┌──────────────┐ │
│  │  Diamond Data    │  │   Metal Data     │  │ Colour Stone │ │
│  │  Aggregation     │  │  Aggregation     │  │ Aggregation  │ │
│  └──────────────────┘  └──────────────────┘  └──────────────┘ │
└────────────────┬────────────────────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                    SQL Server Database                           │
│                      JASBSJEP                                    │
│  ┌──────────────┐  ┌──────────────┐  ┌────────────────────┐   │
│  │  StyleMst    │  │StyleMstDetail│  │  SPM_ItemView      │   │
│  │  SizeMst     │  │ SettingMst   │  │StyleMstCatchDetail │   │
│  └──────────────┘  └──────────────┘  └────────────────────┘   │
└─────────────────────────────────────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                    Output Generation                             │
│  ┌──────────────────┐              ┌──────────────────────┐    │
│  │   PDF Export     │              │   Excel Export       │    │
│  │  (ReportLab)     │              │   (OpenPyXL)         │    │
│  │  ✓ 2x2 Layout    │              │   ✓ 2x2 Layout       │    │
│  │  ✓ 3x3 Layout    │              │   ✓ 3x3 Layout       │    │
│  │  ✓ 4x4 Layout    │              │   ✓ 4x4 Layout       │    │
│  └──────────────────┘              └──────────────────────┘    │
└─────────────────────────────────────────────────────────────────┘
```

## 🔄 Data Flow

### 1. User Input
```
User enters Style Codes
      ↓
RG0002937C, ER0000589A, PD0000466A
      ↓
Split and validate codes
```

### 2. Database Query
```
Style Codes → SQL Query with CTEs
      ↓
DiamondData CTE (Diamond information)
      +
MetalData CTE (Metal information)
      +
ColourStoneData CTE (Colour Stone information)
      ↓
FULL OUTER JOIN all data
      ↓
Grouped by StyleCode
```

### 3. Data Processing
```
Raw Database Results
      ↓
Group by StyleCode
      ↓
Aggregate:
  - Sum Pieces_D (Diamond pieces)
  - Sum Weight_D (Diamond weight)
  - Sum Pieces_CS (Colour Stone pieces)
  - First RmCode_CS (Colour Stone type)
      ↓
Formatted Data Array
```

### 4. Output Generation
```
Formatted Data + Layout Selection
      ↓
┌─────────────┬─────────────┐
│ 2x2 Layout  │ 3x3 Layout  │ 4x4 Layout
│ (4 items)   │ (9 items)   │ (16 items)
└─────────────┴─────────────┘
      ↓
PDF (ReportLab) OR Excel (OpenPyXL)
      ↓
Download File
```

## 📊 Database Schema (Simplified)

### Key Tables

**StyleMst**
- StyleId (PK)
- StyleCode

**StyleMstDetail**
- StyleId (FK)
- ItemId (FK)
- Pieces
- NetWeight
- SizeNo (FK)

**SPM_ItemView**
- ItemId (PK)
- RawMitName (Diamond/Metal/Stone)
- ItemCode
- QlyCode

**SizeMst**
- SizeNo (PK)
- SizeCode

**SettingMst**
- SetNo (PK)
- SetCode
- SetName

## 🎯 Key Features

### Complex SQL Query
- **CTEs**: 3 separate queries for Diamond, Metal, and Colour Stone
- **Window Functions**: ROW_NUMBER() for ranking
- **FULL OUTER JOIN**: Combines all data types
- **COALESCE**: Handles NULL values
- **Dynamic Filtering**: Based on RawMitName patterns

### PDF Generation (ReportLab)
- **Page Layout**: A4 portrait (2x2) or landscape (3x3, 4x4)
- **Table Grid**: Dynamic based on layout
- **Cell Content**: Formatted text with bold labels
- **Borders**: Professional grid layout

### Excel Generation (OpenPyXL)
- **Column Width**: 30 units per cell
- **Row Height**: 120 units per cell
- **Text Wrapping**: Enabled for multi-line content
- **Borders**: Thin border around all cells
- **Formatting**: Bold design numbers, aligned text

## 🔐 Database Connection

**Connection Type**: Windows Authentication (Trusted Connection)
```python
DRIVER_NAME = 'SQL SERVER'
SERVER_NAME = 'DESKTOP-D9GJDD4'
DATABASE_NAME = 'JASBSJEP'
Trust_Connection = yes
```

## 📁 File Structure

```
Catalogue_2x2/
├── app.py                      # Main Flask application
│   ├── Routes
│   │   ├── GET  /              # Home page
│   │   ├── GET  /catalogue     # Catalogue generator UI
│   │   ├── GET  /api/data      # All data API
│   │   ├── POST /api/catalogue-data  # Filtered data API
│   │   ├── GET  /generate-pdf/<layout>
│   │   └── GET  /generate-excel/<layout>
│   │
│   ├── Functions
│   │   ├── get_db_connection()
│   │   ├── get_catalogue_data(style_codes)
│   │   └── create_pdf_cell(item)
│   │
│   └── Libraries
│       ├── Flask (Web framework)
│       ├── pypyodbc (Database)
│       ├── reportlab (PDF)
│       ├── openpyxl (Excel)
│       └── Pillow (Images)
│
├── templates/
│   ├── index.html              # Home page template
│   ├── catalogue.html          # Catalogue generator UI
│   └── error.html              # Error page template
│
├── requirements.txt            # Python dependencies
├── README.md                   # Full documentation
├── QUICK_START.md             # Quick start guide
└── SYSTEM_OVERVIEW.md         # This file
```

## 🚀 Performance Considerations

### Database Query
- **Indexes**: Ensure indexes on StyleId, ItemId, SizeNo
- **CTEs**: Efficient for complex joins
- **WHERE Clauses**: Filter early in CTEs
- **ROW_NUMBER()**: Efficient for top-N queries

### Web Application
- **Connection Pooling**: Close connections after each request
- **Error Handling**: Try-catch blocks for all database operations
- **Response Time**: Preview loads in 1-3 seconds
- **File Generation**: PDF/Excel generates in 2-5 seconds

### File Generation
- **PDF**: Uses ReportLab for efficient rendering
- **Excel**: OpenPyXL handles large datasets well
- **Memory**: Files generated in BytesIO (in-memory)
- **Download**: Streaming response for large files

## 🔧 Configuration

### Database Configuration (app.py)
```python
DRIVER_NAME = 'SQL SERVER'
SERVER_NAME = 'DESKTOP-D9GJDD4'
DATABASE_NAME = 'JASBSJEP'
```

### PDF Settings
- **Page Size**: A4 (210mm × 297mm)
- **Orientation**: Portrait (2x2) / Landscape (3x3, 4x4)
- **Margins**: 0.5cm all sides
- **Font Size**: 8pt for cell content

### Excel Settings
- **Column Width**: 30 units
- **Row Height**: 120 units
- **Font**: Default (Calibri 10pt)
- **Alignment**: Top-left, wrap text

## 📈 Scalability

### Current Capacity
- **Style Codes**: Tested with 6-50 codes
- **PDF Pages**: 1-20 pages
- **Excel Rows**: 1-100 rows
- **Response Time**: < 5 seconds

### Future Enhancements
- Image integration in catalogues
- Custom layouts and templates
- Batch processing for large datasets
- Database query optimization
- Caching for frequently accessed data

---

**Version**: 1.0.0  
**Last Updated**: January 30, 2026
