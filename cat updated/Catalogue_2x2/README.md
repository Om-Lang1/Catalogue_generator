# Catalogue 2x2 - Product Catalogue Generator

A Flask web application that connects to a SQL Server database and generates beautiful product catalogues in PDF and Excel formats with multiple layout options (2x2, 3x3, 4x4).

## Features

- 📊 Database connection to SQL Server with complex query support
- 🎨 Modern gradient UI design
- 📱 Responsive table display
- 📋 **Catalogue Generator** - Generate catalogues in multiple formats
- 📄 **PDF Export** - Create PDF catalogues in 2x2, 3x3, or 4x4 layouts
- 📊 **Excel Export** - Create Excel catalogues in 2x2, 3x3, or 4x4 layouts
- 🔍 **Data Preview** - Preview product data before generating catalogues
- 🔌 REST API endpoints for JSON data
- ⚡ Real-time data fetching
- 💎 Support for Diamond, Metal, and Colour Stone data

## Requirements

- Python 3.x
- SQL Server with database named `JASBSJEP`
- Server: `DESKTOP-D9GJDD4`
- Windows authentication enabled
- Required Python packages (see requirements.txt):
  - Flask
  - pypyodbc
  - reportlab (for PDF generation)
  - openpyxl (for Excel generation)
  - Pillow (for image processing)

## Installation

1. Install dependencies:
```bash
pip install -r requirements.txt
```

2. Ensure SQL Server is running and accessible with the configured credentials.

## Running the Application

1. Start the Flask server:
```bash
python app.py
```

2. Open your browser and navigate to:
```
http://127.0.0.1:5000
```

## Usage

### Catalogue Generator

1. Navigate to the **Catalogue Generator** page from the home page
2. Enter Style Codes (one per line or comma-separated) in the text area
   - Example: `RG0002937C, ER0000589A, PD0000466A`
3. Click **Load Preview** to see the data
4. Select your desired layout:
   - **2x2** - 4 items per page (ideal for detailed view)
   - **3x3** - 9 items per page (balanced layout)
   - **4x4** - 16 items per page (compact view)
5. Click **PDF** or **Excel** button to generate and download the catalogue

### Catalogue Data Format

Each catalogue item displays:
- **Design No**: Style code (e.g., ER0000589A)
- **Dia Details**: Diamond quantity and weight (e.g., 24 Qty / 0.12 Cts)
- **Col Qty**: Colour stone type and quantity (e.g., Emerald - 2)

## API Endpoints

- `GET /` - Main page displaying database records in a table
- `GET /catalogue` - Catalogue generator interface
- `GET /api/data` - JSON API endpoint for all data
- `POST /api/catalogue-data` - Fetch catalogue data for specific style codes
- `GET /generate-pdf/<layout>?style_codes=<codes>` - Generate PDF catalogue
- `GET /generate-excel/<layout>?style_codes=<codes>` - Generate Excel catalogue

## Database Configuration

The application connects to:
- **Server**: DESKTOP-D9GJDD4
- **Database**: JASBSJEP
- **Tables**: 
  - StyleMst
  - StyleMstCatchDetailView
  - StyleMstDetail
  - SPM_ItemView
  - SizeMst
  - SettingMst

The application uses complex SQL queries with CTEs (Common Table Expressions) to join Diamond, Metal, and Colour Stone data. To modify the query, edit the SQL statements in `app.py`.

## Project Structure

```
Catalogue_2x2/
├── app.py                 # Main Flask application with catalogue generation
├── requirements.txt       # Python dependencies
├── templates/
│   ├── index.html        # Main page template with data viewer
│   ├── catalogue.html    # Catalogue generator interface
│   └── error.html        # Error page template
└── README.md             # This file
```

## Development

The application runs in debug mode by default. Changes to the code will automatically reload the server.

To stop the server, press `CTRL+C` in the terminal.
