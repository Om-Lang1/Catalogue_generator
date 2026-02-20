# Database Connection Checker

A simple Flask application to test and monitor the database connection status.

## Features

- **Real-time Status**: Shows whether the database is connected or disconnected
- **Connection Details**: Displays server name, database name, and driver information
- **Auto-refresh**: Automatically refreshes status every 30 seconds
- **Manual Refresh**: Button to manually refresh the connection status
- **Visual Indicators**: Color-coded status (green for connected, red for disconnected)

## How to Run

### Method 1: Using the Batch File (Recommended)
1. Double-click on `run_db_checker.bat`
2. Wait for the server to start
3. Open your web browser and go to `http://localhost:5001`

### Method 2: Using Python Directly
1. Open Command Prompt/PowerShell
2. Navigate to the project directory
3. Run: `python db_connect.py`
4. Open your web browser and go to `http://localhost:5001`

## Database Configuration

The checker uses the same database configuration as your main application:

- **Server**: SJSERVER
- **Database**: TSHISJEP
- **Driver**: SQL SERVER
- **Authentication**: Windows Authentication (Trust_Connection=yes)

## What It Checks

1. **Connection Test**: Attempts to establish a connection to the database
2. **Authentication**: Verifies Windows Authentication works
3. **Database Availability**: Confirms the database is accessible
4. **Driver Compatibility**: Ensures the SQL Server driver is working

## Status Indicators

- ✅ **Connected** (Green): Database connection is successful
- ❌ **Disconnected** (Red): Database connection failed with error details

## Troubleshooting

If the connection fails, check:

1. **Server Status**: Ensure SQL Server is running on SJSERVER
2. **Network**: Verify network connectivity to the server
3. **Authentication**: Ensure your Windows account has database access
4. **Firewall**: Check if firewall is blocking the connection
5. **Driver**: Verify SQL Server ODBC driver is installed

## Auto-refresh

The page automatically refreshes every 30 seconds to provide up-to-date status information. You can also manually refresh using the "Refresh Status" button.

## Stopping the Server

To stop the database checker:
- Press `Ctrl+C` in the command window
- Or close the command window

---

**Created**: February 4, 2026  
**Purpose**: Database connection monitoring for Catalogue System
