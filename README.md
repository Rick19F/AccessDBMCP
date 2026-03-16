# Access Database MCP Server

A standalone MCP (Model Context Protocol) server for interacting with Microsoft Access databases.

## Overview

This MCP server provides tools to connect to and query Microsoft Access databases (.accdb and .mdb files) using pyodbc. It is designed to be used as a standalone service, completely separate from the main AMS-MIS-EDESAL codebase.

## Requirements

- Python 3.8+
- Microsoft Access ODBC Driver (must be installed on the system)
- The dependencies listed in `requirements.txt`

## Installation

1. **Create and activate virtual environment:**

```bash
cd access-mcp
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

2. **Install dependencies:**

```bash
pip install -r requirements.txt
```

3. **Install Microsoft Access ODBC Driver:**

On Windows, the driver is typically included with Microsoft Office.  
On Linux, you may need to use MDBTools or a Windows VM.

## Usage

### Running the MCP Server

```bash
# Activate virtual environment first
source venv/bin/activate

# Run the MCP server
python access_mcp.py
```

The server will start and listen for MCP requests via stdio.

### Available Tools

| Tool                         | Description                               |
| ---------------------------- | ----------------------------------------- |
| `access_connect`             | Connect to an Access database file        |
| `access_disconnect`          | Disconnect from the current database      |
| `access_get_tables`          | List all tables in the database           |
| `access_get_table_structure` | Get column information for a table        |
| `access_execute_query`       | Execute a SQL query                       |
| `access_get_table_data`      | Get data from a table with optional limit |
| `access_connection_status`   | Check current connection status           |

### Example Usage

#### Connect to a database

```json
{
  "name": "access_connect",
  "arguments": {
    "db_path": "/path/to/database.accdb",
    "password": "optional_password"
  }
}
```

#### List tables

```json
{
  "name": "access_get_tables",
  "arguments": {}
}
```

#### Get table structure

```json
{
  "name": "access_get_table_structure",
  "arguments": {
    "table_name": "Customers"
  }
}
```

#### Execute a query

```json
{
  "name": "access_execute_query",
  "arguments": {
    "query": "SELECT * FROM Customers WHERE City = ?",
    "params": ["San Salvador"]
  }
}
```

#### Get table data

```json
{
  "name": "access_get_table_data",
  "arguments": {
    "table_name": "Customers",
    "limit": 100,
    "offset": 0
  }
}
```

## Project Structure

```
access-mcp/
├── venv/              # Python virtual environment
├── requirements.txt   # Python dependencies
├── access_mcp.py      # Main MCP server file
└── README.md          # This file
```

## Dependencies

- `mcp` - Model Context Protocol server SDK
- `pyodbc` - Python ODBC database connector
- `pandas` - Data manipulation and analysis (optional, for advanced data processing)

## Notes

- This MCP server uses stdio for communication, making it compatible with various MCP clients
- The server automatically filters out Microsoft Access system tables (MSys\*)
- Connection strings use the Microsoft Access Driver (_.mdb, _.accdb)
- The driver must be installed on the system for the connection to work

## Troubleshooting

### "Driver not found" error

Make sure the Microsoft Access ODBC Driver is installed on your system. On Windows, you can check installed drivers in ODBC Data Source Administrator.

### "Database file not found" error

Ensure the path to the Access database file is correct and the file exists.

### Connection issues

- Verify the database file is not corrupted
- Check file permissions
- Ensure the database is not opened exclusively by another application
