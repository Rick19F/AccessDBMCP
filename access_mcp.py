#!/usr/bin/env python3
"""
Standalone MCP Server for Microsoft Access Databases

This MCP server provides tools to interact with Microsoft Access databases
using pyodbc. It can list tables, get table structures, execute queries,
and retrieve data.

Usage:
    python access_mcp.py

Requirements:
    - mcp
    - pyodbc
    - pandas
"""

import os
import pyodbc
import pandas as pd
from typing import Any, Dict, List, Optional, Union
from mcp.server import Server
from mcp.server.stdio import stdio_server
from mcp.types import Tool, TextContent
from mcp.server.handlers import ListToolsRequestHandler, CallToolRequestHandler


class AccessDatabase:
    """Handles connections and operations for Microsoft Access databases."""
    
    def __init__(self):
        self.connection: Optional[pyodbc.Connection] = None
        self.current_db_path: Optional[str] = None
    
    def connect(self, db_path: str, password: Optional[str] = None) -> Dict[str, Any]:
        """
        Connect to an Access database.
        
        Args:
            db_path: Path to the .accdb or .mdb file
            password: Optional database password
            
        Returns:
            Dictionary with connection status
        """
        try:
            if not os.path.exists(db_path):
                return {"success": False, "error": f"Database file not found: {db_path}"}
            
            # Build connection string for Access
            if password:
                conn_str = (
                    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
                    f"DBQ={db_path};"
                    f"PWD={password}"
                )
            else:
                conn_str = (
                    r"DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};"
                    f"DBQ={db_path}"
                )
            
            self.connection = pyodbc.connect(conn_str, autocommit=True)
            self.current_db_path = db_path
            
            return {
                "success": True,
                "message": f"Connected to {os.path.basename(db_path)}",
                "db_path": db_path
            }
        except pyodbc.Error as e:
            return {"success": False, "error": str(e)}
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    def disconnect(self) -> Dict[str, Any]:
        """Close the database connection."""
        try:
            if self.connection:
                self.connection.close()
                self.connection = None
                self.current_db_path = None
                return {"success": True, "message": "Disconnected from database"}
            return {"success": True, "message": "No active connection"}
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    def get_tables(self) -> Dict[str, Any]:
        """Get list of all tables in the database."""
        try:
            if not self.connection:
                return {"success": False, "error": "Not connected to database"}
            
            cursor = self.connection.cursor()
            tables = []
            
            # Get all tables (Access system tables start with MSys)
            for row in cursor.tables():
                table_name = row.table_name
                table_type = row.table_type
                # Filter out system tables
                if table_type == 'TABLE' and not table_name.startswith('MSys'):
                    tables.append({
                        "name": table_name,
                        "type": table_type
                    })
            
            cursor.close()
            return {
                "success": True,
                "tables": tables,
                "count": len(tables)
            }
        except pyodbc.Error as e:
            return {"success": False, "error": str(e)}
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    def get_table_structure(self, table_name: str) -> Dict[str, Any]:
        """
        Get the structure of a specific table.
        
        Args:
            table_name: Name of the table to get structure for
            
        Returns:
            Dictionary with column information
        """
        try:
            if not self.connection:
                return {"success": False, "error": "Not connected to database"}
            
            cursor = self.connection.cursor()
            columns = []
            
            for row in cursor.columns(table=table_name):
                columns.append({
                    "name": row.column_name,
                    "type": row.type_name,
                    "data_type": row.data_type,
                    "column_size": row.column_size,
                    "nullable": row.nullable == 1,
                    "remarks": row.remarks if row.remarks else ""
                })
            
            cursor.close()
            return {
                "success": True,
                "table_name": table_name,
                "columns": columns,
                "count": len(columns)
            }
        except pyodbc.Error as e:
            return {"success": False, "error": str(e)}
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    def execute_query(self, query: str, params: Optional[tuple] = None) -> Dict[str, Any]:
        """
        Execute a SQL query on the database.
        
        Args:
            query: SQL query to execute
            params: Optional parameters for parameterized query
            
        Returns:
            Dictionary with query results
        """
        try:
            if not self.connection:
                return {"success": False, "error": "Not connected to database"}
            
            cursor = self.connection.cursor()
            
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            
            # Check if the query returns data (SELECT) or just executes
            if cursor.description:
                columns = [column[0] for column in cursor.description]
                rows = cursor.fetchall()
                
                # Convert to list of dictionaries
                results = []
                for row in rows:
                    results.append(dict(zip(columns, row)))
                
                cursor.close()
                return {
                    "success": True,
                    "columns": columns,
                    "rows": results,
                    "row_count": len(results),
                    "query_type": "SELECT"
                }
            else:
                # For INSERT, UPDATE, DELETE statements
                row_count = cursor.rowcount
                cursor.close()
                return {
                    "success": True,
                    "row_count": row_count,
                    "query_type": "ACTION"
                }
                
        except pyodbc.Error as e:
            return {"success": False, "error": str(e)}
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    def get_table_data(
        self, 
        table_name: str, 
        limit: Optional[int] = None,
        offset: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        Get data from a table with optional pagination.
        
        Args:
            table_name: Name of the table to get data from
            limit: Maximum number of rows to return
            offset: Number of rows to skip
            
        Returns:
            Dictionary with table data
        """
        try:
            if not self.connection:
                return {"success": False, "error": "Not connected to database"}
            
            # Build query with optional LIMIT and OFFSET
            query = f"SELECT * FROM [{table_name}]"
            
            if offset:
                query += f" OFFSET {offset}"
            
            if limit:
                query += f" TOP {limit}"
            
            cursor = self.connection.cursor()
            cursor.execute(query)
            
            if cursor.description:
                columns = [column[0] for column in cursor.description]
                rows = cursor.fetchall()
                
                results = []
                for row in rows:
                    results.append(dict(zip(columns, row)))
                
                cursor.close()
                return {
                    "success": True,
                    "table_name": table_name,
                    "columns": columns,
                    "rows": results,
                    "row_count": len(results)
                }
            else:
                cursor.close()
                return {"success": False, "error": "No data returned"}
                
        except pyodbc.Error as e:
            return {"success": False, "error": str(e)}
        except Exception as e:
            return {"success": False, "error": str(e)}
    
    def get_connection_status(self) -> Dict[str, Any]:
        """Get the current connection status."""
        return {
            "connected": self.connection is not None,
            "db_path": self.current_db_path
        }


# Global Access database instance
access_db = AccessDatabase()

# Create MCP server
app = Server("access-mcp")


@app.list_tools()
async def list_tools() -> List[Tool]:
    """List available tools."""
    return [
        Tool(
            name="access_connect",
            description="Connect to a Microsoft Access database (.accdb or .mdb)",
            inputSchema={
                "type": "object",
                "properties": {
                    "db_path": {
                        "type": "string",
                        "description": "Path to the Access database file"
                    },
                    "password": {
                        "type": "string",
                        "description": "Optional database password"
                    }
                },
                "required": ["db_path"]
            }
        ),
        Tool(
            name="access_disconnect",
            description="Disconnect from the current Access database",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        ),
        Tool(
            name="access_get_tables",
            description="Get list of all tables in the connected Access database",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        ),
        Tool(
            name="access_get_table_structure",
            description="Get the structure (columns) of a specific table",
            inputSchema={
                "type": "object",
                "properties": {
                    "table_name": {
                        "type": "string",
                        "description": "Name of the table to get structure for"
                    }
                },
                "required": ["table_name"]
            }
        ),
        Tool(
            name="access_execute_query",
            description="Execute a SQL query on the Access database",
            inputSchema={
                "type": "object",
                "properties": {
                    "query": {
                        "type": "string",
                        "description": "SQL query to execute"
                    },
                    "params": {
                        "type": "array",
                        "description": "Optional parameters for parameterized query",
                        "items": {"type": "string"}
                    }
                },
                "required": ["query"]
            }
        ),
        Tool(
            name="access_get_table_data",
            description="Get data from a specific table with optional limit",
            inputSchema={
                "type": "object",
                "properties": {
                    "table_name": {
                        "type": "string",
                        "description": "Name of the table to get data from"
                    },
                    "limit": {
                        "type": "integer",
                        "description": "Maximum number of rows to return (optional)"
                    },
                    "offset": {
                        "type": "integer",
                        "description": "Number of rows to skip (optional)"
                    }
                },
                "required": ["table_name"]
            }
        ),
        Tool(
            name="access_connection_status",
            description="Get the current database connection status",
            inputSchema={
                "type": "object",
                "properties": {}
            }
        )
    ]


@app.call_tool()
async def call_tool(name: str, arguments: Any) -> List[TextContent]:
    """Handle tool calls."""
    
    if name == "access_connect":
        result = access_db.connect(
            db_path=arguments.get("db_path"),
            password=arguments.get("password")
        )
        return [TextContent(type="text", text=str(result))]
    
    elif name == "access_disconnect":
        result = access_db.disconnect()
        return [TextContent(type="text", text=str(result))]
    
    elif name == "access_get_tables":
        result = access_db.get_tables()
        return [TextContent(type="text", text=str(result))]
    
    elif name == "access_get_table_structure":
        result = access_db.get_table_structure(
            table_name=arguments.get("table_name")
        )
        return [TextContent(type="text", text=str(result))]
    
    elif name == "access_execute_query":
        params = arguments.get("params")
        result = access_db.execute_query(
            query=arguments.get("query"),
            params=tuple(params) if params else None
        )
        return [TextContent(type="text", text=str(result))]
    
    elif name == "access_get_table_data":
        result = access_db.get_table_data(
            table_name=arguments.get("table_name"),
            limit=arguments.get("limit"),
            offset=arguments.get("offset")
        )
        return [TextContent(type="text", text=str(result))]
    
    elif name == "access_connection_status":
        result = access_db.get_connection_status()
        return [TextContent(type="text", text=str(result))]
    
    else:
        return [TextContent(type="text", text=f"Unknown tool: {name}")]


async def main():
    """Run the MCP server."""
    import argparse
    
    parser = argparse.ArgumentParser(description='Access Database MCP Server')
    parser.add_argument('--mode', choices=['stdio', 'tcp'], default='stdio',
                        help='Communication mode: stdio (default) or tcp')
    parser.add_argument('--host', default='0.0.0.0', help='TCP host (default: 0.0.0.0)')
    parser.add_argument('--port', type=int, default=5000, help='TCP port (default: 5000)')
    
    args = parser.parse_args()
    
    if args.mode == 'tcp':
        # Run TCP server
        from mcp.server import Server
        import socket
        
        print(f"Starting TCP MCP server on {args.host}:{args.port}")
        
        async def handle_client(reader, writer):
            addr = writer.get_extra_info('peername')
            print(f"Client connected: {addr}")
            
            await app.run(
                reader,
                writer,
                app.create_initialization_options()
            )
            
            print(f"Client disconnected: {addr}")
        
        server = await asyncio.start_server(
            handle_client, args.host, args.port
        )
        
        async with server:
            await server.serve_forever()
    else:
        # Run stdio server (default)
        async with stdio_server() as (read_stream, write_stream):
            await app.run(
                read_stream,
                write_stream,
                app.create_initialization_options()
            )


if __name__ == "__main__":
    import asyncio
    asyncio.run(main())
