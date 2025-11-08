MCP server config for CATIA V5

This project provides `mcp_server.py` which exposes CATIA V5 automation tools via FastMCP.

Files created:
- `mcpserver.json` — a suggested config that launches the server using stdio transport.

How to use

1. Ensure dependencies are installed (in the project venv):

   D:/hadi/github/catia-v5-mcp-server/.venv/Scripts/python.exe -m pip install -r requirements.txt

2. Launch the server using the config command (the config expects the venv python path):

   D:/hadi/github/catia-v5-mcp-server/.venv/Scripts/python.exe -u mcp_server.py

   - The `-u` flag forces unbuffered binary stdout/stderr which is recommended for stdio-based transports.

3. If your MCP client accepts a JSON config, point it to `mcpserver.json` or use the command from step 2.

Notes and tips

- `pywin32` provides the `win32com` and `pythoncom` modules; it must be installed and the server must run on Windows.
- `mcpserver.json` uses `transport: "stdio"` so the server communicates over stdin/stdout. If you want TCP or another transport, edit the config accordingly.
- If you change the venv location, update the `command` path in `mcpserver.json`.

Troubleshooting

- If you get ImportError for `win32com` or `pythoncom`, make sure you installed `pywin32` into the same Python interpreter you're using to run the server.
- To verify `pythoncom` availability from the project's venv:

  D:/hadi/github/catia-v5-mcp-server/.venv/Scripts/python.exe -c "import pythoncom; print('pythoncom OK')"

