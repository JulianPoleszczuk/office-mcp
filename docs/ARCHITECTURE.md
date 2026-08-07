# How it works

Back to the [main README](../README.md).

```
Claude Code / Claude Desktop
        |  (stdio, MCP protocol)
        v
+---------------------+
|   MCP server        |  Python, official `mcp` SDK
|   (server.py)       |  Registers the ppt_*, xl_*, doc_* tools
+----------+----------+
           |  TCP (localhost, one JSON object per line)
           v
+---------------------+
|   Office Bridge     |  Python, pywin32 (win32com.client)
|   (bridge/*.py)     |  Holds live COM connections, one per app
+----------+----------+
           |  COM
           v
  PowerPoint / Excel / Word
```

**Why two layers.** The MCP client can restart `server.py` whenever it likes.
The Bridge keeps running, so COM connections stay alive and your open documents
are not disturbed. The Bridge is also a plain TCP server, so you can talk to it
with anything else, such as a debug script.

**App isolation.** Each Office app gets its own COM thread (an STA apartment),
its own connection object and its own state. A hung Word does not block Excel.
Every COM call has a time limit. Once it passes, the thread is dropped and the
connection is marked dead, then rebuilt on the next call.

**Lazy connect.** An app only starts when a tool needs it. The Bridge first
tries `GetActiveObject`, which attaches to a window you already opened. Only if
that fails does it `Dispatch` a new instance.

### Protocol

Every tool returns JSON in one shape:
`{"ok": true, "result": {...}}` or `{"ok": false, "error": {"type": ..., "message": ...}}`.

The Bridge speaks one JSON object per line:

```json
{"id": "uuid", "app": "powerpoint", "action": "add_slide", "params": {"layout": "title_content"}}
{"id": "uuid", "ok": true, "result": {"slide_index": 2}}
{"id": "uuid", "ok": false, "error": {"type": "ComConnectionError", "message": "PowerPoint is not responding"}}
```

You can drive it with a plain socket:

```powershell
python -c "import socket, json; s=socket.create_connection(('127.0.0.1',8765)); s.sendall(json.dumps({'id':'1','app':'excel','action':'get_workbook_info','params':{}}).encode()+b'\n'); print(s.recv(65536).decode())"
```


## Environment variables

| Variable | Default | Meaning |
|---|---|---|
| `OFFICE_BRIDGE_HOST` | `127.0.0.1` | Bridge listen address |
| `OFFICE_BRIDGE_PORT` | `8765` | Bridge port |
| `OFFICE_BRIDGE_TIMEOUT` | `15` | Time limit for one COM call, in seconds |
| `OFFICE_BRIDGE_AUTOSTART` | `1` | Set to `0` to stop the MCP server starting the Bridge |
| `OFFICE_MCP_CALL_TIMEOUT` | `60` | How long the MCP server waits for the Bridge |
| `OFFICE_BRIDGE_LOG_LEVEL` | `INFO` | Bridge log level |


## Project layout

```
office-mcp/
|- server.py                    # MCP server (stdio), tool registration, Bridge client
|- bridge/
|  |- main.py                   # TCP server, JSON-line protocol, routing
|  |- protocol.py               # Request/Response and line encoding
|  |- connection_manager.py     # lazy COM connect, STA threads, timeouts
|  |- controllers/
|  |  |- base.py                # action routing, com_error mapping, shared helpers
|  |  |- powerpoint.py          # PowerPointController
|  |  |- excel.py               # ExcelController
|  |  |- word.py                # WordController
|  |- utils/
|     |- com_helpers.py         # colours, units, Office constants, conversions
|     |- errors.py              # exception hierarchy
|- tests/                       # protocol, controller (mocked COM) and MCP layer tests
|- docs/TOOLS.md                # the full tool reference
|- examples/example_prompts.md  # end to end scenarios for live Office
```


## Tests

```powershell
python -m pytest -q
```

367 tests, all of them running without Office installed:

- `tests/test_bridge_protocol.py` covers protocol encoding and decoding, plus an
  integration test of the TCP server using a real socket and a fake controller
- `tests/test_powerpoint_controller.py`, `test_excel_controller.py` and
  `test_word_controller.py` cover the controllers with COM mocked out
- `tests/test_server_tools.py` covers the MCP layer, including a check that
  every action used in `server.py` exists in the matching controller

End to end scenarios to run by hand against live Office are in
`examples/example_prompts.md`.
