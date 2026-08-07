# office-mcp

An MCP server that lets Claude control **open** Microsoft Office apps on Windows:
PowerPoint, Excel and Word. Claude builds and edits real documents through COM
automation, and you watch the changes happen live in the Office window.

There is no intermediate file generation. The model works on the document you
already have open.

- 129 tools: 53 for PowerPoint, 36 for Excel, 39 for Word, plus a status tool
- Create, read, edit, format, chart, animate, export
- Preview built in: the model can export a slide or a cell range to PNG and look
  at its own work instead of guessing
- Attaches to an Office instance you already have running
- 367 unit and integration tests that run without Office installed

## Requirements

| Item | Version |
|---|---|
| System | Windows 10 or 11. Windows only, because Office COM does not exist elsewhere |
| Office | Microsoft Office 2019. Also works with 2016 and 365 desktop |
| Python | 3.11 or newer |
| Libraries | `mcp`, `pywin32`, `pytest` |

## Install

```powershell
git clone https://github.com/JulianPoleszczuk/office-mcp.git
cd office-mcp

python -m venv .venv
.\.venv\Scripts\Activate.ps1

pip install -r requirements.txt
python .venv\Scripts\pywin32_postinstall.py -install
```

The `pywin32_postinstall.py -install` step registers the DLLs that COM needs.
Skip it and `win32com.client` may throw `ImportError: DLL load failed` on the
first call.

Quick check that everything can see Office:

```powershell
python -m bridge.main --log-level DEBUG
```

## Set up Claude

In Claude Code:

```powershell
claude mcp add office -- C:\path\to\office-mcp\.venv\Scripts\python.exe C:\path\to\office-mcp\server.py
```

In Claude Desktop, edit `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "office": {
      "command": "C:\\path\\to\\office-mcp\\.venv\\Scripts\\python.exe",
      "args": ["C:\\path\\to\\office-mcp\\server.py"],
      "env": {
        "OFFICE_BRIDGE_PORT": "8765"
      }
    }
  }
}
```

Restart Claude Desktop afterwards. The `ppt_*`, `xl_*` and `doc_*` tools show up
in the tool list.

You do not need to start anything by hand. The Bridge starts on the first tool
call.

### Environment variables

| Variable | Default | Meaning |
|---|---|---|
| `OFFICE_BRIDGE_HOST` | `127.0.0.1` | Bridge listen address |
| `OFFICE_BRIDGE_PORT` | `8765` | Bridge port |
| `OFFICE_BRIDGE_TIMEOUT` | `15` | Time limit for one COM call, in seconds |
| `OFFICE_BRIDGE_AUTOSTART` | `1` | Set to `0` to stop the MCP server starting the Bridge |
| `OFFICE_MCP_CALL_TIMEOUT` | `60` | How long the MCP server waits for the Bridge |
| `OFFICE_BRIDGE_LOG_LEVEL` | `INFO` | Bridge log level |

## How it works

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

## Tools

There are 129 of them, so the full list lives in its own page:

**[Tool reference](docs/TOOLS.md)**

| Group | Count | Covers |
|---|---|---|
| [PowerPoint](docs/TOOLS.md#powerpoint) | 53 | Slides, text, shapes, charts, SmartArt, themes, animation, export |
| [Excel](docs/TOOLS.md#excel) | 36 | Sheets, values, formulas, formatting, sorting, validation, charts, pivots, export |
| [Word](docs/TOOLS.md#word) | 39 | Paragraphs, styles, tables, images, captions, thesis layout, export |
| [Diagnostics](docs/TOOLS.md#diagnostics) | 1 | Bridge and COM connection state |

## Letting the model see its own work

Without a preview the model places things blind. It cannot tell that a footer
overlaps a panel or that a column is too narrow. Every app has a way to show the
result:

| App | Preview | Whole document |
|---|---|---|
| PowerPoint | `ppt_export_slide` to PNG or JPG | `ppt_export_pdf` |
| Excel | `xl_export_range_image` to PNG or JPG | `xl_export_pdf` |
| Word | none | `doc_export_pdf` |

A normal loop looks like this:

```
ppt_add_textbox(...)
ppt_export_slide(1, "preview.png")      -> look at it
ppt_set_shape_position(1, 42, top=496)  -> fix it
ppt_export_slide(1, "preview.png")      -> check again
```

Excel cannot export a range to an image directly. `xl_export_range_image` copies
the range to the clipboard as a bitmap, drops it on a temporary chart object,
which can export, and then removes that chart.

## Styling a deck once

`ppt_set_theme_colors`, `ppt_set_theme_fonts` and `ppt_set_master_background`
set the look **once**, on the master, instead of repeating the same hex colour
on every shape:

```
ppt_set_theme_colors({"dark1": "#0B1014", "light1": "#ECF2F0", "accent1": "#10A37F"})
ppt_set_theme_fonts(major="Segoe UI", minor="Segoe UI")
ppt_set_master_background(color="#0B1014")
```

`apply_to_slides=True`, the default, turns on `FollowMasterBackground`, so
slides that had their own background from `ppt_set_background` go back to the
master.

## Writing a thesis in Word

The Word tools cover a full dissertation layout: title page, table of contents
after it, numbered chapters, figure and table captions, a table of figures,
footnotes and two sided binding.

```
doc_set_default_font("Times New Roman", 12)
doc_set_page_margins(2.5, 2.5, 3.5, 2.5, unit="cm")
doc_set_page_setup(gutter=0.5, mirror_margins=True)
doc_set_paragraph_format(style="Normal", line_spacing=1.5,
                         first_line_indent=1.25, alignment="justify", unit="cm")

... write chapters with doc_add_heading and doc_add_paragraph ...

doc_set_heading_numbering(levels=3)                        # 1., 1.1, 1.1.1
doc_insert_table_of_contents(levels=3, position=<paragraph>)
doc_update_fields()                                        # without this the TOC is empty
```

**Caption labels.** `doc_add_caption(label="figure")` uses a built in label, and
Word decides whether to call it "Figure" or something else. That can differ
between installations. If you need a specific word, pass it directly:
`label="Figure"`, `label="Rysunek"`, `label="Abbildung"`. The text goes into the
document as written, and `doc_insert_table_of_figures(label=...)` collects those
captions.

## Errors

Controllers never let a raw `pywintypes.com_error` through. Every COM error is
mapped to a type from `bridge/utils/errors.py`:

| Type | When |
|---|---|
| `ComConnectionError` | App closed, not responding, or rejected the call because a dialog is open |
| `ComTimeoutError` | The COM call passed its time limit |
| `DocumentNotFoundError` | Missing file, missing target directory, or no open document |
| `InvalidReferenceError` | Bad slide index, missing sheet, bad range, unknown layout or chart name |
| `UnsupportedOperationError` | Not available in the installed Office version |
| `ProtocolError` | Bad protocol message, unknown action, missing parameters |
| `BridgeUnavailable` | The MCP server could not reach the Bridge process |

Example error reply:

```json
{
  "ok": false,
  "error": {
    "type": "InvalidReferenceError",
    "message": "slide_index = 7 is outside the range 1..3"
  }
}
```

## Things to watch out for

These are real behaviours found while testing against live Office. Most of them
report success while doing the wrong thing, so they are worth knowing.

**Windows only.** Automation is built on COM. `server.py` and `bridge/main.py`
exit with a clear message on other platforms.

**Office must be installed**, the desktop version, not the web one.

**COM is fragile between versions.** The same member can be a method in one
build and a property in another. Versions older than 2016 may lack `AddChart2`
(there is a fallback to `AddChart`) and some conditional formatting types.

**An open dialog blocks the app.** If a save dialog is open, Office rejects COM
calls and the tool returns `ComConnectionError`.

**`ppt_open_presentation` does not change the "active" presentation.**
PowerPoint ignores `Windows.Activate()` when it is not in the foreground, so
`ActivePresentation` can point at a different file than the one you just opened.
The controller therefore remembers the path it works on rather than trusting
`ActivePresentation`. Excel and Word do not have this problem.

**Excel `Range.Sort` parameters are sticky.** Excel remembers `Orientation`,
`MatchCase` and `SortMethod` from the previous sort in the session. Leaving
`Orientation` out can sort left to right and reorder columns instead of rows, so
`xl_sort_range` passes `xlSortColumns` explicitly every time.

**A line chart takes its colour from the outline, not the fill.**
`series_colors` sets both, because `Format.Fill` works on bars and pies while
`Format.Line` works on lines and scatter points.

**Office picks the value axis range itself.** With close values it can start far
from zero, which makes a 486 against 514 gap look like double. Use
`value_axis_min` and `value_axis_max` to fix that.

**`Chart.Export` in Excel can write a zero length file** without reporting an
error. `xl_export_range_image` checks the file size and turns that silent
failure into a clear error.

**`Range.Collapse(wdCollapseEnd)` in Word lands past the paragraph mark**, that
is, inside the next paragraph. A footnote or hyperlink inserted that way shows
up before the first word of the following paragraph. `doc_add_footnote` and
`doc_add_hyperlink` anchor before the mark instead.

**`doc_add_paragraph("")` does not create an empty paragraph.** The controller
reuses the empty paragraph at the end of the document on purpose. Use
`doc_set_paragraph_format(space_before=...)` for vertical spacing.

**Nested lists need a gallery template.** `ApplyNumberDefault` makes a single
level list and going to level 2 fails with OLE error 0x800a1200. The controller
spots items with `level > 1` and reaches for a multi level template.

**Word translates built in table style names**, the same way it translates
SmartArt layout names. `doc_format_table` takes language independent names such
as `light_grid` and `medium_shading1` and maps them to `wdStyle` constants.

**SmartArt layout names are localised.** `ppt_list_smartart_layouts` returns a
`key` (the tail of the URN, such as `bProcess3`) and a `category` next to the
name. Both are the same in every language version, so prefer them.

**`SmartArtLayouts` can start returning "access denied."** After heavy SmartArt
work in one COM session the collection becomes unreachable and only a PowerPoint
restart helps. This is Office behaviour, not something the code controls.

**`ExportAsFixedFormat` works in Excel and Word but not PowerPoint.** In
PowerPoint the pywin32 wrapper puts `PyOleEmpty` into the `ExternalExporter`
parameter and the call cannot be made at all, so `ppt_export_pdf` uses
`SaveCopyAs` instead.

**The first `ppt_add_chart` in a session takes about 13 seconds**, because
`ChartData.Activate()` has to start Excel. That is close to the default
`OFFICE_BRIDGE_TIMEOUT` of 15. For decks with charts, raise the limit or call
any `xl_*` tool first to warm Excel up.

**Charts cannot be styled per data point.** `ppt_format_chart` and
`xl_format_chart` cover series colours, axis and label text, background, legend
and gridlines, but not individual points or a secondary axis.

**Saving a new document needs a path** the first time, for example
`ppt_save(path=...)`.

**The Bridge does not reload code.** It outlives MCP client restarts by design,
so after editing a controller you have to kill it. Otherwise new actions return
`ProtocolError: Unknown action`.

**The Bridge listens on localhost only and has no authentication.** Do not
expose its port.

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

## Licence

MIT
