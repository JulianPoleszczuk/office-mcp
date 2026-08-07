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

### PowerPoint

Coordinates are in points. A 16:9 slide is 960 x 540 pt, and 1 cm is 28.35 pt.
Slide indexes start at 1.

**Files and slides**

| Tool | What it does |
|---|---|
| `ppt_create_presentation(path, template=None)` | New presentation, optionally from a `.potx` |
| `ppt_open_presentation(path)` | Open a file, or activate one already open |
| `ppt_save(path=None)` | Save, or save as a new file |
| `ppt_close(save=True)` | Close the presentation |
| `ppt_get_presentation_info()` | Slide count, slide size, theme, path |
| `ppt_list_slides()` | Titles and layouts of every slide |
| `ppt_get_slide_content(slide_index)` | Shapes, positions, text, notes |
| `ppt_add_slide(layout, index=None, title=None)` | Add a slide |
| `ppt_delete_slide(slide_index)` | Delete a slide |
| `ppt_duplicate_slide(slide_index)` | Duplicate a slide |
| `ppt_reorder_slide(from_index, to_index)` | Move a slide |
| `ppt_set_slide_layout(slide_index, layout_name)` | Change the layout |
| `ppt_copy_slide_to(slide_index, target_path, position=None)` | Copy a slide into another file |

**Text and content**

| Tool | What it does |
|---|---|
| `ppt_set_title(slide_index, text)` | Set the slide title |
| `ppt_add_textbox(slide_index, text, left, top, width, height, ...)` | Text box |
| `ppt_add_bullet_list(slide_index, items, placeholder="content")` | Bulleted list with levels |
| `ppt_set_speaker_notes(slide_index, text)` | Speaker notes |
| `ppt_find_replace_text(old_text, new_text, slide_index=None, ...)` | Replace text, tables and groups included |
| `ppt_set_text_style(slide_index, shape_id, ...)` | Font, size, colour, bold |
| `ppt_set_paragraph_format(slide_index, shape_id, ...)` | Line spacing, alignment, anchor, margins |

**Shapes and objects**

| Tool | What it does |
|---|---|
| `ppt_add_shape(slide_index, shape_type, left, top, width, height, ...)` | Shape |
| `ppt_add_image(slide_index, image_path, left, top, width=None, height=None)` | Image |
| `ppt_add_chart(slide_index, chart_type, categories, series_data, ...)` | Chart with data |
| `ppt_add_table(slide_index, rows, cols, data, left, top, width, height)` | Table |
| `ppt_add_media(slide_index, media_path, left, top, ..., autoplay=False)` | Video or audio |
| `ppt_add_smartart(slide_index, layout, items, left, top, width, height)` | SmartArt diagram |
| `ppt_list_smartart_layouts(search=None, category=None)` | SmartArt layouts: key, name, category |
| `ppt_set_shape_format(slide_index, shape_id, ...)` | Gradient, transparency, shadow, outline, corner radius |
| `ppt_set_shape_position(slide_index, shape_id, ...)` | Move, scale, rotate |
| `ppt_set_shape_order(slide_index, shape_id, order)` | Layer: front, back, forward, backward |
| `ppt_delete_shape(slide_index, shape_id)` | Delete a shape |
| `ppt_group_shapes(slide_index, shape_ids, name=None)` | Group shapes |
| `ppt_ungroup_shapes(slide_index, shape_id)` | Ungroup |
| `ppt_align_shapes(slide_index, shape_ids, align, ...)` | Align to each other or to the slide |
| `ppt_distribute_shapes(slide_index, shape_ids, direction, ...)` | Even spacing, needs 3 or more |
| `ppt_format_chart(slide_index, shape_id, ...)` | Series colours, axes, legend, labels, background |

**Design, motion and navigation**

| Tool | What it does |
|---|---|
| `ppt_get_theme()` | Theme palette and fonts |
| `ppt_set_theme_colors(colors)` | Change the theme palette |
| `ppt_set_theme_fonts(major, minor)` | Heading and body fonts |
| `ppt_apply_theme(theme_name_or_path)` | Theme from a `.thmx`/`.potx` or the Office gallery |
| `ppt_set_master_background(color, image_path, apply_to_slides=True)` | Background once, on the master |
| `ppt_set_background(slide_index, color=None, image_path=None)` | Background of one slide |
| `ppt_add_animation(slide_index, shape_id, effect, trigger, ...)` | Animate a shape |
| `ppt_list_animations(slide_index)` | Animations in playback order |
| `ppt_set_transition(effect, slide_index=None, ...)` | Slide transition |
| `ppt_add_hyperlink(slide_index, shape_id, url=None, target_slide=None, ...)` | Link out, or jump to a slide |
| `ppt_set_headers_footers(slide_index=None, footer_text=None, ...)` | Footer, slide number, date |
| `ppt_list_sections()`, `ppt_add_section(name, before_slide)`, `ppt_delete_section(...)` | Sections |
| `ppt_slideshow(command, slide_index=None)` | Slide show: start, stop, goto |
| `ppt_export_slide(slide_index, path, width=None, height=None)` | Slide to an image |
| `ppt_export_pdf(path)` | Presentation to PDF |

Layouts: `title`, `title_content`, `two_content`, `title_only`, `blank`,
`section_header`, `comparison`, `picture_with_caption`, `content_with_caption`,
`chart`, `table`, `four_objects`.

Charts: `bar`, `column`, `line`, `pie`, `area`, `scatter`, `doughnut`, `radar`,
`bubble`.

Shapes: `rectangle`, `rounded_rectangle`, `oval`, `triangle`, `diamond`, `star`,
`arrow_right`, `callout`, `cloud`, `hexagon`, `chevron`. Both `fill_color` and
`line_color` accept `"none"`.

### Excel

Sheets can be named or numbered. Ranges use A1 notation.

| Tool | What it does |
|---|---|
| `xl_create_workbook(path)` | New workbook |
| `xl_open_workbook(path)` | Open a file, or activate one already open |
| `xl_save(path=None)`, `xl_close(save=True)` | Save and close |
| `xl_get_workbook_info()` | Sheets, their data ranges, active sheet, path |
| `xl_add_sheet(name, index=None)`, `xl_delete_sheet(name)`, `xl_rename_sheet(...)` | Manage sheets |
| `xl_get_range_values(sheet, range_ref)` | Read a range as a 2D array |
| `xl_get_used_range(sheet)` | The filled area, with data |
| `xl_get_cell_formula(sheet, range_ref)` | Formulas plus their results |
| `xl_set_cell(sheet, cell_ref, value)` | One cell |
| `xl_set_range(sheet, start_cell, values_2d)` | A whole matrix at once |
| `xl_set_formula(sheet, cell_ref, formula)` | Formula and computed result |
| `xl_clear_range(sheet, range_ref, contents_only=True)` | Clear a range |
| `xl_copy_range(sheet, range_ref, target_cell, ...)` | Copy: all, values, formats |
| `xl_find_replace(old_text, new_text, sheet=None, ...)` | Replace text |
| `xl_sort_range(sheet, range_ref, sort_by, order, has_headers)` | Sort |
| `xl_set_autofilter(sheet, range_ref=None, enable=True)` | AutoFilter |
| `xl_insert_rows(...)`, `xl_delete_rows(...)` | Rows |
| `xl_insert_columns(...)`, `xl_delete_columns(...)` | Columns |
| `xl_set_column_width(sheet, column, width)` | Column width, `"auto"` fits |
| `xl_set_row_height(sheet, row, height)` | Row height, `"auto"` fits |
| `xl_set_cell_format(sheet, range_ref, ...)` | Font, colours, number format, alignment, wrap |
| `xl_merge_cells(sheet, range_ref, center=True)` | Merge cells |
| `xl_apply_conditional_formatting(sheet, range_ref, rule_type, params)` | Conditional formatting |
| `xl_add_data_validation(sheet, range_ref, ...)` | Dropdowns and value rules |
| `xl_freeze_panes(sheet, cell_ref)` | Freeze panes |
| `xl_add_chart(sheet, chart_type, data_range, ...)` | Chart |
| `xl_format_chart(sheet, chart, ...)` | Series colours, axes, legend, labels |
| `xl_create_table(sheet, range_ref, table_name, ...)` | Native Excel table |
| `xl_add_pivot_table(sheet, source_range, dest_cell, rows, columns, values, ...)` | Pivot table |
| `xl_export_range_image(sheet, range_ref, path)` | A range as an image |
| `xl_export_pdf(path, sheet=None, range_ref=None)` | Workbook, sheet or range to PDF |

Conditional formatting rules: `cell_value` (with operators `greater`, `less`,
`equal`, `not_equal`, `greater_equal`, `less_equal`, `between`, `not_between`),
`expression`, `text_contains`, `color_scale`, `data_bar`.

Pivot functions: `sum`, `count`, `average`, `max`, `min`, `product`,
`count_numbers`, `std_dev`.

### Word

Paragraphs are indexed from 1. Style names can be given in English even in a
localised Word.

| Tool | What it does |
|---|---|
| `doc_create_document(path, template=None)` | New document, optionally from a `.dotx` |
| `doc_open_document(path)` | Open a file, or activate one already open |
| `doc_save(path=None)`, `doc_close(save=True)` | Save and close |
| `doc_get_document_info()` | Pages, words, characters, template, path |
| `doc_get_full_text()` | The whole text |
| `doc_get_outline()` | Heading tree with paragraph indexes |
| `doc_get_paragraph(paragraph_index, count=1)` | Read paragraphs with style and alignment |
| `doc_add_paragraph(text, style=None)` | Paragraph at the end |
| `doc_insert_paragraph(text, paragraph_index=None, after=False, style=None)` | Paragraph at a given place |
| `doc_delete_paragraph(paragraph_index, count=1)` | Delete paragraphs |
| `doc_add_heading(text, level=1)` | Heading, level 1 to 9 |
| `doc_add_bullet_list(items)`, `doc_add_numbered_list(items)` | Lists with levels |
| `doc_find_replace(old_text, new_text, match_case=False)` | Replace text |
| `doc_set_text_style(paragraph_index, ...)` | Font, size, colour, bold, italic |
| `doc_apply_style(paragraph_index, style_name)` | Paragraph style |
| `doc_set_paragraph_alignment(paragraph_index, alignment)` | Alignment |
| `doc_set_paragraph_format(paragraph_index=None, style=None, ...)` | Line spacing, indents, page breaks |
| `doc_set_default_font(name=None, size=None)` | The Normal style font |
| `doc_set_page_margins(top, bottom, left, right, unit="cm")` | Margins |
| `doc_set_page_setup(orientation, gutter, mirror_margins, ...)` | Binding, mirror margins, orientation |
| `doc_insert_page_break()`, `doc_insert_section_break(break_type, ...)` | Breaks |
| `doc_set_columns(count=1, section=1, spacing=None)` | Newspaper columns |
| `doc_insert_image(image_path, width=None, height=None, position, unit)` | Image |
| `doc_insert_table(rows, cols, data=None, position=None)` | Table |
| `doc_format_table(table_index, style, borders, header_bold, ...)` | Table formatting |
| `doc_add_hyperlink(url, text=None, paragraph_index=None, tooltip=None)` | Hyperlink |
| `doc_add_footnote(paragraph_index, text)` | Footnote |
| `doc_add_caption(paragraph_index, text, label, above=False)` | Numbered caption |
| `doc_insert_header(text, section=1)`, `doc_insert_footer(text, section=1)` | Header and footer |
| `doc_add_page_numbers(alignment="center", first_page=True)` | Page numbers |
| `doc_insert_table_of_contents(levels=3, position="start")` | Table of contents |
| `doc_insert_table_of_figures(label, position)` | Table of figures or tables |
| `doc_set_heading_numbering(enable=True, levels=3)` | Chapter numbering 1., 1.1, 1.1.1 |
| `doc_update_fields()` | Refresh tables, captions and numbering |
| `doc_export_pdf(path, open_after=False)` | Document to PDF |

### Diagnostics

| Tool | What it does |
|---|---|
| `office_status()` | Bridge state and the COM connection of all three apps |

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
|- examples/example_prompts.md  # end to end scenarios for live Office
```

## Licence

MIT
