# Errors and gotchas

Back to the [main README](../README.md).

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
