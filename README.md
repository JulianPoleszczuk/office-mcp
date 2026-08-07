# office-mcp

Let Claude drive Microsoft Office on Windows.

This is an MCP server for PowerPoint, Excel and Word. You ask for something in
plain language, and Claude builds it in the app you already have open. No
uploading, no downloading, no generated file to import. You watch the slides and
cells change on screen.

```
> Build a 3 slide deck about our Q3 numbers. Dark theme, green accent,
> add the chart from the spreadsheet, and animate the bullets.
```

## What it can do

- **PowerPoint**: slides, text, shapes, charts, tables, images, SmartArt,
  themes, animations, transitions, sections, slide shows
- **Excel**: values, formulas, formatting, sorting, filters, dropdowns,
  conditional formatting, charts, pivot tables
- **Word**: paragraphs, styles, tables, images, footnotes, captions, a table of
  contents, and everything a thesis layout needs

It can also **look at its own work**. Claude exports a slide or a cell range to
an image, checks whether the layout came out right, and fixes it. That one
feature is the difference between a deck that looks designed and one that looks
generated.

129 tools in total. The full list is in the [tool reference](docs/TOOLS.md).

## Install

You need Windows, Office 2019 or newer, and Python 3.11 or newer.

```powershell
git clone https://github.com/JulianPoleszczuk/office-mcp.git
cd office-mcp

python -m venv .venv
.\.venv\Scripts\Activate.ps1

pip install -r requirements.txt
python .venv\Scripts\pywin32_postinstall.py -install
```

That last line registers the DLLs that Office automation needs. Skipping it is
the usual reason things fail on the first run.

## Connect it to Claude

In Claude Code:

```powershell
claude mcp add office -- C:\path\to\office-mcp\.venv\Scripts\python.exe C:\path\to\office-mcp\server.py
```

In Claude Desktop, add this to `%APPDATA%\Claude\claude_desktop_config.json` and
restart the app:

```json
{
  "mcpServers": {
    "office": {
      "command": "C:\\path\\to\\office-mcp\\.venv\\Scripts\\python.exe",
      "args": ["C:\\path\\to\\office-mcp\\server.py"]
    }
  }
}
```

That is it. You do not need to start Office or anything else by hand. Ask Claude
to make a presentation and it will open PowerPoint for you.

## Documentation

| Page | What is in it |
|---|---|
| [Tool reference](docs/TOOLS.md) | All 129 tools, grouped by what you want to do |
| [Recipes](docs/RECIPES.md) | Preview loops, styling a deck once, writing a thesis |
| [How it works](docs/ARCHITECTURE.md) | Architecture, protocol, settings, tests, layout |
| [Errors and gotchas](docs/GOTCHAS.md) | Error types, and Office behaviour worth knowing |
| [Example prompts](examples/example_prompts.md) | Things to try, end to end |

## Good to know

**Windows only.** Office automation runs on COM, which does not exist anywhere
else.

**It uses the Office you already have.** The server attaches to a running app
rather than opening a second copy, so your open documents stay as they are.

**Nothing leaves your machine.** The server talks to Office locally over COM.

If something behaves oddly, the [gotchas page](docs/GOTCHAS.md) is worth a look
first. Office has a number of quirks that report success while quietly doing the
wrong thing, and most of them are documented there.

## Licence

MIT
