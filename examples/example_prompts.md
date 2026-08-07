# Example prompts

Scenarios to run by hand against live Office. Each one exercises a group of
tools end to end. Ask Claude in plain language, the way you normally would.

Before you start, make sure the MCP server is registered and that Office is
installed. You do not need to open anything first.

## PowerPoint

**A short deck with a consistent look**

> Build a 3 slide deck about the history of ChatGPT. Set a dark theme with a
> green accent, put the background on the master so it applies to every slide,
> and add entrance animations plus transitions. Save it to my desktop.

Covers: `ppt_create_presentation`, `ppt_set_theme_colors`,
`ppt_set_master_background`, `ppt_add_slide`, `ppt_add_textbox`,
`ppt_add_shape`, `ppt_add_animation`, `ppt_set_transition`, `ppt_save`.

**Check your own layout**

> Export slide 2 to a PNG and look at it. If anything overlaps or the spacing
> looks wrong, fix it and export again.

Covers: `ppt_export_slide`, `ppt_get_slide_content`, `ppt_set_shape_position`,
`ppt_delete_shape`.

**Charts that match the slide**

> Add a column chart with our quarterly numbers, then restyle it to the deck
> colours: green bars, muted axis text, transparent background, no legend, and
> start the value axis at zero.

Covers: `ppt_add_chart`, `ppt_format_chart`.

**Diagrams and structure**

> Show me the SmartArt process layouts, then add a three step diagram to slide 2.
> Group the cards on slide 1 and align them to the top.

Covers: `ppt_list_smartart_layouts`, `ppt_add_smartart`, `ppt_group_shapes`,
`ppt_align_shapes`, `ppt_distribute_shapes`.

**Navigation and sections**

> Add a button on slide 1 that jumps to slide 3, split the deck into two
> sections, and put slide numbers in the footer.

Covers: `ppt_add_hyperlink`, `ppt_add_section`, `ppt_list_sections`,
`ppt_set_headers_footers`.

## Excel

**A simulation with a summary**

> Simulate 1000 coin flips in Excel. Freeze the random values so they stop
> changing, add running totals and frequency columns, and build a summary block
> with the count of heads, the frequency, the standard error and the longest run.

Covers: `xl_create_workbook`, `xl_set_range`, `xl_copy_range` with
`paste="values"`, `xl_set_formula`, `xl_set_cell_format`, `xl_freeze_panes`.

Note the freeze step. `RANDBETWEEN` is volatile and recalculates on every sheet
change, so without copying the values in place the numbers in a report will not
match the sheet.

**Look at the formatting**

> Export the summary block as an image so I can see how it looks.

Covers: `xl_export_range_image`.

**Cleaning up data**

> Sort the table by year, turn on the AutoFilter, add a dropdown in the status
> column with the allowed values, and colour the values column with a colour
> scale.

Covers: `xl_sort_range`, `xl_set_autofilter`, `xl_add_data_validation`,
`xl_apply_conditional_formatting`.

**Reading back**

> Show me the formulas in column E, not the results.

Covers: `xl_get_cell_formula`, `xl_get_used_range`.

## Word

**A short report**

> Write a two page report on the coin flip simulation. Pull the numbers from the
> spreadsheet rather than retyping them. Include a results table and a footnote
> about the random number generator.

Covers: `doc_create_document`, `doc_add_heading`, `doc_add_paragraph`,
`doc_insert_table`, `doc_format_table`, `doc_add_footnote`,
`doc_get_document_info`.

**A thesis skeleton**

> Set this document up like a master's thesis: Times New Roman 12, line spacing
> 1.5, justified text, first line indent, mirror margins with a binding gutter.
> Add a title page, numbered chapters, and a table of contents after the title
> page.

Covers: `doc_set_default_font`, `doc_set_page_margins`, `doc_set_page_setup`,
`doc_set_paragraph_format`, `doc_set_heading_numbering`,
`doc_insert_table_of_contents`, `doc_update_fields`.

Remember `doc_update_fields`. A table of contents inserted before the chapters
exist stays empty until it is refreshed.

**Figures and captions**

> Insert the chart image at 13 cm wide, caption it as Figure 1, and add a table
> of figures at the end.

Covers: `doc_insert_image` with `unit="cm"`, `doc_add_caption`,
`doc_insert_table_of_figures`.

**Editing, not just appending**

> Read paragraphs 1 to 10 with their styles, delete the third one, and insert a
> new paragraph before the second.

Covers: `doc_get_paragraph`, `doc_delete_paragraph`, `doc_insert_paragraph`.

## Across apps

> Read the results from the spreadsheet, write them up in Word, and put the same
> numbers on a slide. Export all three to PDF when you are done.

Covers: `xl_get_range_values`, `doc_add_paragraph`, `ppt_add_textbox`,
`xl_export_pdf`, `doc_export_pdf`, `ppt_export_pdf`.

## Diagnostics

> Is the bridge running, and which Office apps are connected right now?

Covers: `office_status`.
