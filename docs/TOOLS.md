# Tool reference

Every tool in office-mcp, grouped by what you are trying to do. All of them
return JSON in one shape: `{"ok": true, "result": {...}}` on success, or
`{"ok": false, "error": {"type": ..., "message": ...}}` on failure.

Back to the [main README](../README.md).

## PowerPoint

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

## Excel

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

## Word

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

## Diagnostics

| Tool | What it does |
|---|---|
| `office_status()` | Bridge state and the COM connection of all three apps |
