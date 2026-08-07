# Recipes

Patterns that come up again and again when driving Office through this
server.

Back to the [main README](../README.md).

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
