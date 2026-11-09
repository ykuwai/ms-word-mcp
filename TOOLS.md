# Tools

By default the server lists three tools, and the model reaches these through
them. Set `WORD_TOOL_MODE=full` to list every tool by its own name. See
[the README](README.md#the-tool-list-catalog-mode).

Every tool that works on a document takes an optional `document` argument (a
full path or file name), so a call lands in the document it names whatever
another conversation has chosen.

| Area | Tools |
|---|---|
| Word and documents | `word_get_app_info`, `word_list_documents`, `word_activate_document`, `word_open_document`, `word_create_document`, `word_save_document`, `word_save_document_as` (also PDF), `word_close_document` |
| Reading | `word_get_document_info`, `word_get_outline`, `word_get_paragraphs` |
| Writing | `word_insert_paragraphs`, `word_set_paragraph_text`, `word_delete_paragraphs`, `word_find_replace` |
| Formatting | `word_format_text`, `word_set_paragraph_format` |
| Styles | `word_apply_style`, `word_list_styles`, `word_get_style`, `word_set_style` |
| Lists | `word_apply_list`, `word_set_list_level`, `word_restart_numbering`, `word_remove_list`, `word_get_list` |
| Tables | `word_insert_table`, `word_list_tables`, `word_get_table`, `word_set_table`, `word_set_table_cells`, `word_format_table_cells`, `word_insert_table_rows`, `word_delete_table_rows`, `word_insert_table_columns` (Windows), `word_delete_table_columns`, `word_merge_table_cells`, `word_delete_table` |
| Pictures, captions, cross-references | `word_insert_picture`, `word_list_pictures`, `word_set_picture`, `word_insert_caption`, `word_list_captions`, `word_insert_cross_reference` |
| The page | `word_get_page_setup`, `word_set_page_setup` (paper, margins, columns, document grid, vertical text), `word_get_headers_footers`, `word_set_header_footer`, `word_set_page_numbers`, `word_add_ruby` (furigana) |
| Equations | `word_insert_equation`, `word_list_equations`, `word_get_equation` |
| Floating shapes and pictures | `word_insert_shape`, `word_insert_text_box`, `word_insert_line`, `word_insert_floating_picture`, `word_list_shapes`, `word_set_shape`, `word_delete_shape` |
| Footnotes and endnotes | `word_insert_note`, `word_list_notes`, `word_set_note`, `word_delete_note` |
| Hyperlinks and bookmarks | `word_insert_hyperlink`, `word_list_hyperlinks`, `word_set_hyperlink`, `word_remove_hyperlink`, `word_add_bookmark`, `word_list_bookmarks`, `word_delete_bookmark` |
| Fields and tables of contents | `word_insert_table_of_contents`, `word_insert_field`, `word_list_fields`, `word_update_fields` |
| Tracked changes | `word_set_track_changes`, `word_list_revisions`, `word_accept_revisions`, `word_reject_revisions` |
| View | `word_get_view`, `word_set_view` |
| Seeing a page | `word_get_page_image`: a picture of a page as Word lays it out, so the model can check the layout |
| Undo | `word_undo`, `word_redo` |
