# Capability Map

This map groups backend-owned action concepts for `excel.agent.run`. It is for routing and explanation, not for normal agents to call primitive MCP tools directly.

## Discovery And Reads

- Workbook/session: `list_open_workbooks`, `get_workbook_info`, `prepare_session`, `refresh_workbook_snapshot`, `get_workbook_snapshot`, `detect_external_changes`.
- Target discovery: `find_target`, `read_schema`, `get_range_summary`, `read_range_compact`.
- Range metadata: `read_hyperlinks`, `read_comments`, `read_notes`, `read_merged_cells`, `read_data_validation`, `read_conditional_formatting`, `search_range`, `find_blank_cells`, `find_range_errors`.
- Names/regions/templates: `read_named_item`, `read_region`, `list_templates`, `read_template`, `detect_templates`, `infer_template_regions`.

## Mutations

- Values/formulas/formats: `write_values`, grouped `values.patches`, `write_formulas`, `write_number_formats`, `format_range`, `write_styles_many`, `write_data_validation`, `write_conditional_formatting`, `clear_style_dimensions`, grouped internal `write_values_many`/`write_number_formats_many`/`clear_many`/`clear_formats_many`/`clear_style_dimensions_many`/`autofit_many`, `clear_values`, `clear_values_raw`, `clear_formats`, `copy_range`, `move_range`.
- Structure: `insert_rows`, `delete_rows`, `insert_columns`, `delete_columns`, `reorder_range_columns`, `merge_range`, `unmerge_range`, `create_sheet`, `copy_sheet`, `rename_sheet`, `delete_sheet`, `hide_sheet`, `unhide_sheet`, `protect_sheet`, `unprotect_sheet`, `clear_sheet`, `set_sheet_tab_color`, `autofit`, `autofit_rows`.
- Tables: `append_table_rows`, `update_table_rows`, `create_table`, `resize_table`, `reorder_table_columns`, `clear_table_data`, `clear_table_filters`, `sort_table`, `filter_range`, `apply_table_view`, `set_table_total_row`, `set_table_style`, `copy_table_structure`.
- Cleaning: `normalize_headers`, `trim_whitespace`, `remove_duplicates`, `parse_dates`, `parse_numbers`, `standardize_currency`, `fill_missing_values`, `split_column`, `merge_columns`.

## Template, Formula, Style, Pivot, And Chart Workflows

- Template workflows: `copy_template_sheet` backed by internal `sheet.copy_clean_data_regions`, `register_template`, `unregister_template`, `clear_template_data_regions`, `fill_template_regions`, `validate_sheet_against_template`, `repair_sheet_from_template`, `create_template_report`.
- Formula workflows: `create_formula_sheet`, `read_formula_patterns`, `get_formula_dependency_graph`, `trace_formula_precedents`, `trace_formula_dependents`, `validate_formula_range`, `validate_formula_against_template`, `find_formula_errors`, `explain_formula`, `copy_formula_patterns`, `fill_formula_down`, `fill_formula_right`, `repair_formula_patterns`, `convert_formulas_to_values`, `recalculate_formulas`, `repair_formulas_from_template`.
- Style workflows: `read_style_summary`, `format_diagnostics`, `read_style_fingerprint`, `compare_style_fingerprint`, `get_theme`, `apply_theme`, `copy_style_from_template`, `repair_style_consistency`, `repair_style_from_template`.
- Pivot/chart workflows: `create_pivot_chart_summary` and related host-limited capability reports.

## Safety, Validation, And Lifecycle

- Safety artifacts: `create_snapshot`, `create_backup`, `list_snapshots`, `read_snapshot`, `compare_snapshots`, `refresh_snapshot`, `invalidate_snapshot`, `delete_snapshot`, `list_backups`, `read_backup`.
- Backup lifecycle: `verify_backup`, `pin_backup`, `unpin_backup`, `delete_backup`, `create_file_backup`, `restore_file_backup`, `prune_backups`, `restore_workbook_backup`.
- Validation: `validate_compact`, `validate_workbook`, `validate_sheet`, `validate_template_consistency`, `validate_formulas`, `validate_styles`, `validate_tables`, `validate_filters`, `validate_print_layout`, `validate_no_broken_references`, `validate_no_formula_errors`, `validate_no_unintended_changes`, `validate_table_against_template`.
- Workbook lifecycle: `calculate`, `save`, `close_workbook`, `export_local_config`, `import_local_config`, `embed_local_config`, `read_embedded_local_config`, `import_embedded_local_config`.
- Combined workflows: `preview_risky_edit`, `inspect_analyze`, `rollback_validate`.

## Mode Defaults

- Use `answer` for read-only discovery, schema, metadata, safety artifact inspection, and deterministic analysis.
- Use `preview_update` then `apply_update` for workbook mutations, backup lifecycle changes, restores, template fills, table writes, and structural edits.
- Use `validate` for explicit validation or post-change proof.
- Use `rollback` for recovery planning and confirmed rollback/restore flows.
