# Stock Statement AI Compiler — New Master Format

This build uses the uploaded **Master Secondary & Closing file(1).xlsx** format as the single source of truth.

- MAIN FILE structure is preserved for output.
- HQ + CUSTOMER NAME + SKU NAME determine the target output row.
- BRAND is retained as the SKU grouping field.
- SEC UNITS / SEC VALUE / CLO UNITS / CLO VALUE are populated in the existing MAIN FILE columns.
- All other master sheets, rows, columns, filters, widths and workbook structure are retained.
- Existing SKU approval, Ignore, duplicate handling, statement removal and PDF pack-format handling remain enabled.
