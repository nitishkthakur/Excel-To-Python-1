# Copilot Instructions — Excel MCP Server

## Project Overview

This repository implements an **MCP (Model Context Protocol) server** that
acts as a retrieval layer for Excel workbooks (`.xlsx`, `.xlsm`, `.xltx`,
1 KB – 100 MB).  An LLM-based agent uses the MCP tools to explore, query,
summarize, and deeply analyze Excel files — from quick overviews to full
formula lineage.

---

## Architecture

### Design Decision: Separate Tools (not a dispatcher)

Each retriever is exposed as its **own MCP tool** rather than using a single
dispatcher function with a retriever-name argument.  This follows Anthropic's
tool-use guidelines: clear, focused descriptions help the model pick the right
tool without extra routing logic.

### Common Parameters

Every retriever tool accepts these parameters:

| Parameter       | Default       | Description |
|-----------------|---------------|-------------|
| `file_path`     | *(required)*  | Absolute path to the Excel file |
| `sheet_name`    | `null`        | Sheet name, or `null` for ALL sheets |
| `content_type`  | `"formulas"`  | `"formulas"` / `"values"` / `"both"` |
| `output_format` | `"markdown"`  | `"markdown"` / `"json"` / `"csv"` |
| `header_row`    | `0`           | Header row number (0 = auto-detect) |

**content_type** defaults to `"formulas"` — tool docs tell the LLM to fetch
formulas first to understand logic, and only request values when needed.

---

## Folder Structure

```
excel_mcp_server/
├── __init__.py                     # Package root, version
├── server.py                       # MCP tool definitions (Annotated[Field])
├── formula_explainer.py            # Formula → plain-English translator
│
└── excel_retrievers/               # ← all retrieval logic lives here
    ├── __init__.py                 # Registry (RETRIEVERS dict), exports
    ├── _base.py                    # Shared utilities:
    │                               #   file validation, workbook I/O,
    │                               #   sheet resolution, header detection,
    │                               #   serialization, discovery functions,
    │                               #   statistics, formula extraction
    ├── _formatters.py              # JSON / Markdown / CSV formatters
    │
    ├── sheet_data.py               # Paginated row retriever
    ├── formula.py                  # Formula extraction retriever
    ├── search.py                   # Cross-sheet search retriever
    ├── summary.py                  # Brief / detailed summary retriever
    ├── cell_info.py                # Single cell deep-dive retriever
    ├── range_data.py               # Arbitrary range retriever
    ├── statistics.py               # Numeric statistics retriever
    └── validation.py               # Data-quality checks retriever

tests/
├── conftest.py                     # Shared fixtures (5 workbook types)
├── test_retrievers.py              # Per-retriever unit tests
├── test_formatters.py              # Formatter unit tests
├── test_formula_explainer.py       # Formula explainer tests
└── test_e2e.py                     # End-to-end tests through MCP tools
```

---

## How Things Tie Together

1. **`server.py`** defines MCP tools with `@mcp.tool()` and Anthropic-style
   `Annotated[Field]` documentation.  Each tool calls a retriever's
   `retrieve()` function and formats the result via `format_output()`.

2. **`excel_retrievers/`** is the retriever package.  Every retriever
   module exposes a `retrieve(file_path, *, ...)` function that returns a
   plain `dict`.  The `RETRIEVERS` dict in `__init__.py` maps names to
   functions for programmatic access.

3. **`_base.py`** provides shared utilities so retrievers stay DRY:
   - `validate_file()`, `open_workbook()` — file I/O with size checks
   - `resolve_sheets()` — returns all sheets or one sheet
   - `get_headers(header_row=0)` + `detect_header_row()` — handles
     unstructured sheets where headers aren't in row 1
   - `extract_formulas()` — extracts formulas with explanations
   - `add_formula_explanations()` — appends formula data to results
   - `package_results()` — wraps single vs multi-sheet results

4. **`_formatters.py`** converts `dict` results into `str` output:
   - JSON: `json.dumps` with pretty-print
   - Markdown: tables for list-of-dicts, sections for multi-sheet
   - CSV: tabular data as CSV, falls back to JSON for non-tabular

5. **`formula_explainer.py`** translates formulas into natural language
   using column headers instead of cell references, e.g.
   `=SUM(B2:B4)` → "adds up 'Revenue' rows 2 to 4".

---

## Retriever Contract

Every retriever follows this pattern:

```python
def retrieve(
    file_path: str,
    *,
    sheet_name: str | None = None,    # None = all sheets
    content_type: str = "formulas",   # "values" | "formulas" | "both"
    header_row: int = 0,              # 0 = auto-detect
    # ... retriever-specific params
) -> dict[str, Any]:
    ...
```

Returns a `dict` with:
- Single sheet: flat dict with `sheet_name`, `content_type`, and data
- Multi-sheet: `{"sheet_count": N, "sheets": [...], "content_type": ...}`

---

## Adding a New Retriever

1. Create `excel_retrievers/new_retriever.py` with a `retrieve()` function
   following the contract above.
2. Import it in `excel_retrievers/__init__.py` and add to `RETRIEVERS`.
3. Add a new `@mcp.tool()` in `server.py` with Anthropic-style docs.
4. Add tests in `tests/test_retrievers.py`.

---

## Common Use Cases

| Use Case | Tool(s) |
|----------|---------|
| Discover files | `discover_excel_files` → `inspect_workbook` |
| Quick overview | `retrieve_summary(detail_level="brief")` |
| Deep report | `retrieve_summary(detail_level="detailed")` per sheet |
| Formula lineage | `retrieve_formulas(sheet_name=null)` for all sheets |
| Trace a variable | `search_workbook(query="variable_name", content_type="formulas")` |
| Find inputs | `validate_data` + `retrieve_formulas` to separate inputs from computed cells |
| Executive summary | `retrieve_summary(detail_level="detailed")` all sheets, markdown |

---

## Build & Test

```bash
pip install -r requirements.txt     # install deps
python -m pytest tests/ -v          # run all 140 tests
python -m excel_mcp_server.server   # start MCP server (stdio)
```

---

## Dependencies

- `mcp>=1.0.0` — Model Context Protocol SDK
- `openpyxl>=3.1.0` — Excel file reader
- `pydantic>=2.0` — used by MCP for tool parameter validation
