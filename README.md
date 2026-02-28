# Excel MCP Server

An MCP (Model Context Protocol) server that acts as a retrieval layer for
Excel workbooks.  It lets LLM-based agents explore, query, summarize, and
analyze `.xlsx` / `.xlsm` / `.xltx` files ranging from **1 KB to 100 MB**.

## Architecture

Each retriever is a **separate MCP tool** with Anthropic-style documentation
(`Annotated[Field]`), so the LLM can discover and choose the right one.
All retrievers share a consistent interface:

- **`sheet_name`** — pass `null` to operate on **all sheets**, or a name for one sheet
- **`content_type`** — `"formulas"` (default, recommended), `"values"`, or `"both"`
- **`output_format`** — `"markdown"`, `"json"`, or `"csv"`
- **`header_row`** — `0` for auto-detect (handles unstructured sheets), or explicit row

## Tools

| Tool | Purpose |
|---|---|
| `discover_excel_files` | Find Excel files in a directory |
| `inspect_workbook` | File metadata, sheet list, dimensions |
| `list_sheets` | Quick list of sheet names |
| `retrieve_sheet_data` | Paginated row retrieval with column filter |
| `retrieve_formulas` | Every formula with plain-English explanations |
| `retrieve_summary` | Brief or detailed sheet summaries |
| `retrieve_cell_info` | Deep-dive into a single cell |
| `retrieve_range` | Read an arbitrary cell range |
| `search_workbook` | Search values/formulas across all sheets |
| `retrieve_statistics` | Descriptive stats for numeric columns |
| `validate_data` | Data-quality checks (mixed types, missing values, duplicates) |

### Formula Explanations

```
=SUM(B2:B4)  →  "This formula adds up 'Revenue' (rows 2 to 4)."
=B2-C2       →  "In row 2: This formula computes: 'Revenue' (row 2) - 'Cost' (row 2)."
```

### Unstructured Sheet Support

Sheets where headers are not in row 1 are handled via auto-detection
(`header_row=0`).  The system scans the first 20 rows to find the most
likely header row.

## Quick Start

```bash
pip install -r requirements.txt
python -m excel_mcp_server.server     # stdio transport
```

### MCP Client Configuration

```json
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": ["-m", "excel_mcp_server.server"]
    }
  }
}
```

## Development

```bash
pip install -r requirements.txt
python -m pytest tests/ -v            # 140 tests
```

## Project Structure

```
excel_mcp_server/
├── server.py                  # MCP tool definitions (entry point)
├── formula_explainer.py       # Formula → plain-English translator
└── excel_retrievers/          # All retrieval logic
    ├── _base.py               # Shared utilities, discovery functions
    ├── _formatters.py         # JSON / Markdown / CSV formatters
    ├── sheet_data.py          # Paginated data retriever
    ├── formula.py             # Formula extraction retriever
    ├── search.py              # Cross-sheet search retriever
    ├── summary.py             # Summary retriever
    ├── cell_info.py           # Single cell retriever
    ├── range_data.py          # Range retriever
    ├── statistics.py          # Statistics retriever
    └── validation.py          # Validation retriever
tests/
├── conftest.py                # Fixtures (5 workbook types)
├── test_retrievers.py         # Retriever unit tests
├── test_formatters.py         # Formatter tests
├── test_formula_explainer.py  # Formula explainer tests
└── test_e2e.py                # End-to-end MCP tool tests
```