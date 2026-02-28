# Excel MCP Server

An MCP (Model Context Protocol) server that acts as a retrieval layer for
Excel workbooks.  It lets LLM-based agents explore, query, summarize, and
analyze `.xlsx` / `.xlsm` / `.xltx` files ranging from **1 KB to 100 MB**.

## Features

| Tool | Purpose |
|---|---|
| `list_excel_files` | Discover Excel files in a directory |
| `get_workbook_info` | File metadata, sheet list, and dimensions |
| `get_sheet_names` | Quick list of all sheet names |
| `get_sheet_summary` | **Brief** (headers, row/col counts) or **detailed** (types, sample rows, statistics, formulas with human-readable explanations) |
| `read_sheet_data` | Paginated row retrieval with optional column filter |
| `get_cell_info` | Value, type, formula + explanation for a single cell |
| `get_range_data` | Read an arbitrary cell range (e.g. `A1:D10`) |
| `get_formulas` | Every formula in a sheet with plain-English explanations |
| `search_in_sheet` | Substring search across all cells |
| `get_sheet_statistics` | Descriptive stats (count, sum, mean, median, min, max) for numeric columns |
| `validate_sheet_data` | Data-quality checks (mixed types, missing values, duplicates) |

### Formula Explanations

Formulas are translated into natural language using column headers and row
numbers instead of raw cell references:

```
=SUM(B2:B4)  →  "This formula adds up 'Revenue' (rows 2 to 4)."
=B2-C2       →  "In row 2: This formula computes: 'Revenue' (row 2) - 'Cost' (row 2)."
```

### Large-File Handling

Files larger than **5 MB** are automatically opened in read-only streaming
mode to keep memory usage low.  Data retrieval is paginated (default 100
rows per page) so even 100 MB files can be queried efficiently.

## Quick Start

```bash
# Install
pip install -r requirements.txt

# Run the server (stdio transport)
python -m excel_mcp_server.server
```

### MCP Client Configuration

Add the following to your MCP client config (e.g. Claude Desktop):

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
# Install dev dependencies
pip install -r requirements.txt

# Run tests
python -m pytest tests/ -v
```

## Project Structure

```
excel_mcp_server/
├── __init__.py            # Package marker
├── server.py              # MCP tool definitions (entry point)
├── excel_reader.py        # Core reading, analysis & data extraction
└── formula_explainer.py   # Formula → plain-English translator
tests/
├── conftest.py            # Shared fixtures (sample workbooks)
├── test_excel_reader.py   # Tests for every reader function
└── test_formula_explainer.py  # Tests for formula explanations
```