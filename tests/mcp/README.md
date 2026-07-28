# MCP conformance tests

`test_conformance.py` launches the installed server over stdio through the
official `mcp` client SDK. It verifies initialization instructions, exact tool
count and annotations, schemas, all happy and error paths, progress, response
caps, cursor round-trips and write invalidation, and `EXCEL_LSP_ROOT` denial.
