# docx

Tools for reading and writing docx files based on the python-docx library.

This supports the following optional environment variable:

- `ALLOWED_DIR`: The allowed directory. If supplied, the tools will only be able to read/write within this directory. To allow multiple directories, supply a strictly valid JSON-encoded list e.g. use double quotes: '["dir_a", "dir_b"]'
