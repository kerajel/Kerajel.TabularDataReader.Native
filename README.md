# Kerajel.TabularDataReader.Native

Kerajel.TabularDataReader.Native is a high-performance Rust library that rapidly converts spreadsheets (XLS, XLSX, etc.) to CSV format. It leverages the calamine and csv crates to provide a C-compatible API with functions for converting data from raw bytes or file paths. The API returns an `OperationResult` containing either the CSV output or an error message, and memory allocated for these results must be freed using the provided cleanup function.

Note: The `OperationResult` model is derived from [Kerajel.Primitives](https://github.com/kerajel/Kerajel.Primitives).

For build instructions, please refer to the `build` file.

MIT Licensed.
