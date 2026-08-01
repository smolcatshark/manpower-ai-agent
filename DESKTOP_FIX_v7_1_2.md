# Desktop compatibility fix v7.1.2

Observed error:

`Failed to resolve Python.Runtime.Loader.Initialize`

The native pywebview window uses Python.NET on Windows. Some computers
cannot initialize the bundled Python.NET runtime.

The launcher now tries the native desktop window first. If it fails,
the app automatically opens in the default browser and displays a small
control window that keeps the local Streamlit server running. Closing
that control window stops the app.

This patch does not change PDF parsing, validation, location rules,
Excel export, or author branding.
