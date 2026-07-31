# Desktop fix v7.1.1

## Observed error

`RuntimeError: server.port does not work when global.developmentMode is true.`

## Root cause

The packaged Streamlit child process started with development mode
enabled while the desktop launcher also selected a free local port.
Streamlit rejects that configuration combination.

## Fix

The launcher now supplies both:

- `STREAMLIT_GLOBAL_DEVELOPMENT_MODE=false`
- `--global.developmentMode=false`

## Packaging improvement

The GitHub Actions workflow uploads the built application directory
directly. The downloaded GitHub artifact therefore has only one ZIP
extraction layer.
