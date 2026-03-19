# Markdown Editor

A desktop Markdown editor built with Python + pywebview + Vditor.

## Features

- **Three editing modes**: WYSIWYG (rich text), Instant Rendering (Typora-like), Split View
- **Import**: Markdown, Word (.docx), HTML
- **Export**: Markdown, HTML, Word (.docx), PDF (via browser print)
- **Dark mode** support
- **Word / character / line count** in status bar

## Tech Stack

- **pywebview** — Desktop webview window
- **Vditor 3.11.0** — Markdown editor (CDN)
- **tkinter** — Native file dialogs
- **python-docx** — Word document handling
- **markdown** — Python Markdown renderer

## Requirements

```
pip install pywebview markdown python-docx
```

## Run

```bash
python app.py
```

## Build Executable

```bash
pip install pyinstaller
pyinstaller --onefile app.py
```

The executable will be at `dist/app.exe`.

## Keyboard Shortcuts (Vditor)

| Shortcut | Action |
|----------|--------|
| Alt+Cmd+7 | WYSIWYG mode |
| Alt+Cmd+8 | Instant Rendering mode |
| Alt+Cmd+9 | Split View mode |
