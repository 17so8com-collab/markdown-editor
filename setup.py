from cx_Freeze import setup, Executable

build_exe_options = {
    "excludes": [],
    "includes": [
        "webview",
        "webview.window",
        "webview.menu",
        "bottle",
        "markdown",
        "markdown.extensions",
        "markdown.extensions.tables",
        "markdown.extensions.fenced_code",
        "markdown.extensions.codehilite",
        "markdown.extensions.toc",
        "markdown_it",
        "markdown_it.common",
        "markdown_it.rules_core",
        "markdown_it.token",
        "markdown_it.node",
        "markdown_it.renderer",
        "markdown_it.parser",
        "markdown_it.presets",
        "mdurl",
        "pygments",
        "docx",
        "docx.api",
        "docx.document",
        "docx.oxml",
        "docx.oxml.text",
        "docx.oxml.table",
        "docx.oxml.styles",
        "docx.oxml.ns",
        "docx.oxml.parse",
        "lxml",
        "lxml.etree",
    ],
    "include_files": [],
    "zip_include_packages": [],
    "optimize": 0,
}

build_options = {
    "build_exe": build_exe_options,
}

setup(
    name="MarkdownEditor",
    version="1.0",
    description="Markdown编辑器 - 支持MD/Word/HTML/PDF互转",
    options=build_options,
    executables=[
        Executable(
            "app.py",
            base="Win32GUI",
            target_name="MarkdownEditor.exe",
            icon=None,
        )
    ],
)
