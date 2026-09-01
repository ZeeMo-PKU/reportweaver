# ReportWeaver

ReportWeaver is a Windows utility for organizing Word technical reports. It generates and formats the main table of contents, list of figures, and list of tables while normalizing body heading styles and preserving an unnumbered introduction entry.

The tool always works on a copy of the selected document. It writes the processed file next to the source and never overwrites the original.

## Features

- Generate or refresh the main table of contents.
- Generate lists of figures and tables from Word captions.
- Normalize first-, second-, and third-level body headings.
- Convert a numbered introduction into an unnumbered TOC entry.
- Exclude introduction subheadings from the main TOC.
- Apply configurable fonts, sizes, bold settings, line spacing, and page breaks.
- Preserve the source document and add a timestamp when an output name already exists.

## Requirements

- Windows
- Desktop Microsoft Word
- Python 3
- Internet access for first-time dependency installation

## Quick Start

1. Download or clone this repository.
2. Double-click `Install-Dependencies.bat` before the first use.
3. Double-click `Run-WordReportTool.bat`.
4. Select the `.docx` report in the file picker.
5. Wait for the processing window to report completion.

The generated file is written next to the original document:

```text
report.with-directories.docx
```

## Expected Document Structure

The front matter should preferably contain standalone paragraphs for the table of contents, list of figures, list of tables, and introduction. For Chinese technical-report templates, these are commonly written as:

```text
目录
插图清单
附表清单
引言
```

Body headings should follow a numbered hierarchy:

```text
引言
1 Chapter title
1.1 Section title
1.1.1 Subsection title
```

Legacy documents that use `1 引言`, introduction subheadings such as `1.1`, and `2` for the first regular chapter are also supported. The tool converts the introduction into an unnumbered entry, removes its internal subheadings from the main TOC, and shifts subsequent chapter numbers accordingly.

Figure and table captions should follow patterns such as:

```text
图 1 Example figure caption
表 1 Example table caption
```

## Configuration

Edit `config\format-settings.json` to adapt the tool to another report template. Available settings include:

- Front-matter title text
- Page breaks before figure lists, table lists, and body content
- Fonts and sizes for titles, TOC entries, headings, captions, and list entries
- Bold behavior
- Default line spacing

Example title configuration:

```json
{
  "titles": {
    "toc": "目录",
    "figures": "插图清单",
    "tables": "附表清单",
    "introduction": "引言"
  }
}
```

Example page-break configuration:

```json
{
  "page_breaks": {
    "before_figures": true,
    "before_tables": true,
    "before_body": true
  }
}
```

Formatting entries use the following shape:

```json
{
  "font": "Times New Roman",
  "size": 12,
  "bold": null
}
```

`bold` accepts `true`, `false`, or `null`; `null` preserves the existing bold state.

## Command-Line Usage

Run with the default configuration:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\tools\Update-ReportDirectories.ps1" -Path "C:\path\to\report.docx"
```

Run with a custom configuration file:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File ".\tools\Update-ReportDirectories.ps1" -Path "C:\path\to\report.docx" -ConfigPath ".\config\format-settings.json"
```

## Troubleshooting

Close the target Word document before processing it. If the tool fails, inspect `WordReportTool-run-log.txt` on the desktop.

If Python is not found, install Python 3 with `Add python.exe to PATH` enabled and rerun `Install-Dependencies.bat`. If Word COM automation is unavailable, verify that the desktop version of Microsoft Word is installed.
