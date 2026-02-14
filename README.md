# Marp-fast PPT Maker MCP Server

This is a powerful Model Context Protocol (MCP) Server that converts any Markdown text into beautifully formatted and highly precise PPTX and PDF presentations with a single click.

Its core innovation is the **Two-Pass Physical Probe Measurement Engine**. By abandoning traditional character-count estimation, it utilizes a headless browser to perform real DOM pixel-level measurements. This completely resolves pain points such as slide content overflow, truncated math formulas, and broken layouts.

## Core Features

- **Physical-Level Anti-Overflow**: Uses the Chromium engine to measure true rendering height, 100% preventing text from overflowing slide boundaries.
- **Semantic-Aware Splitting**: Prioritizes semantic page breaks at higher-level headings, maintaining the logical coherence of human reading.
- **Cell-Level Element Protection**: Automatically repairs Markdown tables (auto-completing headers) and list structures (auto-reverting indentation and injecting parent context) across pages, preventing LaTeX math rendering failures.
- **Multi-Theme Support**: Built-in `default`, `gaia`, and `uncover` themes, with support for mounting a local `themes` folder to extend community themes (e.g., `academic`, `rose-pine`).
- **Fully Cross-Platform**: Perfectly compatible with Windows, macOS, and Linux.

## Prerequisites

1. **Python 3.10+**
2. **Node.js** (for running Marp CLI)
3. **Google Chrome or Microsoft Edge** (Locally installed; the program will automatically detect the path, eliminating the need for manual downloads)

## Installation Guide

### 1. Install Python Dependencies
It is recommended to run this in a virtual environment:
```bash
python -m venv venv
# Windows activation: venv\Scripts\activate
# macOS/Linux activation: source venv/bin/activate
pip install -r requirements.txt
```

(Note: Because this program automatically calls the system-installed Chrome/Edge, you do not need to manually run playwright install chromium)

### 2. Install Marp CLI (Node.js)
Install the local Marp dependency in the project root directory:

```Bash
npm install
```

## Running and Usage
Debug and run via the MCP Inspector:

```Bash
npx @modelcontextprotocol/inspector ./venv/bin/python -u server.py
```

Alternatively, configure this Server in the config file of an MCP-supported client (like Claude Desktop):

```JSON
{
  "mcpServers": {
    "Marp-PPT-Maker": {
      "command": "/absolute/path/to/venv/bin/python",
      "args": ["-u", "/absolute/path/to/server.py"]
    }
  }
}
```
## Output Artifacts
The generated .md intermediate files, .pptx, and .pdf final files will automatically be saved in the output_slides folder located in the project root directory.

## Related links

- [MCP](https://modelcontextprotocol.io/)
- [Marp](https://marp.app/)
- [Claude](https://docs.anthropic.com/)

## License
MIT
