# ahk-mcp

MCP server exposing AutoHotkey functionality, enabling model interfaces to automation tasks on Windows.

This server only works on Windows and provides the following tools to your AI:

- Ability to enumerate windows/applications
- Ability to control keyboard/mouse (typing, clicking, etc)
- Ability to interrogate Windows APIs (via AutoHotkey) about windows (e.g., to get the text of a window, the position of its GUI controls, etc.)
- Screen capture & OCR functionality (useful when text is not exposed properly by Windows APIs/controls)
- Ability to get accurate window positioning and contextual information about computer monitors (e.g., know what windows are on your primary/secondary monitors)
- Ability to manipulate windows and other actions via AutoHotkey

In total, there are 33 tools currently exposed by the server, the above is just a simple overview. 
While we work on documentation, exploring the source code in `main.py` is encouraged!

## Usage

This project makes use of the [Python MCP SDK](https://github.com/modelcontextprotocol/python-sdk) with FastMCP. Please 
see the Python MCP SDK repo and documentation for detailed information.

Assuming you've already setup `mcp` CLI, you can install this MCP service in Claude Desktop with a simple `mcp` command:

```bash
mcp install main.py
```

This project depends on `ahk-binary` to provide the required AutoHotkey executables and the [`ahk`](https://github.com/spyoungtech/ahk) 
project to interface with AutoHotkey. It uses `mss`, [`easyocr`](https://github.com/JaidedAI/EasyOCR?tab=readme-ov-file#installation), and `numpy`.

## Contributing

The best way to contribute is right here on GitHub. Please feel free to [open an issue](https://github.com/spyoungtech/ahk-mcp/issues) 
to get started. Pull requests are also welcome, but it is strongly recommended to open an issue first, especially for 
significant changes.