# Screenshot MCP Server

A Model Context Protocol (MCP) server that enables AI assistants running in WSL to capture Windows screenshots with advanced features like monitor selection and window-specific capture.

## Features

- 📸 **Full Desktop Capture** - Capture all monitors (default behavior)
- 🖥️ **Monitor Selection** - Capture specific monitors (primary, 1, 2, etc.)
- 🪟 **Window Capture** - Capture specific windows by title match with DPI awareness
- 🚀 **Process Capture** - Capture windows by process name (e.g., notepad.exe)
- 📂 **Custom Save Locations** - Save to any folder using WSL or Windows paths
- 🔄 **Automatic Path Conversion** - Converts WSL paths to Windows paths
- 📁 **Organized Storage** - Screenshots saved to `workspace/screenshots/` by default
- 🎯 **DPI Aware** - Proper scaling for high-DPI displays
- 🖼️ **Full Window Capture** - Includes window shadows and borders without clipping

## Prerequisites

- Windows with WSL (Windows Subsystem for Linux)
- Node.js installed in WSL
- Claude Desktop or Claude Code with MCP support

## Installation

1. Clone this repository:
```bash
git clone https://github.com/peterparker57/screenshot-mcp.git
cd screenshot-mcp
```

2. Install dependencies:
```bash
npm install
# or with bun:
bun install
```

3. Add to your Claude configuration (`~/.claude.json` or `~/.claude/claude_desktop_config.json`):
```json
{
  "mcpServers": {
    "screenshot": {
      "command": "node",
      "args": [
        "/path/to/screenshot-mcp/index.js"
      ],
      "disabled": false,
      "alwaysAllow": [
        "take_screenshot"
      ]
    }
  }
}
```

4. Restart Claude Desktop/Code

## Usage

Once installed, you can ask Claude to take screenshots:

### Basic Usage
```
Take a screenshot
```

### Capture Specific Monitor
```
Take a screenshot of monitor 1
Take a screenshot of the primary monitor
```

### Capture Specific Window by Title
```
Take a screenshot of the "Chrome" window
Take a screenshot of window titled "Visual Studio Code"
```

### Capture Specific Window by Process Name
```
Take a screenshot of notepad.exe
Take a screenshot of the chrome process
Take a screenshot of process "Code"
```

### Custom Filename
```
Take a screenshot and save it as "test.png"
```

### Save to Custom Folder
```
Take a screenshot and save it to /mnt/c/Users/username/Pictures/
Take a screenshot and save to C:\Users\username\Desktop\
Take a screenshot of monitor 1 and save to folder "../docs/images"
```

## API Reference

The MCP server provides two tools:

### `list_windows`

**Description:** List all available windows with numbers for easy selection.

**Parameters:**
- `filter` (optional): Filter windows by title or process name
- `format` (optional): Output format
  - `"simple"` - Basic numbered list (default)
  - `"detailed"` - Include process IDs and window handles

**Returns:**
- Numbered list of windows with their titles and process names
- Total count of windows found

**Example:**
```
list_windows()
// Output:
// 1. Chrome - Google Search (Process: chrome)
// 2. Visual Studio Code (Process: Code)
// 3. screenshots (Process: explorer)
// Total windows: 3
```

### `take_screenshot`

**Parameters:**
- `filename` (optional): Name for the screenshot file (default: "screenshot.png")
- `monitor` (optional): Which monitor to capture
  - `"all"` - Capture all monitors (default)
  - `"primary"` - Capture primary monitor only
  - `1`, `2`, etc. - Capture specific monitor by index
- `windowTitle` (optional): Capture a specific window by its title (partial match supported)
- `processName` (optional): Capture a specific window by process name (e.g., "notepad.exe" or "notepad")
- `windowNumber` (optional): Capture a specific window by its number from list_windows output
- `folder` (optional): Custom folder path to save the screenshot
  - Supports WSL paths: `/mnt/c/Users/...`
  - Supports Windows paths: `C:\Users\...`
  - Supports relative paths: `../images`
  - Default: `workspace/screenshots/`

**Returns:**
- Success message with the file path
- Error message if capture fails

**Notes:** 
- `windowNumber` takes precedence over `windowTitle` and `processName`
- If both `windowTitle` and `processName` are provided, `windowTitle` takes precedence
- Custom folders are created automatically if they don't exist
- Path formats are automatically converted between WSL and Windows as needed

**Workflow Example:**
```
// First, list windows to see what's available
list_windows()

// Then capture by number
take_screenshot(windowNumber=2, filename="vscode.png")
```

## Technical Details

### Architecture
- **MCP Server**: Node.js with `@modelcontextprotocol/sdk`
- **Screenshot Capture**: PowerShell with .NET Windows Forms
- **Communication**: Executes PowerShell commands from WSL
- **Encoding**: Base64 encoding for reliable command execution

### Window Capture Features
- **DPI Awareness**: Automatically handles high-DPI displays for crisp captures
- **Window Padding**: Adds 10px padding to capture window shadows and borders
- **Render Wait**: Waits 200ms after focusing window to ensure complete rendering
- **Bounds Checking**: Prevents negative coordinates when windows are near screen edges
- **Process Matching**: Intelligent process name matching (strips .exe extension automatically)

### How It Works
1. MCP server receives screenshot request from Claude
2. Determines save location (custom folder or default)
3. Converts paths between WSL and Windows formats as needed
4. Constructs appropriate PowerShell script based on parameters
5. Encodes script in base64 to avoid escaping issues
6. Executes PowerShell command from WSL
7. PowerShell captures screenshot using Windows Forms APIs
8. Image is saved to the specified location

### Error Handling
- Filters PowerShell CLIXML output (verbose logging, not errors)
- Validates monitor indices
- Provides clear error messages for missing windows or processes
- Lists available windows when capture fails
- Automatically creates directories if needed
- Handles both WSL and Windows path formats

## Troubleshooting

### Screenshots folder not created
The server automatically creates folders as needed. Ensure you have write permissions to the target location.

### PowerShell execution errors
Check your PowerShell execution policy:
```powershell
Get-ExecutionPolicy
```

### Window not found errors
- Ensure the window is open and not minimized
- The title match is case-insensitive and supports partial matches
- Try using a more specific window title
- When searching by process, the tool will list all available windows to help you identify the correct one

### Window capture is clipped
The latest version includes automatic padding and DPI awareness. If you still experience clipping:
- Ensure you're using the latest version
- Try maximizing the window before capture
- Check if the window has unusual rendering (some apps use custom chrome)

### Path conversion issues
The server automatically converts between WSL and Windows path formats:
- WSL paths like `/mnt/c/...` are converted to `C:\...` for PowerShell
- Windows paths like `C:\...` are converted to `/mnt/c/...` for file verification
- Ensure your paths are accessible from both WSL and Windows

## Recent Updates

### v1.3.4 (Fork)
- **Better UX guidance** - Tool descriptions guide agents to display window lists properly
- **Pre-check for minimized windows** - Prevents screenshot attempts on minimized windows
- **Improved output formatting** - Adds helpful tips to window list output
- **Cleaner error messages** - Shows user-friendly errors instead of PowerShell commands

### v1.3.3 (Fork)
- **Improved error handling** - Clean error messages instead of raw PowerShell output
- **Better error extraction** - Properly parses thrown errors from PowerShell scripts
- **Graceful error display** - No more CLIXML garbage in error messages

### v1.3.2 (Fork)
- **Window state detection** - Shows [MINIMIZED] or [MAXIMIZED] indicators in window lists
- **Prevents capturing minimized windows** - Throws helpful error if trying to capture a minimized window
- **Enhanced window information** - Detailed view now includes window state (Normal/Minimized/Maximized)
- **Better error handling** - Clear messaging when attempting invalid captures

### v1.3.1 (Fork)
- **Added list_windows tool** - List all available windows with numbers for easy selection
- **Window selection by number** - Use `windowNumber` parameter to capture windows by their list number
- **Improved discovery workflow** - No more error-based window discovery
- **Filter support** - Filter windows by title or process name in list_windows
- **Detailed view option** - Get process IDs and window handles with format="detailed"

### v1.3.0 (Fork)
- **Improved window enumeration** - Now uses Windows API EnumWindows instead of Get-Process
- **Better Explorer window detection** - Can now find File Explorer and other system windows
- **More comprehensive window listing** - Shows all visible windows with titles
- **Maintains all existing features** - Fully backward compatible

### v1.2.0
- Added custom folder support with the `folder` parameter
- Supports both WSL paths (`/mnt/...`) and Windows paths (`C:\...`)
- Automatic path conversion between WSL and Windows formats
- Creates custom directories automatically if they don't exist
- Maintains backward compatibility with default screenshots folder

### v1.1.0
- Added process name capture support
- Can now capture windows by process name (e.g., "notepad.exe")
- Intelligent .exe extension handling
- Enhanced error messages showing available windows

### v1.0.1
- Fixed window capture clipping issues
- Added DPI awareness for high-DPI displays
- Added padding to capture window shadows
- Improved window rendering wait time
- Added bounds checking for edge cases

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

MIT License - see LICENSE file for details

## Acknowledgments

- Built for use with [Claude Desktop](https://claude.ai) and Claude Code
- Uses the [Model Context Protocol](https://modelcontextprotocol.io)
- Special thanks to the WSL team for making Windows/Linux integration possible
