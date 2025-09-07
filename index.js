#!/usr/bin/env node

import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { CallToolRequestSchema, ListToolsRequestSchema } from '@modelcontextprotocol/sdk/types.js';
import { exec } from 'child_process';
import { promisify } from 'util';
import fs from 'fs/promises';
import path from 'path';

const execAsync = promisify(exec);

const server = new Server(
  {
    name: 'screenshot-server',
    version: '1.3.0-triage',
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

// This triage build merges v1.2.0 baseline with advanced features
// observed in prior sessions: list_windows, windowNumber/windowHandle,
// allowFocus/restoreIfMinimized, and CopyFromScreen-based window capture.

server.setRequestHandler(ListToolsRequestSchema, async () => {
  return {
    tools: [
      {
        name: 'take_screenshot',
        description: 'Take a screenshot of all monitors, specific monitor, or a specific window',
        inputSchema: {
          type: 'object',
          properties: {
            filename: {
              type: 'string',
              description: 'Filename for the screenshot (default: screenshot.png)',
              default: 'screenshot.png'
            },
            monitor: {
              type: ['string', 'number'],
              description: 'Which monitor to capture: "all" (default), "primary", or monitor number (1, 2, etc.)',
              default: 'all'
            },
            windowTitle: {
              type: 'string',
              description: 'Capture a specific window by its title (partial match supported)'
            },
            processName: {
              type: 'string',
              description: 'Capture a specific window by process name (e.g., "notepad.exe" or just "notepad")'
            },
            windowNumber: {
              type: 'number',
              description: 'Capture a specific window by its number from list_windows output (1-indexed)'
            },
            windowHandle: {
              type: 'string',
              description: 'Capture a specific window by exact handle (hex like 0x00000000 or decimal). Use list_windows format: "detailed" to obtain.'
            },
            filter: {
              type: 'string',
              description: 'Optional filter to narrow window selection (title or process contains, case-insensitive)'
            },
            allowFocus: {
              type: 'boolean',
              description: 'If true, briefly focuses the target window before capture to avoid black frames for some GPU apps.',
              default: false
            },
            restoreIfMinimized: {
              type: 'boolean',
              description: 'If true, restores the window if minimized, captures, then returns focus to prior window.',
              default: false
            },
            folder: {
              type: 'string',
              description: 'Custom folder path to save the screenshot (supports both WSL and Windows paths). Defaults to workspace/screenshots/'
            }
          }
        }
      },
      {
        name: 'list_windows',
        description: 'List available windows with numbers for targeting. Use filter to narrow results; set format to "detailed" for handles.',
        inputSchema: {
          type: 'object',
          properties: {
            filter: { type: 'string', description: 'Filter windows by title or process name (optional)' },
            format: { type: 'string', description: 'Output format: "simple" (default) or "detailed"', default: 'simple' }
          }
        }
      }
    ]
  };
});

function toWindowsPath(wslPath) {
  return wslPath.replace(/^\/mnt\/([a-z])\//, '$1:\\').replace(/\//g, '\\');
}

function toWindowsUNC(wslPath) {
  const distro = process.env.WSL_DISTRO_NAME || '';
  const rel = wslPath.replace(/^\//, '');
  // Build \\wsl$\<distro>\... UNC path
  const unc = distro ? `\\\\wsl$\\${distro}\\${rel.replace(/\//g, '\\')}` : wslPath;
  return unc;
}

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  if (request.params.name === 'list_windows') {
    const filter = request.params.arguments?.filter || '';
    const format = request.params.arguments?.format || 'simple';

    try {
      const escPS = (s) => String(s ?? '').replace(/'/g, "''");
      const psFilter = `'${escPS(filter)}'`;
      const psScript = `
        $ProgressPreference = 'SilentlyContinue'
        Add-Type @"
          using System;
          using System.Text;
          using System.Runtime.InteropServices;
          using System.Collections.Generic;
          public class Win32 {
            public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
            [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);
            [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
            [DllImport("user32.dll")] public static extern int GetWindowTextLength(IntPtr hWnd);
            [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);
            [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
            [DllImport("user32.dll")] public static extern bool IsZoomed(IntPtr hWnd);
            [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
            [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, int nFlags);
          }
"@
        $results = New-Object System.Collections.ArrayList
        $filter = ${psFilter}.ToLower()
        [Win32]::EnumWindows({ param($hWnd, $lParam)
          if ([Win32]::IsWindowVisible($hWnd)) {
            $len = [Win32]::GetWindowTextLength($hWnd)
            if ($len -gt 0) {
              $sb = New-Object System.Text.StringBuilder ($len + 1)
              [void][Win32]::GetWindowText($hWnd, $sb, $sb.Capacity)
              $title = $sb.ToString()
              if ($title.Length -gt 0) {
                $procId = 0; [void][Win32]::GetWindowThreadProcessId($hWnd, [ref]$procId)
                try { $p = Get-Process -Id $procId -ErrorAction Stop } catch { $p = $null }
                $proc = if ($p) { $p.ProcessName } else { '' }
                $state = if ([Win32]::IsIconic($hWnd)) { 'Minimized' } elseif ([Win32]::IsZoomed($hWnd)) { 'Maximized' } else { 'Normal' }
                $handleHex = ('0x{0:X8}' -f $hWnd.ToInt64())
                $line = [PSCustomObject]@{ Title=$title; Process=$proc; PID=$procId; Handle=$handleHex; State=$state }
                if ([string]::IsNullOrEmpty($filter) -or $title.ToLower().Contains($filter) -or $proc.ToLower().Contains($filter)) {
                  [void]$results.Add($line)
                }
              }
            }
          }
          return $true
        }, [IntPtr]::Zero) | Out-Null
        for ($i=0; $i -lt $results.Count; $i++) {
          $w = $results[$i]
          if ('${format}' -eq 'detailed') {
            Write-Host ("$($i+1). $($w.Title) (Process: $($w.Process), PID: $($w.PID), Handle: $($w.Handle), State: $($w.State))")
          } else {
            Write-Host ("$($i+1). $($w.Title) (Process: $($w.Process))")
          }
        }
      `;
      const encoded = Buffer.from(psScript, 'utf16le').toString('base64');
      const { stdout } = await execAsync(`powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -OutputFormat Text -EncodedCommand ${encoded}`);
      const output = stdout?.trim() || 'No windows found';
      return { content: [{ type: 'text', text: output }] };
    } catch (error) {
      const parts = [];
      parts.push('Failed to list windows:');
      if (error && error.message) parts.push(String(error.message));
      if (error && error.stderr) parts.push('stderr:', String(error.stderr).trim());
      if (error && error.stdout) parts.push('stdout:', String(error.stdout).trim());
      return { content: [{ type: 'text', text: parts.join(' ') }], isError: true };
    }
  }

  if (request.params.name === 'take_screenshot') {
    const filename = request.params.arguments?.filename || 'screenshot.png';
    const monitor = request.params.arguments?.monitor || 'all';
    const windowTitle = request.params.arguments?.windowTitle;
    const processName = request.params.arguments?.processName;
    const windowNumber = request.params.arguments?.windowNumber;
    const windowHandleArg = request.params.arguments?.windowHandle;
    const filter = request.params.arguments?.filter;
    const allowFocus = request.params.arguments?.allowFocus || false;
    const restoreIfMinimized = request.params.arguments?.restoreIfMinimized || false;
    const customFolder = request.params.arguments?.folder;

    // Determine target folder and Windows output path
    let screenshotsDir;
    let windowsPath;
    if (customFolder) {
      if (customFolder.match(/^[A-Za-z]:\\\\/)) {
        windowsPath = path.join(customFolder, filename).replace(/\\/g, '\\');
        const driveLetter = customFolder[0].toLowerCase();
        screenshotsDir = customFolder.replace(/^[A-Za-z]:/, `/mnt/${driveLetter}`).replace(/\\/g, '/');
      } else if (customFolder.startsWith('/mnt/')) {
        screenshotsDir = customFolder;
        const windowsFolder = toWindowsPath(customFolder);
        windowsPath = windowsFolder + '\\' + filename;
      } else {
        screenshotsDir = path.resolve(customFolder);
        windowsPath = screenshotsDir.startsWith('/mnt/')
          ? path.join(toWindowsPath(screenshotsDir), filename)
          : path.join(toWindowsUNC(screenshotsDir), filename);
      }
    } else {
      screenshotsDir = path.resolve(process.cwd(), 'screenshots');
      const outputPath = path.join(screenshotsDir, filename);
      windowsPath = outputPath.startsWith('/mnt/') ? toWindowsPath(outputPath) : toWindowsUNC(outputPath);
    }

    await fs.mkdir(screenshotsDir, { recursive: true });

    try {
      let psScript;

      if (windowTitle || processName || windowNumber || windowHandleArg) {
        // Window-specific capture. Insight: prefer CopyFromScreen after optional focus to avoid black frames.
        const escPS = (s) => String(s ?? '').replace(/'/g, "''");
        const psFilter = `'${escPS(filter)}'`;
        const psWindowTitle = `'${escPS(windowTitle)}'`;
        const psProcessName = `'${escPS(processName)}'`;
        const psAllowFocus = allowFocus ? "'$true'" : "'$false'";
        const psRestore = restoreIfMinimized ? "'$true'" : "'$false'";
        psScript = `
          $ProgressPreference = 'SilentlyContinue'
          Add-Type -AssemblyName System.Windows.Forms
          Add-Type -AssemblyName System.Drawing
          Add-Type @"
            using System;
            using System.Text;
            using System.Runtime.InteropServices;
            public class Win32 {
              public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
              [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);
              [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
              [DllImport("user32.dll")] public static extern int GetWindowTextLength(IntPtr hWnd);
              [DllImport("user32.dll")] public static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);
              [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
              [DllImport("user32.dll")] public static extern bool IsZoomed(IntPtr hWnd);
              [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
              [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
              [DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
              [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
              [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
              [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, int nFlags);
              public struct RECT { public int Left; public int Top; public int Right; public int Bottom; }
            }
"@
          Add-Type @"
            using System.Runtime.InteropServices;
            public class DPI { [DllImport("shcore.dll")] public static extern int SetProcessDpiAwareness(int value); }
"@
          [DPI]::SetProcessDpiAwareness(2)

          # Injected variables
          $psAllowFocus = ${psAllowFocus}
          $psRestore = ${psRestore}
          $psFilter = ${psFilter}
          $psWindowTitle = ${psWindowTitle}
          $psProcessName = ${psProcessName}

          function Get-WindowsList([string]$flt) {
            $arr = @()
            [Win32]::EnumWindows({ param($hWnd, $lParam)
              if ([Win32]::IsWindowVisible($hWnd)) {
                $len = [Win32]::GetWindowTextLength($hWnd)
                if ($len -gt 0) {
                  $sb = New-Object System.Text.StringBuilder ($len + 1)
                  [void][Win32]::GetWindowText($hWnd, $sb, $sb.Capacity)
                  $title = $sb.ToString()
                  if ($title.Length -gt 0) {
                    $procId = 0; [void][Win32]::GetWindowThreadProcessId($hWnd, [ref]$procId)
                    try { $p = Get-Process -Id $procId -ErrorAction Stop } catch { $p = $null }
                    $proc = if ($p) { $p.ProcessName } else { '' }
                    if ([string]::IsNullOrEmpty($flt) -or $title.ToLower().Contains($flt.ToLower()) -or $proc.ToLower().Contains($flt.ToLower())) {
                      $obj = [PSCustomObject]@{ Handle=$hWnd; Title=$title; PID=$procId; Process=$proc }
                      $arr += $obj
                    }
                  }
                }
              }
              return $true
            }, [IntPtr]::Zero) | Out-Null
            return ,$arr
          }

          $filter = $psFilter
          $target = [IntPtr]::Zero

          if ('${windowHandleArg}' -ne '' -and '${windowHandleArg}' -ne 'undefined') {
            $h = '${windowHandleArg}'
            if ($h -like '0x*') { $hVal = [Convert]::ToInt64($h, 16) } else { $hVal = [int64]$h }
            $target = [IntPtr]$hVal
          } elseif ('${windowNumber}' -ne '' -and '${windowNumber}' -ne 'undefined') {
            $list = Get-WindowsList $filter
            $idx = [int]'${windowNumber}' - 1
            if ($idx -lt 0 -or $idx -ge $list.Count) { throw "Invalid windowNumber: ${windowNumber}. Available: 1..$($list.Count)" }
            $target = $list[$idx].Handle
          } elseif ('${windowTitle}' -ne '') {
            $list = Get-WindowsList $filter
            $match = $list | Where-Object { $_.Title -like "*$psWindowTitle*" } | Select-Object -First 1
            if ($null -eq $match) { throw "No window found with title containing: ${windowTitle}" }
            $target = $match.Handle
          } elseif ('${processName}' -ne '') {
            $list = Get-WindowsList $filter
            $search = $psProcessName.Replace('.exe','')
            $match = $list | Where-Object { $_.Process -like "*$search*" } | Select-Object -First 1
            if ($null -eq $match) { throw "No window found for process: ${processName}" }
            $target = $match.Handle
          }

          if ($target -eq [IntPtr]::Zero) { throw 'No target window resolved' }

          # Minimized checks and optional restore
          $wasMinimized = [Win32]::IsIconic($target)
          if ($wasMinimized -and -not $psRestore) {
            throw 'Target window is minimized. Set restoreIfMinimized: true to capture.'
          }
          $prevForeground = [Win32]::GetForegroundWindow()
          if ($psRestore -and $wasMinimized) {
            [void][Win32]::ShowWindowAsync($target, 9) # SW_RESTORE
            for ($i=0; $i -lt 40; $i++) {
              if (-not [Win32]::IsIconic($target)) { break }
              Start-Sleep -Milliseconds 50
            }
          }

          # Get bounds and capture with padding
          $rc = New-Object Win32+RECT
          [void][Win32]::GetWindowRect($target, [ref]$rc)
          $pad = 10
          $left = [Math]::Max(0, $rc.Left - $pad)
          $top = [Math]::Max(0, $rc.Top - $pad)
          $width = ($rc.Right + $pad) - $left
          $height = ($rc.Bottom + $pad) - $top

          # Try background capture first using PrintWindow (PW_RENDERFULLCONTENT)
          $bmp = New-Object System.Drawing.Bitmap ([int]$width), ([int]$height)
          $gfx = [System.Drawing.Graphics]::FromImage($bmp)
          $hdc = $gfx.GetHdc()
          $pwOk = [Win32]::PrintWindow($target, $hdc, 2)
          $gfx.ReleaseHdc($hdc)

          if (-not $pwOk) {
            # Fallback only if focusing is allowed
            $gfx.Dispose(); $bmp.Dispose()
            if ($psAllowFocus) {
              [void][Win32]::SetForegroundWindow($target)
              Start-Sleep -Milliseconds 200
              $bmp2 = New-Object System.Drawing.Bitmap ([int]$width), ([int]$height)
              $gfx2 = [System.Drawing.Graphics]::FromImage($bmp2)
              $gfx2.CopyFromScreen($left, $top, 0, 0, $bmp2.Size)
              $bmp2.Save('${windowsPath.replace(/\\/g, '\\\\')}', [System.Drawing.Imaging.ImageFormat]::Png)
              $gfx2.Dispose(); $bmp2.Dispose()
              Write-Host 'Window screenshot saved successfully (foreground fallback)'
            } else {
              throw 'Background capture produced an invalid image for this app. Retry with allowFocus: true.'
            }
          } else {
            $bmp.Save('${windowsPath.replace(/\\/g, '\\\\')}', [System.Drawing.Imaging.ImageFormat]::Png)
            $gfx.Dispose(); $bmp.Dispose()
            Write-Host 'Window screenshot saved successfully (background)'
          }

          # Attempt to restore original state/focus
          if ($wasMinimized) { [void][Win32]::ShowWindowAsync($target, 6) } # SW_MINIMIZE
          if ($prevForeground -ne [IntPtr]::Zero) { [void][Win32]::SetForegroundWindow($prevForeground) }
        `;
      } else if (monitor === 'all') {
        psScript = `
          Add-Type -AssemblyName System.Windows.Forms
          Add-Type -AssemblyName System.Drawing
          Add-Type @"
            using System.Runtime.InteropServices; public class DPI { [DllImport("shcore.dll")] public static extern int SetProcessDpiAwareness(int value); }
"@
          [DPI]::SetProcessDpiAwareness(2)
          $screen = [System.Windows.Forms.SystemInformation]::VirtualScreen
          $bmp = New-Object System.Drawing.Bitmap $screen.Width, $screen.Height
          $gfx = [System.Drawing.Graphics]::FromImage($bmp)
          $gfx.CopyFromScreen($screen.Left, $screen.Top, 0, 0, $bmp.Size)
          $bmp.Save('${windowsPath.replace(/\\/g, '\\\\')}', [System.Drawing.Imaging.ImageFormat]::Png)
          $gfx.Dispose(); $bmp.Dispose()
          Write-Host 'Screenshot saved successfully'
        `;
      } else {
        psScript = `
          Add-Type -AssemblyName System.Windows.Forms
          Add-Type -AssemblyName System.Drawing
          Add-Type @"
            using System; using System.Runtime.InteropServices; public class DPIAware { [DllImport("user32.dll")] public static extern bool SetProcessDPIAware(); [DllImport("shcore.dll")] public static extern int SetProcessDpiAwareness(int value); }
"@
          [DPIAware]::SetProcessDpiAwareness(2)
          $screens = [System.Windows.Forms.Screen]::AllScreens | Sort-Object { $_.Bounds.X }
          if ('${monitor}' -eq 'primary') { $target = [System.Windows.Forms.Screen]::PrimaryScreen }
          elseif ('${monitor}' -match '^\\d+$') { $i = [int]'${monitor}' - 1; if ($i -ge 0 -and $i -lt $screens.Count) { $target = $screens[$i] } else { throw "Monitor ${monitor} not found. Available monitors: 1 to $($screens.Count)" } }
          if ($null -eq $target) { throw "Invalid monitor parameter: ${monitor}" }
          $b = $target.Bounds
          $bmp = New-Object System.Drawing.Bitmap($b.Width, $b.Height)
          $gfx = [System.Drawing.Graphics]::FromImage($bmp)
          $gfx.CopyFromScreen($b.X, $b.Y, 0, 0, $bmp.Size)
          $bmp.Save('${windowsPath.replace(/\\/g, '\\\\')}', [System.Drawing.Imaging.ImageFormat]::Png)
          $gfx.Dispose(); $bmp.Dispose()
          Write-Host "Screenshot of monitor ${monitor} saved successfully"
        `;
      }

      const encoded = Buffer.from(psScript, 'utf16le').toString('base64');
      const { stdout, stderr } = await execAsync(`powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -OutputFormat Text -EncodedCommand ${encoded}`);

      const hasRealError = stderr && (
        stderr.includes('throw') ||
        stderr.includes('Exception') ||
        stderr.toLowerCase().includes('methodnotfound') ||
        stderr.toLowerCase().includes("does not contain a method named 'printwindow'") ||
        stderr.toLowerCase().includes('does not contain a method named') ||
        stderr.includes('not found') ||
        (stderr.includes('Error') && !stderr.includes('ErrorId'))
      );
      if (hasRealError) {
        throw new Error(stderr);
      }

      const outputPath = path.join(screenshotsDir, filename);
      await fs.access(outputPath);

      let successPath;
      if (customFolder) {
        successPath = path.join(customFolder, filename).replace(/\\/g, '/');
      } else {
        successPath = `screenshots/${filename}`;
      }
      return { content: [{ type: 'text', text: `Screenshot saved successfully to: ${successPath}` }] };
    } catch (error) {
      const parts = [];
      parts.push('Failed to take screenshot:');
      if (error && error.message) parts.push(String(error.message));
      if (error && error.stderr) parts.push('stderr:', String(error.stderr).trim());
      if (error && error.stdout) parts.push('stdout:', String(error.stdout).trim());
      return { content: [{ type: 'text', text: parts.join(' ') }], isError: true };
    }
  }

  return { content: [{ type: 'text', text: `Unknown tool: ${request.params.name}` }], isError: true };
});

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
  console.error('Screenshot MCP server (triage) running...');
}

main().catch((error) => {
  console.error('Server error:', error);
  process.exit(1);
});
