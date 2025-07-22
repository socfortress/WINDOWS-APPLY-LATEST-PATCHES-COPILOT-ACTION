# PowerShell Apply Latest Patches Template

This repository provides a template for PowerShell-based active response scripts for security automation and incident response. The template ensures consistent logging, error handling, and execution flow for patch management and update automation.

---

## Overview

The `Apply-Latest-Patches.ps1` script automates the process of checking for and installing critical Windows updates. It logs all actions, results, and errors in both a script log and an active-response log, making it suitable for integration with SOAR platforms, SIEMs, and incident response workflows.

---

## Template Structure

### Core Components

- **Parameter Definitions**: Configurable script parameters
- **Logging Framework**: Comprehensive logging with rotation
- **Error Handling**: Structured exception management
- **JSON Output**: Standardized response format
- **Execution Timing**: Performance monitoring

---

## How Scripts Are Invoked

### Command Line Execution

```powershell
.\Apply-Latest-Patches.ps1 [-LogPath <string>] [-ARLog <string>]
```

### Parameters

| Parameter | Type   | Default Value                                                    | Description                                  |
|-----------|--------|------------------------------------------------------------------|----------------------------------------------|
| `LogPath` | string | `$env:TEMP\Apply-Latest-Patches-script.log`                      | Path for execution logs                      |
| `ARLog`   | string | `C:\Program Files (x86)\ossec-agent\active-response\active-responses.log` | Path for active response JSON output         |

---

### Example Invocations

```powershell
# Basic execution with default parameters
.\Apply-Latest-Patches.ps1

# Custom log path
.\Apply-Latest-Patches.ps1 -LogPath "C:\Logs\ApplyPatches.log"

# Integration with OSSEC/Wazuh active response
.\Apply-Latest-Patches.ps1 -ARLog "C:\ossec\active-responses.log"
```

---

## Template Functions

### `Write-Log`
**Purpose**: Standardized logging with severity levels and console output.

**Parameters**:
- `Message` (string): The log message
- `Level` (ValidateSet): Log level - 'INFO', 'WARN', 'ERROR', 'DEBUG'

**Features**:
- Timestamped, color-coded output
- File logging
- Verbose/debug support

**Usage**:
```powershell
Write-Log "Checking for updates..." 'INFO'
Write-Log "Results JSON logged to $ARLog" 'INFO'
Write-Log "Error occurred: ..." 'ERROR'
```

---

### `Rotate-Log`
**Purpose**: Manages log file size and rotation.

**Features**:
- Monitors log file size (default: 100KB)
- Maintains a configurable number of backups (default: 5)
- Rotates logs automatically

**Configuration Variables**:
- `$LogMaxKB`: Max log file size in KB
- `$LogKeep`: Number of rotated logs to retain

---

## Script Execution Flow

1. **Initialization**
   - Parameter validation and assignment
   - Error action preference
   - Log rotation

2. **Execution**
   - Logs script start
   - Checks for PSWindowsUpdate module, installs if missing
   - Checks for available updates and logs results
   - Installs all available updates and logs results

3. **Completion**
   - Logs results as JSON to the active response log
   - Logs script end and duration

4. **Error Handling**
   - Catches and logs exceptions
   - Outputs error details as JSON

---

## JSON Output Format

### Success Response

```json
{
  "timestamp": "2025-07-22T10:30:45.123Z",
  "host": "HOSTNAME",
  "action": "check_critical_updates",
  "update_count": 2,
  "updates": [
    {
      "guid": "...",
      "title": "...",
      "kb_article": "...",
      "categories": "...",
      "severity": "...",
      "download_sizeMB": 12.34,
      "is_downloaded": true,
      "is_installed": false,
      "publication_date": "..."
    }
  ],
  "status": "success"
}
```

```json
{
  "timestamp": "2025-07-22T10:31:10.456Z",
  "host": "HOSTNAME",
  "action": "install_critical_updates",
  "installed": [
    {
      "title": "...",
      "kb_article": "...",
      "result": "Installed",
      "reboot": false
    }
  ],
  "status": "completed"
}
```

### Error Response

```json
{
  "timestamp": "2025-07-22T10:31:10.456Z",
  "host": "HOSTNAME",
  "action": "install_critical_updates",
  "status": "error",
  "error": "Access is denied"
}
```

---

## Implementation Guidelines

1. Use the provided logging and error handling functions.
2. Customize the main logic as needed for your environment.
3. Ensure JSON output matches your SOAR/SIEM requirements.
4. Test thoroughly in a non-production environment.

---

## Security Considerations

- Run with the minimum required privileges.
- Validate all input parameters.
- Secure log files and output locations.
- Monitor for errors and failed updates.

---

## Troubleshooting

- **Permission Errors**: Run as Administrator.
- **Module Not Found**: Ensure internet access for module installation.
- **Log Rotation**: Check file permissions and disk space.

Enable verbose logging with `-Verbose` for debugging:
```powershell
.\Apply-Latest-Patches.ps1 -Verbose
```

---

## License

This template is provided as-is for security automation and incident response purposes.
