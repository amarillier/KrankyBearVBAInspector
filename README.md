# VBA Inspector

A cross-platform command-line tool to inspect Microsoft Office files (Excel, Word, PowerPoint) for 32-bit only VBA function declarations that are not compatible with 64-bit Office.

## Purpose

This tool helps identify VBA code that uses `Declare Function` without `PtrSafe`, which only works in 32-bit Office versions. Modern 64-bit Office requires `Declare PtrSafe Function` for API calls.

## Features

- ✅ Cross-platform support (Windows, Linux, macOS)
- ✅ Supports Excel (.xlsm, .xlsx), Word (.docm, .docx), and PowerPoint (.pptm, .pptx) files
- ✅ Identifies 32-bit only function declarations
- ✅ Reports 64-bit compatible declarations
- ✅ Color-coded output for easy reading
- ✅ Exit codes for CI/CD integration

## Installation

### Building from Source

```bash
# Clone or download the repository
cd KrankyBearVBAinspector

# Build for your current platform
go build -o vbainspector

# Or build for all platforms (Windows, Linux, macOS Intel & ARM)
make all

# Or build for your current platform using make
make build

# Or build for specific platforms:
make windows      # Windows 64-bit
make linux        # Linux 64-bit
make macos-intel  # macOS Intel
make macos-arm    # macOS Apple Silicon
```

Binaries for all platforms will be created in the `bin/` directory.

## Usage

```bash
vbainspector [options] <path-to-office-file>
vbainspector -v | --version       # Show version information
vbainspector -? | -help | --help  # Show help message
```

### Options

- `-b, -brief, --brief` - Brief output mode (simple text list of unsafe declarations only)
- `-v, --version` - Show version and platform information
- `-?, -help, --help` - Show help message

### Examples

```bash
# Show help message
./vbainspector -help

# Show version and platform information
./vbainspector -v

# Inspect an Excel file (detailed output)
./vbainspector 01-TrainingSchedules.xlsm

# Inspect with brief output (simple text list)
./vbainspector -brief 01-TrainingSchedules.xlsm
./vbainspector -b myfile.xlsm
./vbainspector --brief myfile.xlsm

# Inspect a Word document
./vbainspector myDocument.docm

# Inspect a PowerPoint presentation
./vbainspector myPresentation.pptm
```

## Output

### Standard Mode (Default)

The tool provides a detailed report showing:
- Summary statistics (total declarations, unsafe count, safe count)
- Whether a VBA project was found
- Any 32-bit only function declarations (WARNING in red)
- Any 64-bit compatible declarations (SAFE in green)
- Recommendations for fixing compatibility issues

### Brief Mode (`-b`, `-brief`, or `--brief`)

Simple text output with just the unsafe declarations, one per line:
```
Unsafe 32 bit: Declare Function GetTempPathAllan32
Unsafe 32 bit: Declare Function GetUserName
```

Perfect for:
- Scripting and automation
- Piping to other tools
- Quick scanning of multiple files
- CI/CD pipelines with simple parsing needs

### Exit Codes

- `0`: No issues found (no 32-bit only declarations)
- `1`: Issues found (32-bit only declarations detected) or error occurred

## What It Checks

### ⚠️ 32-bit Only (Unsafe)
```vba
Declare Function GetUserName Lib "advapi32.dll" ...
```

### ✅ 64-bit Compatible (Safe)
```vba
Declare PtrSafe Function GetUserName Lib "advapi32.dll" ...
```

## Recommendations

If 32-bit only declarations are found, you should update them to use conditional compilation:

```vba
#If VBA7 Then
    Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, nSize As Long) As Long
#Else
    Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
        (ByVal lpBuffer As String, nSize As Long) As Long
#End If
```

## Technical Details

Office Open XML files (.xlsx, .xlsm, .docx, .docm, .pptx, .pptm) are ZIP archives containing XML and binary files. The VBA code is stored in a binary file called `vbaProject.bin` located in:
- Excel: `xl/vbaProject.bin`
- Word: `word/vbaProject.bin`
- PowerPoint: `ppt/vbaProject.bin`

### How It Works

This tool operates **entirely in memory** without extracting files to disk:

1. Opens the Office file as a ZIP archive (direct read access)
2. Locates the `vbaProject.bin` file within the ZIP structure
3. Reads `vbaProject.bin` directly into memory (no temp files created)
4. Searches for `Declare Function` and `Declare PtrSafe Function` patterns in the binary data
5. Reports findings with detailed information
6. Automatically cleans up memory (no manual cleanup needed)

**Benefits:**
- ✅ No temporary directory or files created
- ✅ Fast operation (no disk extraction overhead)
- ✅ Safe and secure (no leftover files)
- ✅ Works on read-only filesystems

## License

This tool is provided as-is for inspecting VBA code compatibility.

## Contributing

Contributions, bug reports, and feature requests are welcome!

