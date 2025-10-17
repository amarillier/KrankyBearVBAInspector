package main

import (
	"archive/zip"
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"runtime"
	"strings"
)

const (
	version     = "0.1.0"
	colorRed    = "\033[31m"
	colorGreen  = "\033[32m"
	colorYellow = "\033[33m"
	colorReset  = "\033[0m"
)

var briefMode bool // Global flag for brief output mode

type VBAInspectionResult struct {
	FileName             string
	Has32BitDeclarations bool
	Declarations32Bit    []string
	DeclarationsSafe     []string
	VBAProjectFound      bool
}

func main() {
	if len(os.Args) < 2 {
		printHelp()
		os.Exit(1)
	}

	// Parse arguments
	var filePath string
	briefMode = false

	for i := 1; i < len(os.Args); i++ {
		arg := os.Args[i]

		// Handle help flags
		if arg == "-?" || arg == "-help" || arg == "--help" {
			printHelp()
			os.Exit(0)
		}

		// Handle version flag
		if arg == "-v" || arg == "--version" {
			printVersion()
			os.Exit(0)
		}

		// Handle brief flag
		if arg == "-b" || arg == "-brief" || arg == "--brief" {
			briefMode = true
			continue
		}

		// Remaining argument is the file path
		if filePath == "" {
			filePath = arg
		} else {
			fmt.Fprintf(os.Stderr, "Error: Unexpected argument '%s'\n", arg)
			printHelp()
			os.Exit(1)
		}
	}

	// Check if we have a file path
	if filePath == "" {
		fmt.Fprintf(os.Stderr, "Error: No file specified\n\n")
		printHelp()
		os.Exit(1)
	}

	// Check if file exists
	if _, err := os.Stat(filePath); os.IsNotExist(err) {
		fmt.Fprintf(os.Stderr, "Error: File '%s' does not exist\n", filePath)
		os.Exit(1)
	}

	result, err := inspectVBAProject(filePath)
	if err != nil {
		fmt.Fprintf(os.Stderr, "Error: %v\n", err)
		os.Exit(1)
	}

	printResults(result)

	// Exit with error code if 32-bit declarations found
	if result.Has32BitDeclarations {
		os.Exit(1)
	}
}

// printHelp displays usage information
func printHelp() {
	progName := filepath.Base(os.Args[0])
	fmt.Fprintf(os.Stderr, "VBA Inspector v%s - Detect 32-bit only VBA declarations\n\n", version)
	fmt.Fprintf(os.Stderr, "Usage:\n")
	fmt.Fprintf(os.Stderr, "  %s [options] <office-file>\n", progName)
	fmt.Fprintf(os.Stderr, "  %s -v | --version\n", progName)
	fmt.Fprintf(os.Stderr, "  %s -? | -help | --help\n\n", progName)
	fmt.Fprintf(os.Stderr, "Arguments:\n")
	fmt.Fprintf(os.Stderr, "  <office-file>     Path to Excel (.xlsx, .xlsm), Word (.docx, .docm),\n")
	fmt.Fprintf(os.Stderr, "                    or PowerPoint (.pptx, .pptm) file to inspect\n\n")
	fmt.Fprintf(os.Stderr, "Options:\n")
	fmt.Fprintf(os.Stderr, "  -b, -brief, --brief       Brief output mode (simple text list of unsafe declarations)\n")
	fmt.Fprintf(os.Stderr, "  -v, --version             Show version and platform information\n")
	fmt.Fprintf(os.Stderr, "  -?, -help, --help         Show this help message\n\n")
	fmt.Fprintf(os.Stderr, "Description:\n")
	fmt.Fprintf(os.Stderr, "  Inspects Microsoft Office files for VBA function declarations that use\n")
	fmt.Fprintf(os.Stderr, "  'Declare Function' without 'PtrSafe', which are only compatible with\n")
	fmt.Fprintf(os.Stderr, "  32-bit Office and will fail in 64-bit Office environments.\n\n")
	fmt.Fprintf(os.Stderr, "Exit Codes:\n")
	fmt.Fprintf(os.Stderr, "  0  No issues found (no 32-bit only declarations)\n")
	fmt.Fprintf(os.Stderr, "  1  Issues found or error occurred\n\n")
	fmt.Fprintf(os.Stderr, "Examples:\n")
	fmt.Fprintf(os.Stderr, "  %s myfile.xlsm\n", progName)
	fmt.Fprintf(os.Stderr, "  %s -brief myfile.xlsm\n", progName)
	fmt.Fprintf(os.Stderr, "  %s -v\n", progName)
	fmt.Fprintf(os.Stderr, "  %s -help\n", progName)
}

// printVersion displays version and platform information
func printVersion() {
	fmt.Printf("VBA Inspector v%s\n", version)
	fmt.Printf("Platform: %s/%s\n", runtime.GOOS, runtime.GOARCH)
	fmt.Printf("Go version: %s\n", runtime.Version())
}

// inspectVBAProject opens an Office file and inspects its VBA project
func inspectVBAProject(filePath string) (*VBAInspectionResult, error) {
	result := &VBAInspectionResult{
		FileName:          filepath.Base(filePath),
		Declarations32Bit: []string{},
		DeclarationsSafe:  []string{},
		VBAProjectFound:   false,
	}

	// Open the Office file as a ZIP archive
	reader, err := zip.OpenReader(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open file as ZIP archive: %w", err)
	}
	defer reader.Close()

	// Look for vbaProject.bin in various locations
	// Excel: xl/vbaProject.bin
	// Word: word/vbaProject.bin
	// PowerPoint: ppt/vbaProject.bin
	possiblePaths := []string{
		"xl/vbaProject.bin",
		"word/vbaProject.bin",
		"ppt/vbaProject.bin",
	}

	var vbaProjectData []byte
	for _, file := range reader.File {
		for _, path := range possiblePaths {
			if file.Name == path {
				result.VBAProjectFound = true
				vbaProjectData, err = readZipFile(file)
				if err != nil {
					return nil, fmt.Errorf("failed to read vbaProject.bin: %w", err)
				}
				break
			}
		}
		if result.VBAProjectFound {
			break
		}
	}

	if !result.VBAProjectFound {
		return result, nil
	}

	// Search for function declarations in the binary data
	result.Declarations32Bit, result.DeclarationsSafe = findFunctionDeclarations(vbaProjectData)
	result.Has32BitDeclarations = len(result.Declarations32Bit) > 0

	return result, nil
}

// readZipFile reads the contents of a file from a ZIP archive
func readZipFile(file *zip.File) ([]byte, error) {
	rc, err := file.Open()
	if err != nil {
		return nil, err
	}
	defer rc.Close()

	buf := new(bytes.Buffer)
	_, err = buf.ReadFrom(rc)
	if err != nil {
		return nil, err
	}

	return buf.Bytes(), nil
}

// findFunctionDeclarations searches for VBA function declarations in binary data
func findFunctionDeclarations(data []byte) (unsafe []string, safe []string) {
	// Convert binary data to string for pattern matching
	content := string(data)

	// Pattern to find "Declare" followed by optional "PtrSafe" and then "Function"
	// We'll look for both patterns and distinguish between them

	// Find all occurrences of "Declare" in the content
	declarePattern := regexp.MustCompile(`Declare\s+(?:PtrSafe\s+)?(?:Function|Sub)\s+\w+`)
	ptrSafePattern := regexp.MustCompile(`Declare\s+PtrSafe\s+(?:Function|Sub)`)

	matches := declarePattern.FindAllString(content, -1)

	// Track unique declarations
	unsafeMap := make(map[string]bool)
	safeMap := make(map[string]bool)

	for _, match := range matches {
		// Clean up the match (remove null bytes and non-printable characters)
		cleanMatch := cleanString(match)
		if cleanMatch == "" {
			continue
		}

		// Check if this is a PtrSafe declaration
		if ptrSafePattern.MatchString(match) {
			safeMap[cleanMatch] = true
		} else {
			unsafeMap[cleanMatch] = true
		}
	}

	// Convert maps to slices
	for decl := range unsafeMap {
		unsafe = append(unsafe, decl)
	}
	for decl := range safeMap {
		safe = append(safe, decl)
	}

	return unsafe, safe
}

// cleanString removes null bytes and non-printable characters, keeping only readable text
func cleanString(s string) string {
	// Remove null bytes and control characters except space
	var result strings.Builder
	for _, r := range s {
		if r >= 32 && r < 127 || r == '\t' || r == '\n' || r == '\r' {
			result.WriteRune(r)
		} else if r == 0 {
			// Skip null bytes
			continue
		}
	}

	cleaned := result.String()
	// Normalize whitespace
	cleaned = regexp.MustCompile(`\s+`).ReplaceAllString(cleaned, " ")
	return strings.TrimSpace(cleaned)
}

// printResults displays the inspection results
func printResults(result *VBAInspectionResult) {
	// Brief mode: simple text output of unsafe declarations only
	if briefMode {
		if !result.VBAProjectFound {
			// No output in brief mode if no VBA project
			return
		}

		if result.Has32BitDeclarations {
			for _, decl := range result.Declarations32Bit {
				fmt.Printf("Unsafe 32 bit: %s\n", decl)
			}
		}
		// No output if no unsafe declarations in brief mode
		return
	}

	// Standard detailed output
	fmt.Printf("\n")
	fmt.Printf("═══════════════════════════════════════════════════════════════\n")
	fmt.Printf("  VBA Inspector Report\n")
	fmt.Printf("═══════════════════════════════════════════════════════════════\n")
	fmt.Printf("File: %s\n\n", result.FileName)

	if !result.VBAProjectFound {
		fmt.Printf("%s⚠ No VBA project found in this file%s\n", colorYellow, colorReset)
		fmt.Printf("This file does not contain any macros or VBA code.\n")
		return
	}

	// Report summary
	totalDeclarations := len(result.Declarations32Bit) + len(result.DeclarationsSafe)
	fmt.Printf("Summary: Found %d API declaration(s)\n", totalDeclarations)
	fmt.Printf("  - 32-bit only (unsafe): %d\n", len(result.Declarations32Bit))
	fmt.Printf("  - 64-bit compatible (safe): %d\n\n", len(result.DeclarationsSafe))

	if result.Has32BitDeclarations {
		fmt.Printf("%s⚠ WARNING: 32-bit only VBA declarations found!%s\n\n", colorRed, colorReset)
		fmt.Printf("The following function declarations are NOT 64-bit compatible:\n")
		fmt.Printf("(They use 'Declare Function' without 'PtrSafe')\n\n")

		for i, decl := range result.Declarations32Bit {
			fmt.Printf("  %d. %s%s%s\n", i+1, colorRed, decl, colorReset)
		}

		fmt.Printf("\n%sRecommendation:%s\n", colorYellow, colorReset)
		fmt.Printf("Update these declarations to use 'Declare PtrSafe Function' for\n")
		fmt.Printf("64-bit compatibility, or use conditional compilation:\n")
		fmt.Printf("  #If VBA7 Then\n")
		fmt.Printf("    Declare PtrSafe Function ...\n")
		fmt.Printf("  #Else\n")
		fmt.Printf("    Declare Function ...\n")
		fmt.Printf("  #End If\n")
	} else {
		fmt.Printf("%s✓ No 32-bit only declarations found%s\n", colorGreen, colorReset)
		if totalDeclarations > 0 {
			fmt.Printf("All function declarations appear to be 64-bit compatible.\n")
		} else {
			fmt.Printf("No API declarations found in this VBA project.\n")
		}
	}

	if len(result.DeclarationsSafe) > 0 {
		fmt.Printf("\n%s✓ 64-bit compatible declarations found:%s\n", colorGreen, colorReset)
		for i, decl := range result.DeclarationsSafe {
			fmt.Printf("  %d. %s\n", i+1, decl)
		}
	}

	fmt.Printf("\n═══════════════════════════════════════════════════════════════\n\n")
}
