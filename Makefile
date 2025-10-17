# Makefile for VBA Inspector

APP_NAME=vbainspector
VERSION=0.1.0
BUILD_DIR=bin

# Go parameters
GOCMD=go
GOBUILD=$(GOCMD) build
GOCLEAN=$(GOCMD) clean
GOTEST=$(GOCMD) test
GOMOD=$(GOCMD) mod

# Build flags
LDFLAGS=-ldflags "-s -w"

.PHONY: all build clean test windows linux macos-intel macos-arm release help

## all: Build for all platforms (Windows, Linux, macOS Intel & ARM)
all: clean windows linux macos-intel macos-arm
	@echo ""
	@echo "Built all platforms:"
	@ls -lh $(BUILD_DIR)/

## build: Build for current platform
build:
	@echo "Building for current platform..."
	$(GOBUILD) -o $(APP_NAME) $(LDFLAGS) .

## clean: Clean build artifacts
clean:
	@echo "Cleaning..."
	$(GOCLEAN)
	rm -rf $(BUILD_DIR)
	rm -f $(APP_NAME) $(APP_NAME).exe

## test: Run tests
test:
	@echo "Running tests..."
	$(GOTEST) -v ./...

## windows: Build for Windows (64-bit)
windows:
	@echo "Building for Windows (64-bit)..."
	@mkdir -p $(BUILD_DIR)
	GOOS=windows GOARCH=amd64 $(GOBUILD) -o $(BUILD_DIR)/$(APP_NAME)-windows-amd64.exe $(LDFLAGS) .

## linux: Build for Linux (64-bit)
linux:
	@echo "Building for Linux (64-bit)..."
	@mkdir -p $(BUILD_DIR)
	GOOS=linux GOARCH=amd64 $(GOBUILD) -o $(BUILD_DIR)/$(APP_NAME)-linux-amd64 $(LDFLAGS) .

## macos-intel: Build for macOS (Intel)
macos-intel:
	@echo "Building for macOS (Intel)..."
	@mkdir -p $(BUILD_DIR)
	GOOS=darwin GOARCH=amd64 $(GOBUILD) -o $(BUILD_DIR)/$(APP_NAME)-darwin-amd64 $(LDFLAGS) .

## macos-arm: Build for macOS (Apple Silicon)
macos-arm:
	@echo "Building for macOS (Apple Silicon)..."
	@mkdir -p $(BUILD_DIR)
	GOOS=darwin GOARCH=arm64 $(GOBUILD) -o $(BUILD_DIR)/$(APP_NAME)-darwin-arm64 $(LDFLAGS) .
	cp $(BUILD_DIR)/$(APP_NAME)-darwin-arm64 ./$(APP_NAME)
## release: Alias for 'all' - Build for all platforms
release: all

## deps: Download dependencies
deps:
	@echo "Downloading dependencies..."
	$(GOMOD) download
	$(GOMOD) tidy

## help: Show this help message
help:
	@echo "Usage: make [target]"
	@echo ""
	@echo "Targets:"
	@sed -n 's/^##//p' ${MAKEFILE_LIST} | column -t -s ':' | sed -e 's/^/ /'

