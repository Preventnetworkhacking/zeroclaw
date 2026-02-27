#!/bin/bash
#
# ZeroClaw ADB Installer
#
# Installs ZeroClaw APK to connected Android device via ADB.
#
# Usage: curl -fsSL https://zeroclaw.dev/adb | bash
#
# Requirements:
# - ADB installed on computer
# - Android device connected with USB debugging enabled
# - "Install from unknown sources" enabled on device
#

set -e

# Config
APK_URL="https://github.com/zeroclaw-labs/zeroclaw/releases/latest/download/zeroclaw-android.apk"
APK_NAME="zeroclaw-android.apk"

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
NC='\033[0m'

echo -e "${CYAN}"
echo "  ╔═══════════════════════════════════════╗"
echo "  ║      🦀 ZeroClaw ADB Installer        ║"
echo "  ╚═══════════════════════════════════════╝"
echo -e "${NC}"

# Check for ADB
if ! command -v adb &> /dev/null; then
    echo -e "${RED}Error: ADB not found${NC}"
    echo ""
    echo "Install ADB:"
    echo "  macOS:   brew install android-platform-tools"
    echo "  Ubuntu:  sudo apt install adb"
    echo "  Windows: Download from developer.android.com"
    exit 1
fi

# Check for connected device
echo -e "${BLUE}[1/4]${NC} Checking for connected device..."
DEVICE=$(adb devices | grep -w "device" | head -1 | cut -f1)

if [ -z "$DEVICE" ]; then
    echo -e "${RED}Error: No device connected${NC}"
    echo ""
    echo "Make sure:"
    echo "  1. USB debugging is enabled on your phone"
    echo "  2. Phone is connected via USB"
    echo "  3. You've authorized this computer on your phone"
    echo ""
    echo "To enable USB debugging:"
    echo "  Settings → About Phone → Tap 'Build number' 7 times"
    echo "  Settings → Developer Options → Enable USB debugging"
    exit 1
fi

echo -e "  Found device: ${GREEN}$DEVICE${NC}"

# Download APK
echo -e "${BLUE}[2/4]${NC} Downloading ZeroClaw APK..."
TEMP_DIR=$(mktemp -d)
cd "$TEMP_DIR"

if command -v curl &> /dev/null; then
    curl -fsSL -o "$APK_NAME" "$APK_URL"
elif command -v wget &> /dev/null; then
    wget -q -O "$APK_NAME" "$APK_URL"
else
    echo -e "${RED}Error: curl or wget required${NC}"
    exit 1
fi

echo -e "  Downloaded to: ${CYAN}$TEMP_DIR/$APK_NAME${NC}"

# Install APK
echo -e "${BLUE}[3/4]${NC} Installing APK to device..."
adb -s "$DEVICE" install -r "$APK_NAME"

# Cleanup
echo -e "${BLUE}[4/4]${NC} Cleaning up..."
rm -rf "$TEMP_DIR"

echo ""
echo -e "${GREEN}╔═══════════════════════════════════════╗${NC}"
echo -e "${GREEN}║     ✓ ZeroClaw installed!             ║${NC}"
echo -e "${GREEN}╚═══════════════════════════════════════╝${NC}"
echo ""
echo -e "Open ${CYAN}ZeroClaw${NC} on your phone to get started"
echo ""
echo -e "${YELLOW}Optional: Grant all permissions${NC}"
echo "  adb shell pm grant ai.zeroclaw.android android.permission.POST_NOTIFICATIONS"
echo ""
