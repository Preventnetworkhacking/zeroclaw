#!/data/data/com.termux/files/usr/bin/bash
#
# ZeroClaw Termux Installer
# 
# Usage: curl -fsSL https://zeroclaw.dev/termux | bash
#
# This script installs ZeroClaw in Termux with minimal user interaction.
#

set -e

# Colors
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
NC='\033[0m'

echo -e "${CYAN}"
echo "  ╔═══════════════════════════════════════╗"
echo "  ║     🦀 ZeroClaw Termux Installer      ║"
echo "  ╚═══════════════════════════════════════╝"
echo -e "${NC}"

# Check if running in Termux
if [ ! -d "/data/data/com.termux" ]; then
    echo -e "${RED}Error: This script must be run in Termux${NC}"
    echo "Install Termux from F-Droid: https://f-droid.org/packages/com.termux/"
    exit 1
fi

echo -e "${BLUE}[1/5]${NC} Updating packages..."
pkg update -y && pkg upgrade -y

echo -e "${BLUE}[2/5]${NC} Installing dependencies..."
pkg install -y rust git openssl

echo -e "${BLUE}[3/5]${NC} Installing ZeroClaw..."
cargo install zeroclaw

echo -e "${BLUE}[4/5]${NC} Setting up config directory..."
mkdir -p ~/.config/zeroclaw

echo -e "${BLUE}[5/5]${NC} Running setup wizard..."
echo ""

# Check if already configured
if [ -f ~/.config/zeroclaw/config.toml ]; then
    echo -e "${YELLOW}Config already exists. Skipping wizard.${NC}"
else
    echo -e "${GREEN}Starting configuration wizard...${NC}"
    echo ""
    zeroclaw init
fi

echo ""
echo -e "${GREEN}╔═══════════════════════════════════════╗${NC}"
echo -e "${GREEN}║     ✓ ZeroClaw installed!             ║${NC}"
echo -e "${GREEN}╚═══════════════════════════════════════╝${NC}"
echo ""
echo -e "Run ${CYAN}zeroclaw${NC} to start"
echo -e "Run ${CYAN}zeroclaw config${NC} to change settings"
echo ""
echo -e "${YELLOW}Tip: Add to .bashrc for auto-start:${NC}"
echo "  echo 'zeroclaw daemon &' >> ~/.bashrc"
echo ""
