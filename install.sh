#!/usr/bin/env bash
# ==============================================================================
# Ping Drop Monitor - Automated Installation Script
# ==============================================================================
set -e

INSTALL_DIR="/ping"
SERVICE_NAME="ping-monitor"
SOURCE_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

echo "=========================================================="
echo " Installing 24x7 Ping Drop Monitor into: $INSTALL_DIR"
echo "=========================================================="

# 1. Install Python3 & Pip if missing
if ! command -v python3 &> /dev/null; then
    echo "[+] Installing Python3..."
    if command -v dnf &> /dev/null; then
        dnf install -y python3 python3-pip
    elif command -v apt-get &> /dev/null; then
        apt-get update && apt-get install -y python3 python3-pip
    elif command -v yum &> /dev/null; then
        yum install -y python3 python3-pip
    fi
fi

# 2. Install Python Dependencies
echo "[+] Installing Python dependencies (Flask, xlwt)..."
python3 -m pip install -r "$SOURCE_DIR/requirements.txt"

# 3. Create target directory and copy files
echo "[+] Setting up $INSTALL_DIR directory..."
mkdir -p "$INSTALL_DIR/templates"

cp "$SOURCE_DIR/monitor.py" "$INSTALL_DIR/monitor.py"
cp "$SOURCE_DIR/status.sh" "$INSTALL_DIR/status.sh"
cp "$SOURCE_DIR/templates/index.html" "$INSTALL_DIR/templates/index.html"

chmod +x "$INSTALL_DIR/monitor.py"
chmod +x "$INSTALL_DIR/status.sh"

if [ ! -f "$INSTALL_DIR/targets.json" ]; then
    echo "[+] Creating targets.json from template..."
    cp "$SOURCE_DIR/targets.example.json" "$INSTALL_DIR/targets.json"
fi

# 4. Open Firewall Port 8084 if firewalld is active
if command -v firewall-cmd &> /dev/null && systemctl is-active firewalld &> /dev/null; then
    echo "[+] Opening port 8084 in firewalld..."
    firewall-cmd --permanent --add-port=8084/tcp || true
    firewall-cmd --reload || true
fi

# 5. Install Systemd Service
echo "[+] Configuring systemd service: /etc/systemd/system/${SERVICE_NAME}.service"
cp "$SOURCE_DIR/systemd/ping-monitor.service" "/etc/systemd/system/${SERVICE_NAME}.service"

systemctl daemon-reload
systemctl enable "${SERVICE_NAME}.service"
systemctl restart "${SERVICE_NAME}.service"

echo ""
echo "=========================================================="
echo " Setup Completed Successfully!"
echo "=========================================================="
echo " - Web GUI URL    : http://<your-server-ip>:8084"
echo " - Excel Reports  : $INSTALL_DIR/report.xls"
echo " - Targets Config : $INSTALL_DIR/targets.json"
echo " - Service Status : systemctl status $SERVICE_NAME"
echo " - Quick CLI Test : $INSTALL_DIR/status.sh"
echo "=========================================================="
