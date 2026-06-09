#!/usr/bin/env bash
# setup-ipu6-webcam.sh
# Enables the Intel IPU6 MIPI webcam on Ubuntu 24.04 (Noble) for Alder/Raptor Lake laptops.
# Tested on: Dell XPS with Intel Alder Lake (12th gen), kernel 6.17, ov01a10 sensor.
#
# What this does:
#   1. Adds the Intel IPU6 OEM PPA (ppa:oem-solutions-group/intel-ipu6)
#   2. Installs kernel modules, camera HAL, and GStreamer plugin for your IPU6 variant
#   3. Installs and configures v4l2-relayd to expose the camera as /dev/videoN
#   4. Enables the service at boot
#   5. Captures a test frame to verify everything works
#
# NOTE: The PPA is marked as a development channel by Canonical's OEM team.
# It works well but may receive breaking updates. If the camera stops working
# after an apt upgrade, re-run this script.
#
# Usage: sudo ./setup-ipu6-webcam.sh

set -euo pipefail

# --- helpers -----------------------------------------------------------------

info()  { echo "[INFO]  $*"; }
warn()  { echo "[WARN]  $*"; }
die()   { echo "[ERROR] $*" >&2; exit 1; }

require_root() {
    [[ $EUID -eq 0 ]] || die "Run this script with sudo: sudo $0"
}

# --- detect IPU6 variant -----------------------------------------------------

detect_ipu6_variant() {
    local pci_desc
    pci_desc=$(lspci | grep -i 'imaging signal\|ipu6\|IPU' || true)

    if echo "$pci_desc" | grep -qi 'meteor lake\|ipu6epmtl'; then
        echo "ipu6epmtl"
    elif echo "$pci_desc" | grep -qi 'alder lake\|raptor lake\|ipu6ep'; then
        echo "ipu6ep"
    elif echo "$pci_desc" | grep -qi 'tiger lake\|ipu6'; then
        echo "ipu6"
    else
        warn "Could not auto-detect IPU6 variant from PCI; defaulting to ipu6ep (Alder/Raptor Lake)."
        echo "ipu6ep"
    fi
}

# --- main --------------------------------------------------------------------

require_root

KERNEL=$(uname -r)
info "Kernel: $KERNEL"

# Detect IPU6 variant
VARIANT=$(detect_ipu6_variant)
info "Detected IPU6 variant: $VARIANT"

# 1. Add PPA
info "Adding Intel IPU6 OEM PPA..."
add-apt-repository -y ppa:oem-solutions-group/intel-ipu6

# 2. Install base packages
info "Installing base packages..."
apt-get install -y \
    libcamera-tools \
    libcamera-ipa \
    linux-modules-ipu6-"${KERNEL}" \
    v4l-utils

# Fix DKMS directory if missing (sometimes absent on fresh installs)
mkdir -p /var/lib/dkms

# 3. Install v4l2-relayd and v4l2loopback
info "Installing v4l2-relayd and v4l2loopback..."
apt-get install -y v4l2-relayd v4l2loopback-dkms

# 4. Install the GStreamer plugin and HAL for the detected variant
info "Installing GStreamer icamera plugin and libcamhal for $VARIANT..."
apt-get install -y \
    gstreamer1.0-icamera \
    gstreamer1.0-tools \
    "libcamhal-${VARIANT}" \
    "libcamhal-${VARIANT}-common"

# 5. Load the IPU6 processing system module if not already loaded
if ! lsmod | grep -q intel_ipu6_psys; then
    info "Loading intel_ipu6_psys kernel module..."
    modprobe intel_ipu6_psys
fi

# 6. Load v4l2loopback with the label expected by libcamhal-common
if ! lsmod | grep -q v4l2loopback; then
    info "Loading v4l2loopback..."
    modprobe v4l2loopback devices=1 exclusive_caps=1 card_label="Intel MIPI Camera"
else
    info "v4l2loopback already loaded."
fi

# 6b. Apply colour correction (hue/saturation tuned for this hardware)
# libcamhal-common writes /etc/v4l2-relayd with a plain icamerasrc source;
# we extend it to include a videobalance filter. hue=-0.08/saturation=0.88
# counteracts the warm/red cast of the ov01a10 sensor on this Dell XPS.
RELAYD_CONF=/etc/v4l2-relayd
if grep -q '^VIDEOSRC=icamerasrc' "$RELAYD_CONF" && ! grep -q 'videobalance' "$RELAYD_CONF"; then
    info "Applying colour correction to $RELAYD_CONF ..."
    sed -i 's|^VIDEOSRC=icamerasrc buffer-count=7$|VIDEOSRC=icamerasrc buffer-count=7 ! videoconvert ! videobalance hue=-0.08 saturation=0.88 ! videoconvert|' "$RELAYD_CONF"
fi

# 7. Start and enable v4l2-relayd
info "Enabling and starting v4l2-relayd..."
systemctl daemon-reload
systemctl enable v4l2-relayd
systemctl restart v4l2-relayd
sleep 2

# 8. Verify
STATUS=$(systemctl is-active v4l2-relayd || true)
if [[ "$STATUS" != "active" ]]; then
    warn "v4l2-relayd is not active. Check: journalctl -u v4l2-relayd"
    exit 1
fi
info "v4l2-relayd is active."

# Find the virtual camera device
VDEV=$(grep -rl "Intel MIPI Camera" /sys/devices/virtual/video4linux/*/name 2>/dev/null | head -1 | cut -d/ -f6 || true)
if [[ -z "$VDEV" ]]; then
    warn "Could not find 'Intel MIPI Camera' virtual device. Service may still be initialising."
else
    VDEV="/dev/$VDEV"
    info "Virtual camera device: $VDEV"

    # 9. Capture a test frame
    TESTFILE=$(mktemp /tmp/webcam_test_XXXXXX.jpg)
    info "Capturing test frame to $TESTFILE ..."
    if gst-launch-1.0 -q v4l2src device="$VDEV" num-buffers=1 ! videoconvert ! jpegenc ! filesink location="$TESTFILE" 2>/dev/null; then
        SIZE=$(stat -c%s "$TESTFILE")
        info "Test frame captured: $SIZE bytes."
        info "View it with: xdg-open $TESTFILE"
    else
        warn "Test frame capture failed. The service is running but the camera pipeline may need a moment."
    fi
fi

echo ""
info "Setup complete."
info "Your webcam is exposed as a standard V4L2 device (look for 'Intel MIPI Camera')."
info "It will start automatically on every boot."
info ""
info "To verify at any time:"
info "  v4l2-ctl --list-devices"
info "  systemctl status v4l2-relayd"
