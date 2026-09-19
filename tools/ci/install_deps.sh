#!/usr/bin/env bash
set -euo pipefail
if [[ "$(uname)" == Darwin ]]; then
    brew install ccache libxml2 minizip-ng zlib-ng pkgconf
    brew list --versions ccache libxml2 minizip-ng zlib-ng
elif command -v dnf >/dev/null; then
    dnf install -y --setopt=install_weak_deps=False epel-release
    dnf install -y --setopt=install_weak_deps=False ccache libxml2-devel minizip-devel zlib-devel
    rpm -q ccache libxml2-devel minizip-devel zlib-devel | tee /host/tmp/turboxl-ccache/dependencies.txt
else
    sudo apt-get update
    sudo apt-get install -y ccache libxml2-dev libminizip-dev zlib1g-dev ninja-build
    dpkg-query -W ccache libxml2-dev libminizip-dev zlib1g-dev
fi
