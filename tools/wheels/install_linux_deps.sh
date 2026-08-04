#!/usr/bin/env bash
set -euxo pipefail

# manylinux_2_28 is AlmaLinux 8 based. Using its ABI-compatible development
# packages avoids rebuilding the same small C libraries for every wheel run;
# auditwheel vendors non-policy shared libraries into the finished wheel.
dnf install -y --setopt=install_weak_deps=False epel-release
dnf install -y --setopt=install_weak_deps=False \
    ccache \
    libxml2-devel \
    minizip-devel \
    zlib-devel
dnf clean all
