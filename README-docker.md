# Docker Deployment Guide

## Prerequisites (Windows)

1. **Docker Desktop** - Install from [docker.com](https://www.docker.com/products/docker-desktop)
2. **X11 Server** (for GUI) - Choose one:
   - **VcXsrv** (Recommended): Download from [sourceforge](https://sourceforge.net/projects/vcxsrv/)
   - **Xming**: Download from [xming.com](http://www.straightrunning.com/XmingNotes/)

## Setup X11 Server (VcXsrv)

1. Install VcXsrv
2. Launch "XLaunch"
3. Choose settings:
   - Display: "Multiple windows"
   - Client startup: "Start no client"
   - Extra settings: Check "Disable access control"
4. Save configuration for future use

## Quick Start

1. **Build the image:**
   ```cmd
   docker-build.bat