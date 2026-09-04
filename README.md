# Microsoft 365 for Linux (Unofficial)

An Electron-based Flatpak desktop application integrating the Microsoft 365 web suite with native OneDrive synchronization on Linux. Features continuous background sync, local document editing in Office Online, and atomic local saving with real-time renaming.

## Features
- Full Office Online Integration (Word, Excel, PowerPoint, OneDrive, Outlook, OneNote, Copilot)
- Native OneDrive Sync using abraunegg/onedrive
- Inotify-driven real-time background synchronization
- Open local documents directly in Office Online via managed cloud links
- Atomic local file saving with rename tracking
- System tray and background portal autostart support

## Prerequisites
Install flatpak and flatpak-builder on your distribution:
- Debian/Ubuntu/Mint: `sudo apt install flatpak flatpak-builder`
- Arch Linux: `sudo pacman -S flatpak flatpak-builder`
- Fedora: `sudo dnf install flatpak flatpak-builder`

Install required runtimes:
```bash
flatpak install flathub org.gnome.Platform//46 org.gnome.Sdk//46
```

## Building and Running
```bash
./build.sh
flatpak install --user --reinstall -y Microsoft365-for-Linux-Unofficial.flatpak
flatpak run com.microsoft365linux.Unofficial
```

## License
Unofficial open-source software. Microsoft 365 and OneDrive are trademarks of Microsoft Corporation.
