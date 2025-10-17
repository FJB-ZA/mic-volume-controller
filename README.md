# Mic Volume Controller

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![.NET](https://img.shields.io/badge/.NET-9.0-blue.svg)](https://dotnet.microsoft.com/download)
[![Platform](https://img.shields.io/badge/platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

A lightweight Windows application that allows you to lock your microphone volume to a specific level and prevent other applications from changing it.

<img width="386" height="253" alt="Mic Volume Controller Screenshot" src="https://github.com/user-attachments/assets/2e0702b3-6ea8-4d85-a67c-5aae7f47f360" />

## ✨ Features

- 🎤 **Select Any Microphone** - Works with all audio input devices
- 🔒 **Lock Volume** - Continuously enforces your chosen volume level
- 💾 **Per-Microphone Memory** - Remembers individual volume settings for each microphone
- 🎚️ **Dual Control** - Use slider for quick adjustments or numeric input for precision
- 🚀 **Start with Windows** - Optional auto-start on system boot
- 📌 **System Tray Support** - Minimize to tray to keep it running in the background
- ⚙️ **Persistent Settings** - All your preferences are saved between sessions
- 🪶 **Lightweight** - Minimal resource usage (~10MB RAM)

## 📥 Installation

### Option 1: Download Executable (Recommended)

1. Go to the [Releases](https://github.com/FJB-ZA/mic-volume-controller/releases) page
2. Download the latest `MicVolumeController.exe`
3. Run the executable - no installation required!
4. Windows SmartScreen might show a warning - click "More info" → "Run anyway"

### Option 2: Build from Source

See the [Building](#-building-from-source) section below.

## 🚀 Usage

1. **Launch the application** - Double-click the executable
2. **Select your microphone** - Choose from the dropdown list
3. **Set your desired volume** - Use the slider or type a precise value (0-100%)
4. **Let it run** - The app continuously maintains your volume setting

### Features Explained

#### Start with Windows
Check this option to automatically launch the app when Windows starts. Perfect for ensuring your microphone volume is always correct.

#### Close to System Tray
When enabled (default), clicking the X button minimizes the app to the system tray instead of closing it. This keeps volume control running in the background.
- **Double-click** the tray icon to restore the window
- **Right-click** for menu options (Show/Exit)

#### Per-Microphone Settings
The app remembers different volume levels for each microphone:
- Set your USB mic to 75%
- Switch to your headset and set it to 40%
- Switch back - each microphone automatically loads its saved volume

## 🎯 Use Cases

- **Streaming/Recording** - Prevent game audio or system sounds from changing mic levels
- **Online Meetings** - Lock your microphone volume for consistent audio quality
- **Gaming** - Stop voice chat applications from auto-adjusting your mic
- **Multiple Microphones** - Different volume levels for different recording setups

## 🛠️ Building from Source

### Prerequisites

- Windows 10/11
- [.NET 9.0 SDK](https://dotnet.microsoft.com/download/dotnet/9.0) or later
- Visual Studio 2022 (optional, for IDE development)

### Build Steps

#### Using Command Line

```bash
# Clone the repository
git clone https://github.com/FJB-ZA/mic-volume-controller.git
cd mic-volume-controller

# Build the project
dotnet build -c Release

# Run the application
dotnet run --project MicVolumeController
```

#### Using Visual Studio

1. Open `MicVolumeController.sln` in Visual Studio
2. Set build configuration to **Release**
3. Build → Build Solution (or press `Ctrl+Shift+B`)
4. Run with `F5` or press the Start button

### Publishing a Single Executable

To create a single-file executable:

```bash
cd MicVolumeController
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -o ./publish
```

The executable will be in the `./publish` folder.

## 📦 Dependencies

- [NAudio](https://github.com/naudio/NAudio) (v2.2.1) - Audio device management and control
- .NET 9.0 Windows Forms - UI framework

## 🤝 Contributing

Contributions are welcome! Here's how you can help:

1. **Fork the repository**
2. **Create a feature branch** (`git checkout -b feature/AmazingFeature`)
3. **Commit your changes** (`git commit -m 'Add some AmazingFeature'`)
4. **Push to the branch** (`git push origin feature/AmazingFeature`)
5. **Open a Pull Request**

### Code Style

- Follow standard C# naming conventions
- Add XML documentation comments for public methods and classes
- Keep methods focused and single-purpose
- Test your changes thoroughly

### Reporting Issues

Found a bug or have a feature request? Please [open an issue](https://github.com/FJB-ZA/mic-volume-controller/issues) with:
- Clear description of the problem/suggestion
- Steps to reproduce (for bugs)
- Your Windows version and .NET version
- Screenshots if applicable

## 📝 Project Structure

```
MicVolumeController/
├── MainForm.cs              # Main application form and logic
├── Program.cs               # Application entry point
├── Properties/
│   └── Settings.settings    # Application settings storage
├── mic_controller_icon.ico  # Application icon
└── MicVolumeController.csproj  # Project configuration
```

## 🔒 Privacy & Security

- **No data collection** - This app does not collect, transmit, or store any personal data
- **Local settings only** - All settings are stored locally on your computer
- **Registry access** - Only used for the "Start with Windows" feature
- **Open source** - All code is available for review

## ❓ FAQ

**Q: Why does Windows SmartScreen block the app?**  
A: The app isn't digitally signed. Click "More info" → "Run anyway" to bypass this warning.

**Q: Does this work with USB microphones?**  
A: Yes! It works with all audio input devices that Windows recognizes.

**Q: Will this interfere with Discord/Teams/Zoom?**  
A: No, the app only sets the system volume level. Application-specific audio processing is unaffected.

**Q: Can I use this on multiple PCs?**  
A: Yes! Just copy the .exe to each computer. Settings are stored per-machine.

**Q: How do I uninstall?**  
A: Simply delete the .exe file. To remove startup entry, uncheck "Start with Windows" first.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 👏 Acknowledgments

- Built with [NAudio](https://github.com/naudio/NAudio) by Mark Heath
- Icon design by Francois
- Inspired by the need for consistent microphone volumes during streaming

## 📧 Contact

Francois - [@FJB-ZA](https://github.com/FJB-ZA)

Project Link: [https://github.com/FJB-ZA/mic-volume-controller](https://github.com/FJB-ZA/mic-volume-controller)

---

⭐ **Star this repository** if you find it helpful!

💡 **Have suggestions?** Open an issue or submit a pull request!
