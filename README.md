 InfraCommand (Support Cockpit)

InfraCommand is a lightweight **Support Cockpit** for Windows admins.  
It started as a small launcher for daily support tasks and is evolving into a broader infrastructure command tool.

---

## ✨ Features (v1.1.0)

- ✅ **Baseline Summary** – one-click system health snapshot  
- ✅ **Cache & Reset Tools** – Outlook, Teams, OneDrive cleanup (safe defaults)  
- ✅ **Sysinternals Integration** – Process Explorer, TCPView, RAMMap  
- ✅ **Session Awareness** – RDSH / Horizon detection  
- ✅ **Smart Window Placement** – grid-based quadrant layout  
- ✅ **Config-Driven Menus** – extend with JSON (`grid.config.json`)  

---

## 📸 Screenshots

![Support Cockpit Main Menu](screenshots/main-menu.png)

---

## 🚀 Installation

### Prerequisites

- Windows 10/11  
- Python 3.12 (tested), Python 3.13 (compatible)  
- PowerShell 5.1+ (or PowerShell 7.x)

### Setup

1. **Clone repository:**
```bash
git clone https://github.com/lukaspfisterch/InfraCommand.git
cd InfraCommand
Install dependencies:

bash
Code kopieren
pip install -r requirements.txt
Configure settings:

bash
Code kopieren
copy grid.config.example.json grid.config.json
# Adjust tools and paths as needed
Start:

bash
Code kopieren
python infracommand.py
⚙️ Configuration
Example: grid.config.json
json
Code kopieren
{
  "TOOLS_DIR": "C:\\Support\\Tools",
  "SCRIPTS_DIR": "./scripts",
  "LOCAL_LOG_DIR": "C:\\Support\\Logs\\Cockpit",
  "CENTRAL_LOG_DIR": "C:\\Support\\Logs_CENTRAL\\Cockpit",
  "TOOLS": {
    "Main": [
      {
        "label": "PowerShell",
        "type": "ps1",
        "path": "PowerShell.ps1",
        "elevate": false
      }
    ]
  }
}
Supported Tool Types
ps1 – PowerShell scripts

cmd – CMD commands

exe – Executables

url – Web links

🎮 Usage
Basic Operation
Start tool: Click button

Window positioning: Automatic quadrant placement

UAC: Automatic elevation for admin tools

Keyboard Shortcuts
F11 – Fullscreen

ESC – Close

Ctrl+Q – Quit

🛠 Development
Project Structure
bash
Code kopieren
InfraCommand/
├── infracommand.py          # Main application
├── window_utils.py          # Window management
├── grid.config.example.json # Example configuration
├── scripts/                 # PowerShell scripts
│   ├── PowerShell.ps1
│   ├── CMD.ps1
│   └── ...
├── requirements.txt         # Python dependencies
├── README.md                # This file
└── LICENSE                  # License
Contributing
Fork the repository

Create feature branch (git checkout -b feature/AmazingFeature)

Commit changes (git commit -m 'Add some AmazingFeature')

Push branch (git push origin feature/AmazingFeature)

Create Pull Request

🔮 Vision / Roadmap
InfraCommand started as a small Support Cockpit – but the long-term vision goes further:

State-machine window orchestration – each tool tracked in a defined state

Multi-monitor layouts – distribute monitoring & action windows across screens

Framework-first design – orchestrator module instead of isolated scripts

Future modules – FSLogix intelligence, AI log analysis, team-wide config sharing

📜 License
This project is licensed under the MIT License – see LICENSE for details.

📝 Changelog
v1.1.0
First public release

Baseline Summary (system checks)

Cache/reset tools for Outlook, Teams, OneDrive

Sysinternals integration (Process Explorer, TCPView, RAMMap)

Session-aware window placement

Config-driven menus and clusters