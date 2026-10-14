<div align="center">

# 🚀 MultiProfile Launcher & Window Layout Manager

**مدیریت حرفه‌ای پروفایل‌های مرورگر، بررسی آفلاین سشن‌ها و چیدمان هوشمند پنجره‌ها در ویندوز**  
*A high-performance Windows desktop suite (PySide6 + Win32 DWM) for managing 100+ isolated browser profiles, offline session/cookie inspection, and real-time DWM window tiling.*

[![Platform](https://img.shields.io/badge/Platform-Windows%2010%2F11%20(x64)-0078D6?style=for-the-badge&logo=windows)](https://github.com/)
[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?style=for-the-badge&logo=python&logoColor=white)](https://python.org)
[![UI](https://img.shields.io/badge/UI-PySide6%20(Qt6)-41CD52?style=for-the-badge&logo=qt&logoColor=white)](https://wiki.qt.io/Qt_for_Python)

</div>

---

## ✨ Key Features | قابلیت‌های کلیدی

### 1. 🌐 Multi-Engine Browser Profile Manager (`main.py`)
- **Auto-Discovery Across Chromium & Gecko**: Automatically detects installed browsers (**Chrome, Edge, Brave, Vivaldi, Yandex, Opera, Opera GX, Firefox, Waterfox, Floorp, LibreWolf**) and reads exact profile nicknames from Chromium `Local State` and Firefox/Waterfox `Profile Groups` SQLite databases.
- **Bulk Profile Creator (`➕ Create Profiles…`)**: Create up to 100 isolated profiles in one click with custom naming patterns (`Profile {n}`), automatic avatar spreading, color tags, and one-click rollback/undo.
- **Workspaces & Staggered Bulk Launch**: Launch groups of profiles with configurable stagger delay (`Delay: 2.0 s`) and optional auto-open URL (`🌐 URL on launch…`) so your disk/CPU never freezes.

### 2. 🔍 Offline Background Session & Cookie Inspector
Inspects profile storage (`LevelDB` + `SQLite` + `DPAPI/AES-GCM Cookies` with Chrome 127+ domain-binding support) **in the background without launching any browser window**:
- **Snapp (`🔍 اسنپ`)**: Detects active logins, phone numbers, JWT tokens, and cookies across `snapp.taxi`, `snappfood.ir`, and `snapp.express`. Supports 1-click **`PAPA_...` session code copy** and bulk JSON export (`papayar_bulk_import.json`) for the **Papa Yar** extension.
- **Digikala (`🛒 دیجی‌کالا`)**: Extracts active sessions, customer phone numbers, full names, and tokens from both LevelDB (`persist:DIGIKALA`, UTF-8/UTF-16LE) and encrypted Chromium/Gecko cookies.
- **Universal Domain Checker (`🌐 بررسی هر لینک…`)**: Enter any website URL (Divar, Basalam, Okala, SnappFood, Telegram Web, etc.) to instantly see which profiles have active cookies or LocalStorage for that domain—and **launch all active profiles with one click**.

### 3. 🪟 Window Layout Manager (`window_layout_manager.py`)
Can be opened right inside the launcher header (`🪟 Window Layout Manager`) or run as a standalone app:
- **Live DWM Thumbnails**: Streams real-time Desktop Window Manager (`dwmapi.dll`) thumbnails of every open window inside an interactive preview canvas.
- **5 Tiling Templates**: `Equal Grid`, `Master+Stack`, `Columns`, `Vertical Stack`, and `Custom Grid (Rows × Cols)` with draggable dividers and DWM shadow-offset compensation (`Gap = 0` means true pixel-perfect edge-to-edge tiling).
- **Global Hotkey (`Ctrl + Alt + L`)**: Tile selected windows from anywhere in Windows with smooth 60 FPS cubic-ease animations.

---

## 🛠️ Installation & Usage | نصب و اجرا

```bash
# 1. Clone the repository
git clone https://github.com/YOUR_USERNAME/multiprofile-launcher.git
cd multiprofile-launcher

# 2. Install dependencies
pip install -r requirements.txt

# 3. Run the unified suite
python main.py

# Or run the Window Layout Manager standalone
python window_layout_manager.py
```

### 📦 Building Standalone `.exe` (PyInstaller)
```bash
pyinstaller --noconfirm --onefile --windowed --name MultiProfileLauncher main.py
pyinstaller --noconfirm --onefile --windowed --name WindowLayoutManager window_layout_manager.py
```

---

## 🔒 Security Note | نکته امنیتی
This repository includes a strict `.gitignore` that blocks `config.json`, `session_*.json`, `papayar_bulk_import.json`, and `accounts_list.txt` from ever being committed to Git.
