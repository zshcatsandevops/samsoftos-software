#!/usr/bin/env python3
"""
Samsoft Update Manager CE (Community Edition)
- Inspired by WSUS Offline Update
- Online + Offline update modes
- Local repository for Windows/Office/.NET/VC++ updates
- Tkinter GUI frontend
"""

import sys, os, ctypes, subprocess, threading
import tkinter as tk
from tkinter import ttk, messagebox

# ---------- Auto-elevation ----------
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    params = " ".join([f'"{arg}"' for arg in sys.argv])
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
    sys.exit()

# ---------- Repo Directory ----------
REPO_DIR = os.path.join(os.getcwd(), "SamsoftRepo")
os.makedirs(REPO_DIR, exist_ok=True)

# ---------- GUI Application ----------
class UpdateManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Samsoft Update Manager CE")
        self.root.geometry("800x550")
        self.root.configure(bg="#1e1e1e")

        self.status_var = tk.StringVar(value="Idle.")
        self.create_ui()

    def create_ui(self):
        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True, padx=10, pady=10)

        # Tabs
        self.tab_windows = ttk.Frame(nb)
        self.tab_office = ttk.Frame(nb)
        self.tab_extras = ttk.Frame(nb)

        nb.add(self.tab_windows, text="Windows")
        nb.add(self.tab_office, text="Office")
        nb.add(self.tab_extras, text="Extras")

        # Windows Tab
        ttk.Button(self.tab_windows, text="Check Online", command=self.check_updates).pack(pady=5)
        ttk.Button(self.tab_windows, text="Download to Repo", command=self.download_updates).pack(pady=5)
        ttk.Button(self.tab_windows, text="Install Online", command=self.install_updates).pack(pady=5)
        ttk.Button(self.tab_windows, text="Install Offline (Repo)", command=self.install_offline).pack(pady=5)

        # Office Tab
        ttk.Button(self.tab_office, text="Update Office (C2R)", command=self.update_office).pack(pady=5)

        # Extras Tab
        ttk.Button(self.tab_extras, text="Update .NET", command=self.update_dotnet).pack(pady=5)
        ttk.Button(self.tab_extras, text="Update VC++ Redists", command=self.update_vcredist).pack(pady=5)

        # Log frame
        log_frame = tk.LabelFrame(self.root, text="Update Log", bg="#1e1e1e", fg="white")
        log_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.log_text = tk.Text(log_frame, wrap="word", bg="#0d0d0d", fg="#00ff00",
                                insertbackground="white", font=("Consolas", 10))
        self.log_text.pack(fill="both", expand=True)

        # Status
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief="sunken", anchor="w")
        status_bar.pack(fill="x", side="bottom")

    # ---------- Helpers ----------
    def log(self, msg):
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.status_var.set(msg)

    def run_powershell(self, command):
        completed = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command],
                                   capture_output=True, text=True)
        return completed.stdout.strip(), completed.stderr.strip()

    def ensure_module(self):
        check_cmd = "Get-Module -ListAvailable -Name PSWindowsUpdate"
        out, _ = self.run_powershell(check_cmd)
        if not out.strip():
            self.log("[INFO] Installing PSWindowsUpdate...")
            install_cmd = "Install-PackageProvider -Name NuGet -Force; Install-Module PSWindowsUpdate -Force"
            _, err = self.run_powershell(install_cmd)
            if err:
                self.log(f"[ERROR] {err}")
                return False
        return True

    # ---------- Windows ----------
    def check_updates(self):
        threading.Thread(target=self._check_updates_thread, daemon=True).start()

    def _check_updates_thread(self):
        self.log("[CHECK] Searching online for updates...")
        if not self.ensure_module():
            return
        cmd = "Import-Module PSWindowsUpdate; Get-WindowsUpdate -MicrosoftUpdate"
        out, err = self.run_powershell(cmd)
        if err:
            self.log(f"[ERROR] {err}")
        elif not out.strip():
            self.log("[OK] System is up to date.")
        else:
            self.log("[FOUND]\n" + out)

    def download_updates(self):
        threading.Thread(target=self._download_thread, daemon=True).start()

    def _download_thread(self):
        self.log(f"[DOWNLOAD] Saving updates into {REPO_DIR}...")
        if not self.ensure_module():
            return
        cmd = f"Import-Module PSWindowsUpdate; Get-WindowsUpdate -Download -MicrosoftUpdate -Verbose | Out-File {os.path.join(REPO_DIR, 'updates.txt')}"
        _, err = self.run_powershell(cmd)
        if err:
            self.log(f"[ERROR] {err}")
        else:
            self.log("[DONE] Updates downloaded to repo.")

    def install_updates(self):
        threading.Thread(target=self._install_online_thread, daemon=True).start()

    def _install_online_thread(self):
        self.log("[INSTALL] Installing updates online...")
        if not self.ensure_module():
            return
        cmd = "Import-Module PSWindowsUpdate; Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -AutoReboot"
        _, err = self.run_powershell(cmd)
        if err:
            self.log(f"[ERROR] {err}")
        else:
            self.log("[DONE] Online updates installed.")

    def install_offline(self):
        self.log(f"[OFFLINE] Applying updates from {REPO_DIR}...")
        for file in os.listdir(REPO_DIR):
            if file.endswith(".msu") or file.endswith(".cab") or file.endswith(".exe"):
                full_path = os.path.join(REPO_DIR, file)
                self.log(f"[INSTALL-OFFLINE] Installing {file}...")
                cmd = f'Start-Process "{full_path}" -ArgumentList "/quiet /norestart" -Wait'
                _, err = self.run_powershell(cmd)
                if err:
                    self.log(f"[ERROR] {err}")
        self.log("[DONE] Offline updates applied.")

    # ---------- Office ----------
    def update_office(self):
        self.log("[OFFICE] Running Office Click-to-Run updater...")
        office_path = r'"C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeC2RClient.exe" /update user'
        out, err = self.run_powershell(office_path)
        if err:
            self.log(f"[ERROR] {err}")
        else:
            self.log("[DONE] Office updated.")

    # ---------- Extras ----------
    def update_dotnet(self):
        self.log("[.NET] Installing latest .NET via winget...")
        cmd = "winget install --id Microsoft.DotNet.DesktopRuntime.8 --exact --silent --accept-source-agreements --accept-package-agreements"
        out, err = self.run_powershell(cmd)
        if err:
            self.log(f"[ERROR] {err}")
        else:
            self.log("[DONE] .NET Runtime updated.")

    def update_vcredist(self):
        self.log("[VC++] Installing Visual C++ Redistributables via winget...")
        pkgs = [
            "Microsoft.VCRedist.2015+.x64",
            "Microsoft.VCRedist.2015+.x86"
        ]
        for pkg in pkgs:
            cmd = f"winget install --id {pkg} --exact --silent --accept-source-agreements --accept-package-agreements"
            out, err = self.run_powershell(cmd)
            if err:
                self.log(f"[ERROR] {err}")
            else:
                self.log(f"[DONE] {pkg} installed.")


if __name__ == "__main__":
    root = tk.Tk()
    app = UpdateManagerApp(root)
    root.mainloop()
