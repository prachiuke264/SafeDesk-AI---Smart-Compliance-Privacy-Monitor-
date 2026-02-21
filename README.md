SafeDesk AI - Smart Compliance & Privacy Monitor 🛡️
SafeDesk AI is a lightweight, edge-computing desktop application designed to enforce "Clean Desk Policies" in secure environments like BPOs, Banks, and Exam Centers. It uses YOLOv8 to detect unauthorized mobile devices in real-time and takes automated security actions.

🚀 Key Features
Real-time AI Detection: Powered by YOLOv8 for high-speed mobile phone detection.
Privacy-First Design: The live camera feed is hidden from the UI to ensure employee comfort; only AI scans the background.
Automated Security: Instantly locks the Windows Workstation (Win + L) upon violation.
Audit Logging: Saves every violation with a timestamp, employee ID (Windows Login), and high-res photo evidence in an SQLite database.
Compliance Reporting: One-click export to Excel (.xlsx) for HR and Management audits.
Stealth & Persistence: Supports "Minimize to System Tray" and "Auto-start with Windows" via Registry integration.

Language: Python 3.10+
AI Engine: Ultralytics YOLOv8
GUI Framework: CustomTkinter (Modern Dark Theme)
Database: SQLite3
Data Export: Pandas & OpenPyxl
Packaging: PyInstaller (Executable (.exe) for Windows)

📦 Installation & Setup
1.Clone the repo
git clone https://github.com
cd SafeDesk-AI

2.Install dependencies
pip install -r requirements.txt

3.Run the Application
python safedesk_final.py

🖥️ How It Works
Start Monitoring: Activates the webcam in the background.
Detection: If a "Cell Phone" is identified with >40% confidence, a "Violation" is triggered.
Action: The app captures a screenshot, logs it to compliance_logs.db, and calls the Windows LockWorkStation API.
Reporting: Managers can click Export to Excel to get a full audit trail of all employees.






