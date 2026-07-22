# 📋 Smart Attendance System (QR Code)

A QR-code-based attendance verification system that replaces manual roll-call with fast, secure check-ins — preventing proxy attendance through unique, validated QR codes tied to individual student records.

[![Quality Gate Status](https://sonarcloud.io/api/project_badges/measure?project=sumitsaxena2716-sys_QR-Code-Attendance-System&metric=alert_status)](https://sonarcloud.io/summary/new_code?id=sumitsaxena2716-sys_QR-Code-Attendance-System)

---

## 🎯 Overview

Traditional roll-call attendance is slow and prone to proxy marking (one student marking attendance for another). This system solves that by:

- Generating a **unique QR code** for each student, tied to their student ID
- Validating each scan against the database in real time to **block duplicate or invalid check-ins**
- Logging attendance instantly with timestamp, reducing manual paperwork

---

## ✨ Features

- 🔐 Unique QR code generation per student
- ⚡ Real-time scan validation (prevents proxy/duplicate attendance)
- 🗄️ SQL-backed attendance records, query-ready for reports
- 🖥️ Simple HTML/CSS interface for check-in and admin view

---

## 🛠️ Tech Stack

| Layer | Technology |
|---|---|
| Backend | Python |
| Frontend | HTML, CSS |
| Database | SQL |

---

## 🚀 Getting Started

### Prerequisites
- Python 3.9+
- MySQL / SQLite (depending on your configured DB)

### Installation

```bash
# Clone the repository
git clone https://github.com/sumitsaxena2716-sys/QR-Code-Attendance-System.git
cd QR-Code-Attendance-System

# Install dependencies
pip install -r requirements.txt

# Run the application
python app.py
```

Then open `http://localhost:5000` (or the port shown in your terminal) in your browser.

> **Note:** Update the `requirements.txt` and run command above to match your actual entry-point file and dependencies.


## 🧠 What I Learned

- Designing unique, tamper-resistant identifiers (QR codes) tied to real-world records
- Building validation logic to close common loopholes in manual systems
- Structuring a SQL schema for fast, query-ready attendance lookups

---

## 📄 License

This project is licensed under the MIT License.

---

## 👤 Author

**Sumit Saxena**
[LinkedIn](https://linkedin.com/in/sumit-saxena-54566b310) · [Email](mailto:sumitsaxena2716@gmail.com) · [GitHub](https://github.com/sumitsaxena2716-sys)