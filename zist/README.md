# SF9 Automated Report Card Generator (2026–2027)

A web application that takes an Excel list of students and automatically generates official DepEd SF9 report cards in a **2-student-per-page (2-Up)** format for Polangui General Comprehensive High School (PGCHS).

---

## Features

- **2-Up Landscape Formatting:** Places 2 student report cards side-by-side on each printed sheet (Left Card & Right Card).
- **Tab Naming:** Sheet tabs are labeled (e.g. `1-2 FRONT Navarro - Cedro`, `1-2 BACK Navarro - Cedro`).
- **Student Name Header on BACK Pages:** Learner's name is in `B1` (Left Card) and `Q1` (Right Card).
- **Centered Cell Alignment:** All numeric grades, general averages, and attendance counts are centered inside their boxes.
- **Smart Remarks Auto-Wrapping:** Wraps teacher remarks at 56 characters per line across rows 9–11 (Term 1), 12–13 (Term 2), and 15–16 (Term 3).
- **Separate Front & Back Downloads:** Download **FRONT (.xlsx)**, **BACK (.xlsx)**, or **Both (ZIP Bundle)** with 1 click.

---

## Local Setup & Running

1. **Install Dependencies:**
   ```bash
   pip install -r requirements.txt
   ```
2. **Start the App:**
   - On Windows: Double-click `start.bat`
   - Or in terminal:
     ```bash
     python main.py
     ```
3. Open **http://localhost:8000** in your browser.

---

## Deploying to GitHub & Vercel

### Step 1: Push to GitHub
```bash
git init
git add .
git commit -m "Initial commit - SF9 Generator 2026-2027"
git branch -M main
git remote add origin https://github.com/<your-username>/<your-repo-name>.git
git push -u origin main
```

### Step 2: Deploy on Vercel
1. Go to [vercel.com](https://vercel.com) and click **"Add New Project"**.
2. Select your imported GitHub repository.
3. Vercel will automatically detect `vercel.json` and `requirements.txt`.
4. Click **Deploy**.
