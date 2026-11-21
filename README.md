# Protons Engineering - Daily Service Sheet & Excel Generator

A full-stack daily service tracking app deployed on Netlify. Uses a mobile-friendly frontend with client-side localStorage persistence and a serverless Netlify Function (`/api/generate-excel`) to generate branded `.xlsx` sheets matching the template style.

---

## 3 Quick Steps

### 1. How to Edit CONFIG
All sheet settings, headings, brand names, colors, and input rules live in a single place at the top of `index.html`:

```javascript
const CONFIG = {
  title: "Protons Engineering",
  company: "Company:Widecare",
  sheetName: "Client Data",
  fileNamePattern: "Protons-Widecare-{date}",
  excelColors: { title: "D35400", company: "E67E22", header: "B03A2E" },
  summarySuffix: " Calls",
  includeTotalsInExcel: true,
  statusCounters: [
    { label: "Gas Charging", match: "gas charg" }
  ],
  columns: [
    { key: "Claim ID",       required: true, width: 15.66, inputmode: "numeric", placeholder: "e.g. 35431" },
    { key: "Customer Name",  required: true, width: 25.66 },
    { key: "Area",           width: 20.66, suggest: ["Virar (W)", "Virar (E)", "Vasai (W)", ...] },
    { key: "Phone Number",   width: 18, inputmode: "tel", type: "tel", placeholder: "10 digit number" },
    { key: "Product Name",   required: true, width: 25.66, role: "product", chips: ["A/C", "Ref", "W/M"] },
    { key: "Product Status", width: 28, role: "status", chips: ["Jet Service done", "Gas charging done", ...] }
  ]
};
```
- **Change Headings**: Change any `key` (e.g. change `"Area"` to `"Location"`). The form, cards, and downloaded Excel update automatically with zero backend changes needed.
- **Change Colors or Title**: Edit `title`, `company`, or `excelColors`.
- **Add Quick Chips or Suggestions**: Add or remove strings from `chips` or `suggest`.

---

### 2. How to Test Locally
Make sure Node.js is installed, then in this project folder:

```bash
# 1. Install dependencies
npm install

# 2. Start local Netlify development server (serves frontend & serverless functions)
npm run dev
# or: npx netlify dev
```

- Open the local URL (usually `http://localhost:8888`).
- Enter sample rows and click **Download Excel**.
- The backend function `/api/generate-excel` handles the generation. If offline, the built-in browser fallback generates the file.

---

### 3. How to Deploy to Netlify

#### Option A: Connect via GitHub / GitLab (Recommended for Beginners)
1. Push this folder to a GitHub repository.
2. Log in to [Netlify](https://app.netlify.com/).
3. Click **Add new site** > **Import an existing project** > Choose **GitHub**.
4. Select your repository.
5. In the build settings:
   - **Base directory**: (leave blank / root)
   - **Build command**: (leave blank or `npm install`)
   - **Publish directory**: `.` (or leave as configured in `netlify.toml`)
   - **Functions directory**: `netlify/functions` (detected automatically from `netlify.toml`)
6. Click **Deploy site**.
7. Netlify will deploy your site and serverless functions automatically.
8. Your site will be live at:
   `https://<your-site-name>.netlify.app`

#### Option B: Deploy from Terminal with Netlify CLI
1. Log in to your Netlify account:
   ```bash
   npx netlify login
   ```
2. Link or create a new site:
   ```bash
   npx netlify init
   ```
3. Deploy directly to production:
   ```bash
   npx netlify deploy --prod
   ```

---

## Architecture & Security Notes
- **Stateless Backend**: The function receives the payload, generates the Excel binary in memory, sends it back as base64, and retains no state.
- **Privacy & Safety**: Customer names and phone numbers are never logged or stored.
- **Formula Injection Prevention**: Leading characters `=`, `+`, `-`, `@` are escaped with a leading apostrophe (`'`).
- **Input Bounds**: Capped at 500 rows, 12 columns, 200 chars per cell, and ~1MB payload limit.
- **Offline Resilient**: If the backend is unreachable, the frontend automatically falls back to in-browser generation with ExcelJS.
