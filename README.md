# VAPT One-Click Report Generator — Updated Version

## 📁 Folder Structure (jo bhi files honi chahiye)

```
VAPT_OneClick/
├── app.py                              ← Flask backend (Routes + DOCX report generator + Knowledge Base)
├── requirement.txt                     ← Python dependencies (pinned versions)
├── run.bat                             ← Windows par ek-click run karne ke liye
├── instance/
│   └── vapt_report.db                  ← SQLite DB (aapke saved vulnerabilities + templates)
├── static/
│   └── branding/
│       ├── header_logo_default.png     ← ProTechmanize logo (har report me by default)
│       └── watermark_default.png       ← ProTechmanize watermark (har report me by default)
└── templates/
    └── index.html                      ← Web UI (form, autocomplete, preview)
```

Yeh poora folder as-is apne existing `VAPT_OneClick` folder ki jagah use karo (ya overwrite kar do — `instance/vapt_report.db` me aapka purana data already maujood hai, naya overwrite nahi karega use jab tak aap khud naya data add na karo).

## 🛠️ Setup (pehli baar)

```bash
cd VAPT_OneClick
pip install -r requirement.txt
python app.py
```

Ya seedha `run.bat` double-click karo (Windows par).

Browser me kholo: **http://127.0.0.1:5000**

---

## ✅ Kya Naya Hai (is update me)

### 1. Watermark — Image-based, Centered, Faded, DEFAULT (koi upload nahi)
- Purana text watermark ("PROTECH CONFIDENTIAL") completely hata diya gaya
- ProTechmanize ka watermark `static/branding/watermark_default.png` se **automatically** lagta hai har report par — page ke exact center me faded/washed-out
- Koi upload field nahi — yeh fixed branding hai, har report me default

### 2. Header — Logo + Client Name + Date (left-aligned, exact reference jaisa)
- **Logo**: top-left, har page par — `static/branding/header_logo_default.png` se automatically (koi upload nahi)
- **Client Name**: logo ke turant baad, same line par (left-aligned, center nahi)
- **Date**: client name ke neeche, agli line par (left-aligned)
- **"ProTechmanize Solutions Pvt Ltd" jaisa koi firm-name text header me NAHI hai** — sirf logo + client name + date
- Yeh exact alignment aapke diye gaye reference reports (Web_Initial_VAPT_Updated_Report.docx, Powerapp_kotak_amc.docx, Kubernates_API_Prod-BusinessHub.docx) se match karta hai

### 3. Page Border, Footer, Tables
- Dark navy border har page par
- Footer: CERT-IN line + website + live "Page X of Y"
- Saari tables: dark blue header row, white bold text, thin black borders
- Severity colors (Critical/High/Medium/Low/Informational) har jagah consistent

### 4. Real Word Table of Contents
- Ab ek genuine Word TOC field hai (`TOC \o "1-3" \h \z \u`)
- Jab document Word me khulta hai, TOC apne aap update ho jaata hai

### 5. 🧠 Auto Vulnerability Knowledge Base (NAYA — aapne jo maanga tha)

**Workflow:**

```
Day 1: "Misconfigured CSP Header" ko Company A ke liye manually fill kiya
       → Severity, CWE/CVE, Description, Business Impact, Remediation, References

       → Add Vulnerability dabaya
            ↓
       System automatically isko TEMPLATE ki tarah bhi save kar leta hai
       (alag se kuch karne ki zaroorat nahi)

Day 2: Company B ke kisi naye URL pe wahi "Misconfigured CSP" mila
       → "Vulnerability Name" field me type karna shuru kiya: "Misco..."
            ↓
       Live dropdown turant dikhata hai: "Misconfigured CSP Header (Low)"
            ↓
       Click kiya YA poora naam type karke field se bahar click kiya (blur)
            ↓
       Severity, CWE/CVE, Description, Business Impact, Remediation,
       References — SAB AUTOMATICALLY BHAR JAATE HAIN
            ↓
       Bas naya "Vulnerable URL" aur POC daalo, baaki kaam khatam
```

- **Pehli baar** koi naya vulnerability naam save hota hai → naya template ban jaata hai
- **Dusri baar** wahi naam (kisi bhi company ke liye) save hota hai → existing template **refresh** ho jaata hai latest details ke saath (knowledge base time ke saath better hoti jaati hai)
- Matching case-insensitive hai — "misconfigured csp header" aur "Misconfigured CSP Header" dono same template treat honge
- Naam type karte hi live suggestions dikhte hain (jaise "csp" type karne par "Misconfigured CSP Header" dikhega, "click" type karne par "Clickjacking" dikhega)
- Suggestion par click karo YA poora exact naam type karke bahar click karo — dono se auto-fill ho jaata hai

Yeh knowledge base **10 built-in templates** ke saath shuru hoti hai (Default Credentials, Stored XSS, HTML Injection, Clickjacking, CORS, Missing Security Headers, Weak Ciphers, Improper Error Handling, Unrestricted Internal Service Access, Misconfigured CSP Header), aur jaise-jaise aap naye vulnerabilities add karte ho, yeh list khud badhti jaati hai — permanently, SQLite DB me.

---

## 🧪 Testing

Yeh code real Flask routes ke through, pinned dependency versions (`python-docx==0.8.11`, `flask==2.3.2`) ke against, end-to-end test kiya gaya hai:
- Naya vulnerability add karna → DOCX report generate hona
- Brand-new company, **zero uploads** → bundled default ProTechmanize logo + watermark automatically apply hote hain
- Knowledge base: create → search → exact-match fetch → refresh, sab verified

---

## ⚠️ Note

`instance/vapt_report.db` me aapka **original data already hai** — is zip ko extract karte waqt agar pucha jaaye to overwrite na karo agar aap apna purana data rakhna chahte ho. Agar fresh start chahiye, to is DB file ko delete kar do — app khud naya bana lega (sirf 10 built-in templates ke saath).
