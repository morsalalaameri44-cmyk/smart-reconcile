
import os, uuid, re, unicodedata
from pathlib import Path
from datetime import timedelta
from typing import Optional
import pandas as pd
from rapidfuzz import fuzz
from fastapi import FastAPI, Request, UploadFile, File, Form, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

BASE = Path(__file__).parent
UPLOADS = BASE / "uploads"; REPORTS = BASE / "reports"
UPLOADS.mkdir(exist_ok=True); REPORTS.mkdir(exist_ok=True)

app = FastAPI(title="SmartReconcile")
app.mount("/static", StaticFiles(directory=BASE/"static"), name="static")
templates = Jinja2Templates(directory=str(BASE/"templates"))

def normalize_text(v):
    if pd.isna(v): return ""
    s = str(v).strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    trans = str.maketrans({"أ":"ا","إ":"ا","آ":"ا","ى":"ي","ة":"ه"})
    s = s.translate(trans)
    s = re.sub(r"[^0-9a-zA-Z\u0600-\u06FF ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def parse_amount(v):
    if pd.isna(v): return None
    s = str(v).replace(",", "").replace(" ", "")
    s = re.sub(r"[^\d\.\-]", "", s)
    try: return round(float(s), 2)
    except: return None

def parse_date(v):
    if pd.isna(v): return pd.NaT
    return pd.to_datetime(v, errors="coerce", dayfirst=True)

ALIASES = {
    "date": ["date","transaction date","التاريخ","تاريخ","تاريخ العملية"],
    "customer": ["customer","customer name","client","name","اسم العميل","العميل","الاسم"],
    "user": ["user","cashier","operator","المستخدم","الموظف"],
    "amount": ["amount","value","total","المبلغ","القيمة","دائن","مدين","debit","credit"],
    "description": ["description","details","narration","البيان","الوصف","التفاصيل"],
}

def find_col(cols, aliases):
    normalized = {normalize_text(c): c for c in cols}
    for a in aliases:
        na = normalize_text(a)
        for nc, original in normalized.items():
            if na == nc or na in nc or nc in na:
                return original
    return None

def read_file(path: Path):
    ext = path.suffix.lower()
    if ext in [".xlsx", ".xls", ".csv"]:
        if ext == ".csv":
            df = pd.read_csv(path)
        else:
            df = pd.read_excel(path)
        return df
    if ext == ".pdf":
        from pypdf import PdfReader
        text = "\n".join((p.extract_text() or "") for p in PdfReader(str(path)).pages)
        lines = [x.strip() for x in text.splitlines() if x.strip()]
        return pd.DataFrame({"raw_text": lines})
    if ext in [".jpg",".jpeg",".png",".webp"]:
        try:
            import pytesseract
            from PIL import Image
            text = pytesseract.image_to_string(Image.open(path), lang="ara+eng")
            lines = [x.strip() for x in text.splitlines() if x.strip()]
            return pd.DataFrame({"raw_text": lines})
        except Exception as e:
            raise HTTPException(400, "تعذر تنفيذ OCR. ثبّت Tesseract وبيانات اللغة العربية.")
    raise HTTPException(400, "صيغة الملف غير مدعومة في النسخة الأولية.")

def standardize(df):
    # If the file is unstructured, preserve it for manual mapping instead of inventing transactions.
    cols = list(df.columns)
    mapping = {k: find_col(cols, v) for k,v in ALIASES.items()}
    if not mapping["date"] or not mapping["amount"]:
        raise HTTPException(
            400,
            f"تعذر التعرف تلقائياً على أعمدة التاريخ والمبلغ. الأعمدة الموجودة: {', '.join(map(str, cols))}"
        )
    out = pd.DataFrame()
    out["date"] = df[mapping["date"]].map(parse_date)
    out["customer"] = df[mapping["customer"]] if mapping["customer"] else ""
    out["user"] = df[mapping["user"]] if mapping["user"] else ""
    out["description"] = df[mapping["description"]] if mapping["description"] else ""
    out["amount"] = df[mapping["amount"]].map(parse_amount)
    out["customer_norm"] = out["customer"].map(normalize_text)
    out["user_norm"] = out["user"].map(normalize_text)
    out = out.dropna(subset=["date","amount"]).reset_index(drop=True)
    return out

def reconcile(company, bank):
    used = set()
    results = []
    review = []
    for ci, c in company.iterrows():
        candidates = []
        for bi, b in bank.iterrows():
            if bi in used: continue
            if abs(float(c["amount"]) - float(b["amount"])) > 0.001: continue
            days = abs((c["date"].date() - b["date"].date()).days)
            if days > 1: continue
            name_score = fuzz.token_set_ratio(c["customer_norm"], b["customer_norm"]) if c["customer_norm"] and b["customer_norm"] else 0
            user_score = fuzz.ratio(c["user_norm"], b["user_norm"]) if c["user_norm"] and b["user_norm"] else 0
            score = 50 + (20 if days == 0 else 15) + name_score*0.25 + user_score*0.05
            candidates.append((score, bi, days, name_score, user_score))
        candidates.sort(reverse=True)
        if candidates:
            score, bi, days, ns, us = candidates[0]
            b = bank.loc[bi]
            # exact amount is mandatory; accept automatic match when name is strong,
            # or when exact date and only one candidate.
            status = "MATCHED" if (ns >= 80 or (days == 0 and len(candidates) == 1 and ns >= 55)) else "REVIEW_REQUIRED"
            if status == "MATCHED":
                used.add(bi)
            else:
                review.append({"company_index":ci, "bank_index":bi})
            results.append({
                "company_date": c["date"], "company_customer":c["customer"], "company_user":c["user"], "company_amount":c["amount"],
                "bank_date":b["date"], "bank_customer":b["customer"], "bank_user":b["user"], "bank_amount":b["amount"],
                "status":status, "score":round(score,1), "date_difference_days":days,
                "name_similarity":round(ns,1), "user_similarity":round(us,1)
            })
        else:
            results.append({
                "company_date":c["date"],"company_customer":c["customer"],"company_user":c["user"],"company_amount":c["amount"],
                "bank_date":None,"bank_customer":"","bank_user":"","bank_amount":None,
                "status":"MISSING_IN_BANK","score":0,"date_difference_days":None,
                "name_similarity":0,"user_similarity":0
            })
    for bi, b in bank.iterrows():
        if bi not in used:
            # banks referenced by review are not automatically considered matched
            if any(x["bank_index"] == bi for x in review): 
                continue
            results.append({
                "company_date":None,"company_customer":"","company_user":"","company_amount":None,
                "bank_date":b["date"],"bank_customer":b["customer"],"bank_user":b["user"],"bank_amount":b["amount"],
                "status":"MISSING_IN_COMPANY","score":0,"date_difference_days":None,
                "name_similarity":0,"user_similarity":0
            })
    return pd.DataFrame(results)

@app.get("/", response_class=HTMLResponse)
def login(request: Request):
    return templates.TemplateResponse("login.html", {"request":request})

@app.post("/login")
def do_login(username: str = Form(...), password: str = Form(...)):
    # Demo authentication for MVP; replace with hashed users in production.
    if username == "admin" and password == "admin123":
        return RedirectResponse("/dashboard", status_code=303)
    return RedirectResponse("/?error=1", status_code=303)

@app.get("/dashboard", response_class=HTMLResponse)
def dashboard(request: Request):
    return templates.TemplateResponse("dashboard.html", {"request":request})

@app.post("/reconcile", response_class=HTMLResponse)
async def do_reconcile(request: Request, bank_file: UploadFile = File(...), company_file: UploadFile = File(...)):
    job_id = uuid.uuid4().hex[:12]
    def save(upload, prefix):
        ext = Path(upload.filename).suffix
        p = UPLOADS / f"{job_id}_{prefix}{ext}"
        with open(p, "wb") as f: f.write(upload.file.read())
        return p
    bp = save(bank_file, "bank"); cp = save(company_file, "company")
    bank = standardize(read_file(bp))
    company = standardize(read_file(cp))
    result = reconcile(company, bank)
    report = REPORTS / f"reconciliation_{job_id}.xlsx"
    with pd.ExcelWriter(report, engine="openpyxl") as writer:
        summary = result["status"].value_counts().rename_axis("status").reset_index(name="count")
        summary.to_excel(writer, sheet_name="Summary", index=False)
        for status, sheet in [
            ("MATCHED","Matched"),("REVIEW_REQUIRED","Manual Review"),
            ("MISSING_IN_BANK","Missing in Bank"),("MISSING_IN_COMPANY","Missing in Company")
        ]:
            result[result.status == status].to_excel(writer, sheet_name=sheet, index=False)
        result.to_excel(writer, sheet_name="All Results", index=False)
    counts = result["status"].value_counts().to_dict()
    return templates.TemplateResponse("results.html", {
        "request":request, "job_id":job_id, "counts":counts,
        "rows":result.fillna("").to_dict("records")
    })

@app.get("/download/{job_id}")
def download(job_id: str):
    files = list(REPORTS.glob(f"reconciliation_{job_id}.xlsx"))
    if not files: raise HTTPException(404, "التقرير غير موجود")
    return FileResponse(files[0], filename=files[0].name)
