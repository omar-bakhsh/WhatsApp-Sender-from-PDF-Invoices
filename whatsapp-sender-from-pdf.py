import os
import re
import sys
import urllib.parse
from pathlib import Path
import pandas as pd
import subprocess
import pytesseract
from PIL import Image
import io

# مسار tesseract
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# مكتبات
try:
    from pdf2image import convert_from_path
    import fitz  # pymupdf
except Exception:
    print("⚠️ ثبت الحزم المطلوبة: pip install pdf2image pytesseract pillow pandas openpyxl pymupdf")
    sys.exit(1)

# إعدادات (عدّل المسارات حسب جهازك)
PDF_FOLDER = r"C:\DF_Files"   # مجلد الفواتير
OUTPUT_XLSX = r"C:\Users\hp\OneDrive\سطح المكتب\pdf-whats\whatsapp_links.xlsx"
GOOGLE_REVIEW_LINK = "https://g.page/r/YOUR_GOOGLE_REVIEW_LINK"
POPPLER_BIN = r"C:\Users\hp\Downloads\Release-25.07.0-0\poppler-25.07.0\bin"

# أرقام المركز لتجاهلها
CENTER_NUMBERS = {"0566522351", "0556565135"}

# أنماط
PHONE_REGEX = re.compile(r'0?5\d{8}')
ARABIC_LETTERS = r'\u0600-\u06FF'
# كلمات مفتاحية للعميل
CLIENT_KEYWORDS = [
    r'اسم\s*العميل', r'الى\s*المكرم', r'إلى\s*المكرم', r'المكرم',
    r'Mob\.?No', r'الجوال', r'رقم\s*الجوال', r'جوال'
]
NAME_KEYWORDS = [
    r'اسم\s*العميل', r'إلى\s*المكرم', r'الى\s*المكرم', r'المكرم'
]

def normalize_text_for_search(text: str) -> str:
    """
    يفصل الكلمات الملتصقة مثل '556464353Mob.Noمازن' إلى '556464353 Mob.No مازن'
    ويبدل علامات غير قياسية بمسافات.
    """
    if not text:
        return ""
    # استبدال أنواع Mob.No المنوعة بمسافة مفصولة
    text = re.sub(r'(Mob\.?No)', r' \1 ', text, flags=re.IGNORECASE)
    # ضع مسافة بين رقم يتبعه حرف عربي مباشرة (مثال: '556464353مازن' -> '556464353 مازن')
    text = re.sub(r'(\d)([ء-ي])', r'\1 \2', text)
    # وضع مسافة بين حرف عربي يتبعه رقم (مثال: 'مازن556' -> 'مازن 556')
    text = re.sub(r'([ء-ي])(\d)', r'\1 \2', text)
    # استبدال علامات خاصة بمسافة
    text = re.sub(r'[_\-\|,:/()\[\]]+', ' ', text)
    # أضع مسافة حول ':' و '-' و '/'
    text = re.sub(r'\s{2,}', ' ', text)
    return text.strip()

def clean_name_candidate(s: str) -> str:
    """
    ينظف المرشح للاسم: يزيل الأرقام، كلمات Mob.No، 'رقم'، 'جوال'، ويفصل زوائد.
    يعيد 'غير معروف' إذا لم يوجد اسم عربي واضح.
    """
    if not s:
        return "غير معروف"
    s = s.strip()
    # استبدال كلمات غير مرغوبة
    s = re.sub(r'(?i)Mob\.?No', ' ', s)
    s = re.sub(r'(?i)رقم\s*الجوال|رقم|جوال|MobNo', ' ', s)
    # إزالة أرقام ورموز
    s = re.sub(r'[0-9]', ' ', s)
    s = re.sub(r'[_\-\|,:\.\(\)]', ' ', s)
    s = re.sub(r'\s{2,}', ' ', s).strip()
    # الآن نريد أن نأخذ أول سلسلة عربية طويلة بما يكفي (مثلاً كلمتين أو أكثر)
    m = re.search(rf'([ء-ي]+(?:\s+[ء-ي]+)+)', s)
    if m:
        name = m.group(1).strip()
        return name
    # لو لم نجد سلسلة من كلمتين، خذ أول كلمة عربية مفيدة
    m2 = re.search(rf'([ء-ي]{{2,}})', s)
    if m2:
        return m2.group(1).strip()
    return "غير معروف"

def find_candidate_phone(text: str):
    """
    يبحث عن رقم قريب من كلمات العميل أولاً، ثم أي رقم مطابق للـ pattern.
    يعيد الرقم كسلسلة (بدون + أو مسافات).
    """
    if not text:
        return None
    txt = normalize_text_for_search(text)
    # بحث قرب الكلمات المفتاحية
    for kw in CLIENT_KEYWORDS:
        # نأخذ 0?5xxxxxxxx قرب الكلمة
        pattern = re.compile(rf'({kw}).{{0,60}}(0?5\d{{8}})', re.IGNORECASE)
        m = pattern.search(txt)
        if m:
            return m.group(2)
        # العكس: رقم ثم الكلمة بعده
        pattern2 = re.compile(rf'(0?5\d{{8}}).{{0,60}}({kw})', re.IGNORECASE)
        m2 = pattern2.search(txt)
        if m2:
            return m2.group(1)
    # لو لم نجد قرب الكلمات، نبحث عن أول رقم مطابق
    m3 = PHONE_REGEX.search(txt)
    if m3:
        return m3.group(0)
    return None

def find_name(text: str, phone_found: str = None):
    """
    يستخرج اسم العميل بناءً على كلمات مفتاحية أو قرب رقم الهاتف المستخرج.
    phone_found يمرر إذا وُجد للمساعدة في تحديد موقع الاسم.
    """
    if not text:
        return "غير معروف"
    txt = normalize_text_for_search(text)

    # 1) محاولة استخراج الاسم مباشرة بعد كلمات الاسم
    for kw in NAME_KEYWORDS:
        pattern = re.compile(rf'{kw}\s*[:\-]?\s*([ء-ي0-9\s\-]+)', re.IGNORECASE)
        m = pattern.search(txt)
        if m:
            candidate = m.group(1).strip()
            cleaned = clean_name_candidate(candidate)
            if cleaned != "غير معروف":
                return cleaned

    # 2) إذا وُجد رقم الهاتف، حاول أخذ نص قريب (قبل أو بعد) الرقم
    if phone_found:
        # تأكد من أن phone_found موجود في النص بعد التطبيع
        ph = phone_found
        # ابحث عن ph في النص واحصل على الناحية القريبة (50 حرفًا)
        loc = txt.find(ph)
        if loc != -1:
            # احصل على نافذة صغيرة قبل وبعد الرقم
            start = max(0, loc - 60)
            end = loc + len(ph) + 60
            window = txt[start:end]
            # حاول إيجاد سلسلة عربية في النافذة
            m = re.search(rf'([ء-ي]+(?:\s+[ء-ي]+)+)', window)
            if m:
                cleaned = clean_name_candidate(m.group(1))
                if cleaned != "غير معروف":
                    return cleaned

    # 3) كخيار أخير: خذ أول سلسلة عربية من النص (اسم محتمل)
    m = re.search(rf'([ء-ي]+(?:\s+[ء-ي]+)+)', txt)
    if m:
        return clean_name_candidate(m.group(1))

    return "غير معروف"

# بقية الكود (استخراج صفحات PDF -> OCR -> تجميع النتائج)
def build_whatsapp_link(name, phone):
    if not phone:
        return ""
    if phone.startswith("5"):
        phone = "0" + phone
    phone_intl = "966" + phone[1:]
    message = f"مرحباً {name} 👋، نشكرك على زيارتك لمركز مازدا ونأمل تقييم خدمتنا في جوجل 🌟\n\nرابط التقييم: {GOOGLE_REVIEW_LINK}"
    encoded = urllib.parse.quote(message)
    return f"https://wa.me/{phone_intl}?text={encoded}"

def ocr_pdf_and_extract(pdf_path: Path):
    results = []
    pages = []
    # محاولة Poppler/pdf2image
    try:
        pages = convert_from_path(str(pdf_path), dpi=200, fmt='png', poppler_path=POPPLER_BIN)
    except Exception as e:
        # سنحاول PyMuPDF كبديل
        try:
            import fitz
            doc = fitz.open(str(pdf_path))
            for page_index in range(len(doc)):
                page = doc.load_page(page_index)
                pix = page.get_pixmap(dpi=200)
                img = Image.open(io.BytesIO(pix.tobytes("png")))
                pages.append(img)
            doc.close()
        except Exception as e2:
            print(f"⚠️ لم يتمكن أي محول من فتح {pdf_path.name}: {e2}")
            return results

    for i, page in enumerate(pages, start=1):
        text = ""
        try:
            text = pytesseract.image_to_string(page, lang='ara+eng')
        except Exception:
            try:
                text = pytesseract.image_to_string(page)
            except Exception:
                text = ""
        # نظّف النص مؤقتًا
        ntext = normalize_text_for_search(text)
        phone = find_candidate_phone(ntext)
        name = find_name(ntext, phone)
        # تخطي أرقام المركز
        if phone and phone in CENTER_NUMBERS:
            phone = None
        results.append({
            "invoice_file": pdf_path.name,
            "page": i,
            "name": name,
            "phone": phone or ""
        })
    return results

def main():
    pdf_folder = Path(PDF_FOLDER)
    if not pdf_folder.exists():
        print("المجلد غير موجود:", pdf_folder)
        return
    all_rows = []
    for pdf_path in pdf_folder.rglob("*.pdf"):
        print("📄 معالجة:", pdf_path.name)
        items = ocr_pdf_and_extract(pdf_path)
        if not items:
            # ملف فُتح لكن لم يعط أي صفحة/نص
            all_rows.append({
                "اسم العميل": "غير معروف",
                "رقم الجوال": "",
                "رابط واتساب": "",
                "ملف الفاتورة": pdf_path.name,
                "صفحة": ""
            })
            continue
        for it in items:
            wa = build_whatsapp_link(it["name"], it["phone"])
            all_rows.append({
                "اسم العميل": it["name"],
                "رقم الجوال": it["phone"],
                "رابط واتساب": wa,
                "ملف الفاتورة": it["invoice_file"],
                "صفحة": it["page"]
            })
    if not all_rows:
        print("لم توجد نتائج.")
        return
    df = pd.DataFrame(all_rows)
    # اجعل الصفوف التي لديها أرقام أولًا بدون تكرار، ثم الباقي
    df_non_empty = df[df["رقم الجوال"].astype(bool)].drop_duplicates(subset=["رقم الجوال"])
    df_empty = df[~df["رقم الجوال"].astype(bool)]
    final_df = pd.concat([df_non_empty, df_empty], ignore_index=True)
    final_df.to_excel(OUTPUT_XLSX, index=False)
    print("✅ تم إنشاء:", OUTPUT_XLSX)
    try:
        subprocess.Popen(["start", OUTPUT_XLSX], shell=True)
    except Exception:
        pass

if __name__ == "__main__":
    main()
