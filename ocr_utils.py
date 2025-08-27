import os, re
from datetime import date
from typing import List, Tuple
from PIL import Image, ImageOps, ImageFilter
import pytesseract

try:
    import cv2
    HAS_CV2 = True
except: HAS_CV2 = False

pytesseract.pytesseract.tesseract_cmd = os.getenv("TESSERACT_CMD", pytesseract.pytesseract.tesseract_cmd)

def _preprocess(img: Image.Image, strong=False) -> Image.Image:
    g = img.convert("L")
    g = ImageOps.autocontrast(g)
    if strong:
        g = g.filter(ImageFilter.UnsharpMask(radius=1.2, percent=220, threshold=2))
        g = g.point(lambda x: 255 if x > 170 else 0, mode="1").convert("L")
    else:
        g = g.filter(ImageFilter.UnsharpMask(radius=1.0, percent=160, threshold=3))
    return g

def _ocr_text(img: Image.Image, psm=6) -> str:
    try:
        return pytesseract.image_to_string(img, config=f"--oem 3 --psm {psm}", lang="kor+eng")
    except: return ""

def _resize(img: Image.Image, max_w=1400) -> Image.Image:
    w,h = img.size
    if w>max_w:
        s=max_w/float(w)
        return img.resize((max_w,int(h*s)))
    return img

R_010 = re.compile(r"(010)[-\s\.]?(\d{3,4})[-\s\.]?(\d{4})")

def _parse_fields(lines: List[str]) -> dict:
    phone,name,addr="","",""
    for i,ln in enumerate(lines):
        m = R_010.search(ln)
        if m:
            phone=f"010-{m.group(2)}-****"
            left = ln[:m.start()].strip()
            k=re.findall(r"[가-힣]{2,8}", left)
            if k: name=k[-1]
            if i+1<len(lines): addr=lines[i+1].strip()
            break
    return {"대여자명":name,"전화번호":phone,"주소":addr}

def _map_model_device(qr_text:str)->Tuple[str,str]:
    raw=(qr_text or "").strip()
    u=re.sub(r"[^A-Z0-9]","",raw.upper())
    MAP={"SM":"심포니","LT":"락티나","S":"스윙","M":"스윙맥스","F":"프리스타일","G":"각시밀","C":"시밀레"}
    m2=re.match(r"^(SM|LT)(\d{2,})$",u)
    if m2: return MAP.get(m2.group(1),"-"), m2.group(2)
    m1=re.match(r"^([SMFLGC])[A-Z0-9]*$",u)
    if m1: return MAP.get(m1.group(1),"-"), raw
    return "-",""

def make_final_entry(qr_text, path:str):
    im=Image.open(path); im=_resize(im,1400)
    t=_ocr_text(_preprocess(im,False),6)
    lines=[ln.strip() for ln in t.splitlines() if ln.strip()]
    parsed=_parse_fields(lines)
    model,device=_map_model_device(qr_text)
    return {
        "출고일": date.today().isoformat(),
        "대여자명":parsed["대여자명"],
        "전화번호":parsed["전화번호"],
        "주소":parsed["주소"],
        "기기번호":device,
        "기종":model
    }

def make_final_entry_fast(qr_text,path:str):
    im=Image.open(path); im=_resize(im,900)
    t=_ocr_text(_preprocess(im,False),6)
    lines=[ln.strip() for ln in t.splitlines() if ln.strip()]
    parsed=_parse_fields(lines)
    model,device=_map_model_device(qr_text)
    return {
        "출고일": date.today().isoformat(),
        "대여자명":parsed["대여자명"],
        "전화번호":parsed["전화번호"],
        "주소":parsed["주소"],
        "기기번호":device,
        "기종":model
    }




