from PIL import Image
import pytesseract
import re
from datetime import datetime

def extract_shipping_info(image_path):
    image = Image.open(image_path)
    text = pytesseract.image_to_string(image, lang='kor+eng')

    name = None
    phone = None
    address = None
    invoice = None

    lines = text.splitlines()
    for i, line in enumerate(lines):
        if not phone:
            phone_match = re.search(r'010[-\s]?\d{3,4}[-\s]?\d{4}', line)
            if phone_match:
                phone = phone_match.group().replace(' ', '').replace('--', '-')
                name = lines[i - 1].strip() if i > 0 else None
                address = lines[i + 1].strip() if i + 1 < len(lines) else None

        if not invoice:
            invoice_match = re.search(r'\b\d{4}[-]?\d{4}[-]?\d{4}\b', line)
            if invoice_match:
                invoice = invoice_match.group().replace('-', '')

    return {
        "수취인명": name or "",
        "전화번호": phone or "",
        "주소": address or "",
        "송장번호": invoice or "",
        "출고일": datetime.now().strftime("%Y-%m-%d")
    }

def parse_qr_text(qr_text):
    code_map = {
        "SM": "심포니", "LT": "락티나", "SW": "스윙", "MX": "스윙맥스",
        "FR": "프리스타일", "SP": "스펙트라", "GS": "각시밀", "CM": "시밀레",
    }
    prefix = qr_text[:2]
    return {
        "기종": code_map.get(prefix, "알 수 없음"),
        "기기번호": qr_text[2:]
    }

def make_final_entry(qr_text, 송장_image_path):
    qr_data = parse_qr_text(qr_text)
    송장_data = extract_shipping_info(송장_image_path)

    return {
        "출고일": 송장_data["출고일"],
        "대여자명": 송장_data["수취인명"],
        "전화번호": 송장_data["전화번호"],
        "주소": 송장_data["주소"],
        "기기번호": qr_data["기기번호"],
        "기종": qr_data["기종"],
        "송장번호": 송장_data["송장번호"],
    }
