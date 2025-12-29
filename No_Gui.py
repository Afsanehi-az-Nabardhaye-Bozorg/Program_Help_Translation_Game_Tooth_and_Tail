import pandas as pd
import re
import arabic_reshaper
from bidi.algorithm import get_display

def process_excel_to_xml(excel_path, xml_path, output_xml_path):
    # مرحله 1: خواندن فایل اکسل و ایجاد دیکشنری تنظیمات
    excel_data = pd.read_excel(excel_path, header=None)
    
    # ایجاد دیکشنری برای ذخیره IDها و تنظیمات مربوطه
    id_settings = {}
    valid_ids = set()
    
    # پر کردن دیکشنری تنظیمات از ستون‌های 4,5,6
    for idx, row in excel_data.iterrows():
        if len(row) > 6 and pd.notna(row[4]) and pd.notna(row[5]) and pd.notna(row[6]):
            try:
                id_val = int(row[4])
                id_settings[id_val] = {
                    'min_len': int(row[5]),
                    'max_len': int(row[6])
                }
                valid_ids.add(id_val)
            except:
                continue
    
    # ایجاد دیکشنری برای نگاشت کلید به ID از ستون‌های 0 و 3
    key_to_id = {}
    key_to_text = {}
    for idx, row in excel_data.iterrows():
        if len(row) > 3 and pd.notna(row[0]) and pd.notna(row[3]):
            key_to_id[row[0]] = row[3]
        if len(row) > 1 and pd.notna(row[0]) and pd.notna(row[1]):
            key_to_text[row[0]] = row[1]
    
    # مرحله 2: پردازش فایل XML
    with open(xml_path, 'r', encoding='utf-8') as f:
        xml_content = f.read()
    
    # الگو برای یافتن آیتم‌ها
    item_pattern = re.compile(r'<item\s+key="([^"]+)"\s+value="([^"]*)"\s*/>')
    
    # تابع جایگزینی
    def replace_item(match):
        key = match.group(1)
        original_value = match.group(2)
        
        # بررسی آیا کلید در اکسل وجود دارد
        if key in key_to_id:
            id_val = key_to_id[key]
            
            # اگر ID=0 باشد، پردازش بدون محدودیت طول
            if id_val == 0:
                text = key_to_text.get(key, original_value)
                processed_text = process_text(text, 0, 0, "\n")  # min=0, max=0 یعنی بدون محدودیت
                processed_value = convert_special_chars(processed_text)
                return f'<item key="{key}" value="{processed_value}" />'
            
            # اگر ID معتبر باشد (غیر صفر و در تنظیمات وجود داشته باشد)
            elif id_val in valid_ids:
                text = key_to_text.get(key, original_value)
                settings = id_settings[id_val]
                processed_text = process_text(text, settings['min_len'], settings['max_len'], "\n")
                processed_value = convert_special_chars(processed_text)
                return f'<item key="{key}" value="{processed_value}" />'
        
        # اگر نیازی به پردازش نبود، مقدار اصلی را برگردان
        return match.group(0)
    
    # اعمال جایگزینی‌ها در محتوای XML
    processed_xml = item_pattern.sub(replace_item, xml_content)
    
    # ذخیره فایل خروجی
    with open(output_xml_path, 'w', encoding='utf-8') as f:
        f.write(processed_xml)
    
    print(f"پردازش با موفقیت انجام شد. فایل خروجی: {output_xml_path}")

def process_text(text, min_len, max_len, linebreaker):
    if pd.isna(text):
        return ""
    
    text = str(text) if isinstance(text, (int, float)) else text
    
    # اگر min=0 و max=0 باشد، تقسیم خط انجام نمی‌شود
    if min_len > 0 and max_len > 0 and max_len > min_len:
        processed_text = add_linebreaks(text, min_len, max_len, linebreaker)
    else:
        processed_text = text
    
    # اصلاح حروف عربی/فارسی
    processed_text = reshape_arabic(processed_text)
    
    # معکوس کردن جملات
    processed_text = rearrange_sentences(processed_text, linebreaker)
    
    return processed_text

def add_linebreaks(text, min_len, max_len, linebreaker):
    arabic_regex = re.compile(r'[\u0600-\u08FF\uFB50-\uFEFF]')
    output_lines = []
    
    for line in str(text).splitlines():
        if arabic_regex.search(line):
            parts = []
            while len(line) > max_len:
                split_pos = line.rfind(' ', min_len, max_len)
                if split_pos == -1:
                    split_pos = max_len
                parts.append(line[:split_pos])
                line = line[split_pos:].lstrip()
            parts.append(line)
            line = linebreaker.join(parts)
        output_lines.append(line)
    
    return '\n'.join(output_lines)

def reshape_arabic(text):
    arabic_regex = re.compile(r'[\u0600-\u08FF\uFB50-\uFEFF]')
    output_lines = []
    
    for line in str(text).splitlines():
        if arabic_regex.search(line):
            reshaped = arabic_reshaper.reshape(line)
            line = get_display(reshaped)
        output_lines.append(line)
    
    return '\n'.join(output_lines)

def rearrange_sentences(text, linebreaker):
    arabic_regex = re.compile(r'[\u0600-\u08FF\uFB50-\uFEFF]')
    output_lines = []
    
    for line in str(text).splitlines():
        if arabic_regex.search(line):
            sentences = [s.strip() for s in line.split(linebreaker) if s.strip()]
            sentences.reverse()
            line = linebreaker.join(sentences)
        output_lines.append(line)
    
    return '\n'.join(output_lines)

def convert_special_chars(text):
    """
    تبدیل کاراکترهای خاص به entityهای XML با ترتیب صحیح
    """
    if not isinstance(text, str):
        text = str(text)
    
    return (
        text
        .replace("&", "&amp;")  # اول از همه
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
        .replace("\n", "&#10;")  # در انتها
    )

if __name__ == "__main__":
    excel_path = "Book1.xlsx"  # مسیر فایل اکسل
    xml_path = "english_Original.xml"   # مسیر فایل XML ورودی
    output_xml_path = "english.xml"  # مسیر فایل XML خروجی
    
    process_excel_to_xml(excel_path, xml_path, output_xml_path)