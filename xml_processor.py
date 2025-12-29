import pandas as pd
import re
import arabic_reshaper
from bidi.algorithm import get_display

def process_excel_to_xml(excel_path, xml_path, output_xml_path):
    try:
        # خواندن فایل اکسل
        excel_data = pd.read_excel(excel_path, header=None)
        
        # ایجاد ساختارهای داده
        id_settings = {}
        valid_ids = set()
        key_to_id = {}
        key_to_text = {}

        # پردازش ردیف‌های اکسل
        for idx, row in excel_data.iterrows():
            # استخراج تنظیمات طول
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

            # استخراج نگاشت کلید به ID
            if len(row) > 3 and pd.notna(row[0]) and pd.notna(row[3]):
                key_to_id[row[0]] = row[3]

            # استخراج متن‌ها
            if len(row) > 1 and pd.notna(row[0]) and pd.notna(row[1]):
                key_to_text[row[0]] = row[1]

        # خواندن فایل XML
        with open(xml_path, 'r', encoding='utf-8') as f:
            xml_content = f.read()

        # پردازش XML
        item_pattern = re.compile(r'<item\s+key="([^"]+)"\s+value="([^"]*)"\s*/>')
        
        def replace_item(match):
            key = match.group(1)
            original_value = match.group(2)
            
            if key in key_to_id:
                id_val = key_to_id[key]
                
                if id_val == 0:  # پردازش بدون محدودیت طول
                    text = key_to_text.get(key, original_value)
                    processed_text = process_text(text, 0, 0, "\n")
                    return f'<item key="{key}" value="{convert_special_chars(processed_text)}" />'
                
                elif id_val in valid_ids:  # پردازش با محدودیت طول
                    text = key_to_text.get(key, original_value)
                    settings = id_settings[id_val]
                    processed_text = process_text(text, settings['min_len'], settings['max_len'], "\n")
                    return f'<item key="{key}" value="{convert_special_chars(processed_text)}" />'
            
            return match.group(0)

        processed_xml = item_pattern.sub(replace_item, xml_content)

        # ذخیره فایل خروجی
        with open(output_xml_path, 'w', encoding='utf-8') as f:
            f.write(processed_xml)
        
        return True, "پردازش با موفقیت انجام شد"
    
    except Exception as e:
        return False, f"خطا در پردازش: {str(e)}"

def process_text(text, min_len, max_len, linebreaker):
    if pd.isna(text):
        return ""
    
    text = str(text)
    
    if min_len > 0 and max_len > 0 and max_len > min_len:
        processed_text = add_linebreaks(text, min_len, max_len, linebreaker)
    else:
        processed_text = text
    
    processed_text = reshape_arabic(processed_text)
    processed_text = rearrange_sentences(processed_text, linebreaker)
    
    return processed_text

def add_linebreaks(text, min_len, max_len, linebreaker):
    arabic_regex = re.compile(r'[\u0600-\u08FF\uFB50-\uFEFF]')
    output_lines = []
    
    for line in text.splitlines():
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
    
    for line in text.splitlines():
        if arabic_regex.search(line):
            reshaped = arabic_reshaper.reshape(line)
            line = get_display(reshaped)
        output_lines.append(line)
    
    return '\n'.join(output_lines)

def rearrange_sentences(text, linebreaker):
    arabic_regex = re.compile(r'[\u0600-\u08FF\uFB50-\uFEFF]')
    output_lines = []
    
    for line in text.splitlines():
        if arabic_regex.search(line):
            sentences = [s.strip() for s in line.split(linebreaker) if s.strip()]
            sentences.reverse()
            line = linebreaker.join(sentences)
        output_lines.append(line)
    
    return '\n'.join(output_lines)

def convert_special_chars(text):
    if not isinstance(text, str):
        text = str(text)
    
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
        .replace("\n", "&#10;")
    )