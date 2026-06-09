import zipfile
import xml.etree.ElementTree as ET

with zipfile.ZipFile('../250903_Theo dõi cấp CME + Điểm danh_Bài giảng SHCM (Độ)_unlocked.xlsx', 'r') as z:
    xml_content = z.read('xl/worksheets/sheet1.xml')
    root = ET.fromstring(xml_content)
    
    ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
    
    for c in root.findall('.//ns:c', ns):
        text_elements = c.findall('.//ns:t', ns)
        for t in text_elements:
            if t.text and '1drv.ms' in t.text:
                print(f"Cell {c.get('r')}: {t.text}")
