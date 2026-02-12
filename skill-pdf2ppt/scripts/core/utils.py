import hashlib
import os
import math
import re
import logging
from html.parser import HTMLParser

def get_pdf_hash(pdf_path):
    """计算 PDF 文件的 MD5 哈希（用于缓存标识）"""
    hasher = hashlib.md5()
    with open(pdf_path, 'rb') as f:
        hasher.update(f.read())
    return hasher.hexdigest()[:8]

def ensure_cache_dir(cache_path):
    """确保缓存目录存在"""
    if not os.path.exists(cache_path):
        os.makedirs(cache_path)
        logging.info(f"创建缓存目录: {cache_path}")

def recursive_blocks(blocks):
    """递归提取所有 Layout Block (借鉴 MinerU 优化思路)"""
    result = []
    for block in blocks:
        if isinstance(block, dict):
            # 如果 block 自身包含 blocks (嵌套布局)
            if "blocks" in block:
                result.extend(recursive_blocks(block["blocks"]))
            else:
                result.append(block)
    return result

def calculate_font_size_gemini_style(block_fontSize, page_height, slide_height_inches=7.5):
    """
    借鉴 Gemini Canvas 版本的字体计算公式
    参数:
        block_fontSize: PDF中元素的字号（像素）
        page_height: PDF页面高度（像素）
        slide_height_inches: PPT幻灯片高度（英寸）
    """
    # 72 pt per inch
    scale_factor = (slide_height_inches * 72) / page_height
    # 原来用0.8太保守，改为0.95
    font_size = block_fontSize * scale_factor * 0.95
    
    # 限制范围
    font_size = max(6, min(72, font_size))
    
    return int(font_size)

def estimate_font_size_by_area(bbox_w, bbox_h, char_count, is_title=False):
    """根据面积和字符数估算字号"""
    if char_count <= 0:
        return 14
        
    # 如果是标题(字数少)，可以直接用高度估算
    if char_count < 15 and bbox_w / bbox_h > 2:
        return bbox_h * 0.7
    
    # 面积法公式：FontSize = sqrt(Area / (K * CharCount))
    # K 是单个字符占用的平均面积系数，包含行距等
    # 0.8 是经验值
    area = bbox_w * bbox_h
    estimated_fontSize_px = math.sqrt(area / (0.8 * char_count))
    
    # 限制字号不能超过 bbox 高度的一定比例 (防止单行时溢出)
    estimated_fontSize_px = min(estimated_fontSize_px, bbox_h * 0.9)
    
    return estimated_fontSize_px

def clean_latex_symbols(text):
    """
    清理文本中的LaTeX格式符号（如 '$'、'\circ'等）
    """
    if not text:
        return text
    
    # 处理转义的反斜杠：将 '\\%' 转换为 '%'
    text = text.replace('\\%', '%')
    text = text.replace('\\$', '$')  # 先处理转义的$
    
    # 处理常见的LaTeX命令
    latex_replacements = {
        r'\\circ': '°',        # 度数符号
        r'\\degree': '°',      # 度数符号
        r'\\times': '×',       # 乘号
        r'\\div': '÷',         # 除号
        r'\\pm': '±',          # 正负号
        r'\\leq': '≤',         # 小于等于
        r'\\geq': '≥',         # 大于等于
        r'\\neq': '≠',         # 不等于
        r'\\sim': '~',         # 约等于
        r'\\approx': '≈',      # 近似等于
    }
    
    for latex_cmd, unicode_char in latex_replacements.items():
        text = text.replace(latex_cmd, unicode_char)
    
    # 处理LaTeX上标和下标: $360^{\circ}$ -> 360°
    # 匹配模式: $数字^{命令}$ 或 $数字^命令$
    text = re.sub(r'\$(\d+)\^\{\\circ\}\$', r'\1°', text)  # $360^{\circ}$ -> 360°
    text = re.sub(r'\$(\d+)\^\\circ\$', r'\1°', text)      # $360^\circ$ -> 360°
    
    # 更通用的处理: 移除所有 ${...} 格式，保留内部内容
    text = re.sub(r'\$\{([^}]+)\}\$', r'\1', text)
    
    # 移除上标标记 ^{...}
    text = re.sub(r'\^\{([^}]+)\}', r'\1', text)
    
    # 移除下标标记 _{...}
    text = re.sub(r'_\{([^}]+)\}', r'\1', text)
    
    # 移除数学公式的 $ 符号：匹配 $...$ 格式，提取中间内容
    text = re.sub(r'\$([^$]+)\$', r'\1', text)
    
    # 移除剩余的单个 $ 符号
    text = text.replace('$', '')
    
    # 移除剩余的花括号（如果有的话）
    text = text.replace('{', '').replace('}', '')
    
    return text

def parse_html_table(html_content):
    """
    解析HTML表格内容，返回二维数组
    """
    if not html_content:
        return None
    
    try:
        class TableParser(HTMLParser):
            def __init__(self):
                super().__init__()
                self.tables = []
                self.current_table = []
                self.current_row = []
                self.current_cell = []
                self.in_table = False
                self.in_row = False
                self.in_cell = False
            
            def handle_starttag(self, tag, attrs):
                if tag == 'table':
                    self.in_table = True
                    self.current_table = []
                elif tag == 'tr' and self.in_table:
                    self.in_row = True
                    self.current_row = []
                elif tag in ['td', 'th'] and self.in_row:
                    self.in_cell = True
                    self.current_cell = []
            
            def handle_endtag(self, tag):
                if tag == 'table':
                    self.in_table = False
                    if self.current_table:
                        self.tables.append(self.current_table)
                elif tag == 'tr' and self.in_row:
                    self.in_row = False
                    if self.current_row:
                        self.current_table.append(self.current_row)
                elif tag in ['td', 'th'] and self.in_cell:
                    self.in_cell = False
                    cell_text = ''.join(self.current_cell).strip()
                    self.current_row.append(cell_text)
            
            def handle_data(self, data):
                if self.in_cell:
                    self.current_cell.append(data)
        
        parser = TableParser()
        parser.feed(html_content)
        
        if parser.tables and len(parser.tables) > 0:
            return parser.tables[0]  # 返回第一个表格
        return None
    except Exception as e:
        logging.warning(f"HTML表格解析失败: {e}")
        return None

def is_watermark_element(element, all_elements, page_width, page_height):
    """
    检测元素是否为水印（基于重复出现的文本内容、位置和已知水印关键词）
    """
    if not isinstance(element, dict):
        return False
    
    elem_type = element.get('type')
    # 只检测文本类型的元素（text、title、footer）
    if elem_type not in ['text', 'title', 'footer']:
        return False
    
    text_content = element.get('text') or element.get('content', '')
    if not text_content or len(text_content.strip()) == 0:
        return False
    
    text_stripped = text_content.strip()
    
    # 已知水印关键词列表（可扩展）
    watermark_keywords = ['NotebookLM', 'notebook lm', 'notebooklm']
    if any(keyword.lower() in text_stripped.lower() for keyword in watermark_keywords):
        logging.debug(f"  检测到已知水印关键词: '{text_stripped}'")
        return True
    
    # 检测位置：右下角区域（最后20%宽度和高度）
    bbox = element.get('bbox')
    if bbox and len(bbox) == 4:
        x1, y1, x2, y2 = bbox
        # 计算元素中心点
        center_x = (x1 + x2) / 2
        center_y = (y1 + y2) / 2
        
        # 右下角判定：X > 80%宽度 且 Y > 80%高度
        is_bottom_right = (center_x > page_width * 0.8) and (center_y > page_height * 0.8)
        
        if is_bottom_right and len(text_stripped) <= 30:
            logging.debug(f"  检测到右下角短文本水印: '{text_stripped}'")
            return True
    
    # 统计相同文本内容的出现次数（原有逻辑）
    count = 0
    for other_elem in all_elements:
        if not isinstance(other_elem, dict):
            continue
        other_type = other_elem.get('type')
        if other_type not in ['text', 'title', 'footer']:
            continue
        other_text = other_elem.get('text') or other_elem.get('content', '')
        if other_text and other_text.strip() == text_stripped:
            count += 1
    
    # 如果相同内容在页面中出现5次或以上，且文本较短，标记为水印
    # 提高阈值避免误判正常的重复内容
    if count >= 5 and len(text_stripped) <= 20:
        logging.debug(f"  检测到重复短文本水印(出现{count}次): '{text_stripped}'")
        return True
    
    return False
