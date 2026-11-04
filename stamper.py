import fitz  # PyMuPDF
from PIL import Image
import os
import io
import random
import re

class PDFStamper:
    def __init__(self, pdf_path, stamp_image_path):
        if not os.path.exists(pdf_path):
            raise FileNotFoundError(f"PDF file not found: {pdf_path}")
        if not os.path.exists(stamp_image_path):
            raise FileNotFoundError(f"Stamp image file not found: {stamp_image_path}")

        self.pdf_path = pdf_path
        self.stamp_image_path = stamp_image_path
        self.document = fitz.open(pdf_path)
        self.stamp_image_pil = Image.open(stamp_image_path).convert("RGBA")

    def _place_stamp(self, page, stamp_image_pil, x, y, stamp_width, stamp_height):
        """在指定位置放置印章（x,y为左上角坐标，单位：点）"""
        target_w = int(stamp_width)
        target_h = int(stamp_height)
        
        if target_w <= 0 or target_h <= 0:
            return

        if stamp_image_pil.width != target_w or stamp_image_pil.height != target_h:
            stamp_image_pil = stamp_image_pil.resize((target_w, target_h), Image.Resampling.LANCZOS)
        
        img_byte_arr = io.BytesIO()
        stamp_image_pil.save(img_byte_arr, format='PNG')
        img_byte_arr = img_byte_arr.getvalue()

        rect = fitz.Rect(x, y, x + target_w, y + target_h)
        page.insert_image(rect, stream=img_byte_arr)

    def _find_keyword_position(self, page, keyword, occurrence=1):
        """
        在页面中查找关键词的位置，返回关键词区域的中心点坐标（单位：点）
        keyword: 要搜索的关键词
        occurrence: 匹配第几次出现（默认第1次）
        返回: (center_x, center_y) 或 None
        """
        # 搜索关键词（支持正则表达式）
        text_instances = page.search_for(keyword)
        
        if not text_instances:
            # 如果直接搜索不到，尝试用正则表达式搜索
            try:
                text_instances = page.search_for(keyword, quads=True)
                if not text_instances:
                    return None
            except:
                return None
        
        if len(text_instances) < occurrence:
            return None
        
        # 获取指定次数的匹配区域
        rect = text_instances[occurrence - 1]
        
        # 计算中心点
        center_x = (rect.x0 + rect.x1) / 2
        center_y = (rect.y0 + rect.y1) / 2
        
        return (center_x, center_y)

    def stamp_first_page_by_keyword(self, keyword, scale, offset_up_cm=0.1, occurrence=1):
        """
        根据关键词位置盖章（首页）
        keyword: 要搜索的关键词（支持正则表达式）
        scale: 印章缩放比例
        offset_up_cm: 从关键词中心向上偏移的距离（单位：厘米），默认0.1cm
        occurrence: 匹配第几次出现（默认第1次）
        """
        if len(self.document) < 1:
            return
        
        page = self.document[0]
        
        # 查找关键词位置
        pos = self._find_keyword_position(page, keyword, occurrence)
        if pos is None:
            print(f"警告: 在首页未找到关键词 '{keyword}'，跳过盖章")
            return
        
        center_x, center_y = pos
        
        # 计算印章尺寸（单位：点）
        stamp_width = int(self.stamp_image_pil.width * scale)
        stamp_height = int(self.stamp_image_pil.height * scale)
        
        # 将偏移量从厘米转换为点（1点 = 1/72英寸，1英寸 = 2.54厘米）
        # 1点 ≈ 0.03528厘米，所以 1厘米 ≈ 28.346点
        points_per_cm = 28.346
        offset_up_points = offset_up_cm * points_per_cm
        
        # 计算印章左上角坐标（以关键词中心为基准，向上偏移，并居中）
        # 向上偏移：y减小
        stamp_x = center_x - stamp_width / 2
        stamp_y = center_y - stamp_height / 2 - offset_up_points
        
        # 确保印章不超出页面边界
        page_rect = page.rect
        if stamp_x < 0:
            stamp_x = 0
        if stamp_y < 0:
            stamp_y = 0
        if stamp_x + stamp_width > page_rect.width:
            stamp_x = page_rect.width - stamp_width
        if stamp_y + stamp_height > page_rect.height:
            stamp_y = page_rect.height - stamp_height
        
        self._place_stamp(page, self.stamp_image_pil, stamp_x, stamp_y, stamp_width, stamp_height)

    def stamp_first_page(self, x, y, scale):
        """原始方法：在指定坐标盖章（保留兼容）"""
        if len(self.document) < 1:
            return
        page = self.document[0]
        stamp_width = int(self.stamp_image_pil.width * scale)
        stamp_height = int(self.stamp_image_pil.height * scale)
        self._place_stamp(page, self.stamp_image_pil, x, y, stamp_width, stamp_height)

    def stamp_multi_pages(self, x, y, scale, page_limit=None, random_offset=False, jitter_amplitude=20):
        """多页盖章（保留原有逻辑）"""
        total_pages = len(self.document)
        limit = total_pages
        
        if page_limit is not None:
            try:
                limit_int = int(page_limit)
                if limit_int > 0:
                    limit = min(limit_int, total_pages)
            except ValueError:
                pass

        stamp_width = int(self.stamp_image_pil.width * scale)
        stamp_height = int(self.stamp_image_pil.height * scale)

        for i in range(limit):
            page = self.document[i]
            current_x, current_y = x, y
            
            if i > 0 and random_offset:
                jitter_x = random.uniform(-jitter_amplitude, jitter_amplitude)
                jitter_y = random.uniform(-jitter_amplitude, jitter_amplitude)
                current_x += jitter_x
                current_y += jitter_y

            self._place_stamp(page, self.stamp_image_pil, current_x, current_y, stamp_width, stamp_height)

    def stamp_cross_page(self, scale, first_page_extra_ratio=0.15):
        """骑缝章（保留原有逻辑）"""
        num_pages = len(self.document)
        if num_pages == 0:
            return

        scaled_total_width = int(self.stamp_image_pil.width * scale)
        scaled_total_height = int(self.stamp_image_pil.height * scale)
        
        if scaled_total_width < num_pages:
            scaled_total_width = max(scaled_total_width, num_pages)
            
        scaled_stamp_pil = self.stamp_image_pil.resize((scaled_total_width, scaled_total_height), Image.Resampling.LANCZOS)
        
        part_width_per_page = scaled_total_width / num_pages
        
        for i in range(num_pages):
            page = self.document[i]
            page_width = page.rect.width
            page_height = page.rect.height

            left = int(i * part_width_per_page)
            upper = 0
            right = int((i + 1) * part_width_per_page)
            lower = scaled_total_height
            
            if i == 0:
                extra_width = int(scaled_total_width * first_page_extra_ratio)
                right += extra_width
                if right > scaled_total_width:
                    right = scaled_total_width
            
            if i == num_pages - 1:
                right = scaled_total_width

            if right <= left:
                continue

            cropped_stamp_pil = scaled_stamp_pil.crop((left, upper, right, lower))
            current_part_width = right - left
            
            EDGE_MARGIN = 5 
            draw_x = page_width - current_part_width - EDGE_MARGIN
            draw_y = (page_height - scaled_total_height) / 2
            
            self._place_stamp(page, cropped_stamp_pil, draw_x, draw_y, current_part_width, scaled_total_height)

    def save_pdf(self, output_path):
        """保存PDF"""
        self.document.save(output_path)
        self.document.close()
        print(f"Stamped PDF saved to: {output_path}")