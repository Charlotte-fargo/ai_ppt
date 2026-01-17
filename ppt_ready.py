import json
import os
import re
import logging
from datetime import datetime
from PIL import Image

from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# 引入配置文件
import config

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class PPTGenerator:
    def __init__(self, template_path, json_path, images_dir, location_name):
        self.template_path = template_path
        self.json_path = json_path
        self.images_dir = images_dir
        self.location = location_name
        self.prs = None
        self.data = None

    def load_resources(self):
        """加载模板和JSON数据"""
        try:
            if not os.path.exists(self.template_path):
                raise FileNotFoundError(f"模板文件未找到: {self.template_path}")
            self.prs = Presentation(self.template_path)
            # logging.info(f"模板加载成功，共有 {len(self.prs.slide_layouts)} 个布局")

            if not os.path.exists(self.json_path):
                raise FileNotFoundError(f"数据文件未找到: {self.json_path}")
            
            with open(self.json_path, "r", encoding="utf-8") as f:
                self.data = json.load(f)
            logging.info("资源加载成功")
        except Exception as e:
            logging.error(f"资源加载失败: {e}")
            raise

    # --- 通用工具方法 ---

    def _set_text_style(self, run, font_name='华文细黑', size=12, bold=False, color=config.COLOR_BLACK):
        """统一设置文本样式"""
        run.font.name = font_name
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color

    def _get_image_dimensions(self, image_path):
        """获取图片尺寸"""
        try:
            with Image.open(image_path) as img:
                return img.size
        except Exception as e:
            logging.warning(f"无法读取图片尺寸 {image_path}: {e}")
            return 0, 0

    def _calculate_fitted_size(self, img_width, img_height, max_width, max_height):
        """计算适合给定区域的图片大小，保持纵横比"""
        if img_height == 0 or max_height == 0: return 0, 0
        
        aspect = img_width / img_height
        max_aspect = max_width / max_height

        if aspect > max_aspect:
            fitted_width = max_width
            fitted_height = max_width / aspect
        else:
            fitted_height = max_height
            fitted_width = max_height * aspect
        return fitted_width, fitted_height

    def _remove_all_picture_placeholders(self, slide):
        """移除幻灯片中所有的图片占位符"""
        for shape in list(slide.shapes):
            if shape.is_placeholder and shape.placeholder_format.type == 18:
                try:
                    sp = shape._element
                    sp.getparent().remove(sp)
                except Exception:
                    pass

    def _find_matching_image(self, title, match_len=2):
        """根据内容标题开头查找匹配的图片"""
        if not os.path.exists(self.images_dir):
            logging.warning(f"图片目录不存在: {self.images_dir}")
            return None
        
        # 清理标题
        cleaned = re.sub(r'[【】\[\]()（）,，.。、/\\|]', '', title).replace(' ', '').strip()
        parts = cleaned.split('_')
        key = parts[0][:match_len] if len(parts[0]) >= match_len else parts[0]
        # 特殊规则
        if "债市" in parts[0]: key = "债券"
        print(f"    查找匹配图片，标题开头: '{key}'")
        logging.debug(f"查找图片: 原始标题='{title}', 关键字='{key}'")

        image_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp']
        for f in os.listdir(self.images_dir):
            if any(f.lower().endswith(ext) for ext in image_extensions):
                if f.startswith(key):
                    return os.path.join(self.images_dir, f)
        return None

    def _fill_image(self, slide, image_path):
        """在幻灯片中插入图片并调整填充方式"""
        if not os.path.exists(image_path): return None

        img_width, img_height = self._get_image_dimensions(image_path)
        if img_width == 0: return None
        
        # 查找图片占位符
        placeholder = None
        for shape in slide.shapes:
            if shape.placeholder_format.type == 18:
                placeholder = shape
                break

        if placeholder:
            w, h = self._calculate_fitted_size(img_width, img_height, placeholder.width, placeholder.height)
            left = placeholder.left + (placeholder.width - w) // 2
            top = placeholder.top + (placeholder.height - h) // 2
            
            try:
                picture = slide.shapes.add_picture(image_path, left, top, w, h)
                logging.info(f"✓ 插入图片: {os.path.basename(image_path)}")
                return picture
            except Exception as e:
                logging.error(f"插入图片失败: {e}")
        return None

    def _add_image_annotations(self, slide, image_path):
        """添加图表标题和资料来源 (使用 config 中的坐标配置)"""
        try:
            file_name = os.path.basename(image_path)
            parts = os.path.splitext(file_name)[0].split('_')

            # A. 添加图表标题
            if len(parts) >= 2:
                chart_title = parts[1]
                if chart_title and chart_title != "NONE":
                    cfg = config.ANNOTATION_CONFIG['title']
                    
                    # 动态调整左边距
                    offset = Pt(0)
                    if len(chart_title) > 20: offset = Pt(80)
                    elif len(chart_title) > 10: offset = Pt(30)
                    
                    t_left = cfg['left_base'] - offset
                    
                    textbox = slide.shapes.add_textbox(t_left, cfg['top'], cfg['width'], cfg['height'])
                    run = textbox.text_frame.paragraphs[0].add_run()
                    run.text = chart_title
                    self._set_text_style(run, font_name=cfg['font_name'], size=cfg['size'], bold=True)

            # B. 添加资料来源
            source_text = ""
            offset_s = Pt(0)
            
            if len(parts) >= 2:
                # 假设最后一部分是来源
                raw_source = parts[-1].replace('，', ' ')
                source_text = re.sub(r'\s+', ' ', raw_source).strip()
                if len(source_text) > 13: offset_s = Pt(40)
                else: offset_s = Pt(20)
            else:
                source_text = parts[0]
                offset_s = -Pt(15)
            
            source_text = source_text.strip('_ ')
            
            if source_text:
                cfg_s = config.ANNOTATION_CONFIG['source']
                s_left = cfg_s['left_base'] - offset_s
                
                textbox = slide.shapes.add_textbox(s_left, cfg_s['top'], cfg_s['width'], cfg_s['height'])
                # 确保每次都是新段落
                p = textbox.text_frame.paragraphs[0]
                if not p.runs: p.add_run()
                p.runs[0].text = f"资料来源：{source_text}"
                
                self._set_text_style(p.runs[0], font_name=cfg_s['font_name'], size=cfg_s['size'], color=config.COLOR_GRAY)

        except Exception as e:
            logging.warning(f"添加图片注释失败: {e}")

    # --- 页面生成方法 ---

    def create_cover(self):
        """生成封面"""
        slide = self.prs.slides[0]
        doc = self.data.get("document", {})
        
        try:
            # 假设占位符顺序：0=日期, 1=作者, 2=标题
            slide.shapes[0].text = datetime.now().strftime("%Y-%m-%d")
            
           # 1. 设置作者 (Shape 1)
            # -------------------------------------------------------
            # 第一步：先获取文本框
            tf_author = slide.shapes[1].text_frame
            
            # 第二步：先填充文字内容 (这步会重置样式，所以必须先做)
            tf_author.paragraphs[0].text = doc.get("author", "")

            # 第三步：获取 Run 并设置样式 (加粗 + 微软雅黑)
            if tf_author.paragraphs[0].runs:
                self._set_text_style(
                    tf_author.paragraphs[0].runs[0], 
                    size=35, 
                    bold=True,                      # 加粗
                    color=config.COLOR_LIGHT_BLUE, 
                    font_name='Microsoft YaHei'      # <--- 这里指定字体
                )

            # -------------------------------------------------------
            # 2. 设置标题 (Shape 2)
            # -------------------------------------------------------
            # 第一步：先获取文本框
            tf_title = slide.shapes[2].text_frame
            
            # 第二步：先填充文字内容
            tf_title.paragraphs[0].text = doc.get("title", "")

            # 第三步：获取 Run 并设置样式 (加粗 + 微软雅黑)
            if tf_title.paragraphs[0].runs:
                self._set_text_style(
                    tf_title.paragraphs[0].runs[0], 
                    size=44, 
                    bold=True,                       
                    color=config.COLOR_DARK_BLUE, 
                    font_name='Microsoft YaHei'      # <--- 这里指定字体
                )
        except IndexError:
            logging.warning("封面页占位符索引不匹配")

    def create_summary(self):
        """生成摘要页"""
        slide = self.prs.slides[1]
        summary = self.data.get("executive_summary", {})
        
        # 标题
        if slide.shapes[0].has_text_frame:
            self._set_text_style(slide.shapes[0].text_frame.paragraphs[0].runs[0], 
                                 size=29.1, bold=True, color=config.COLOR_DARK_BLUE)

        # 表格
        try:
            table = slide.shapes[1].table
            cols = summary.get("columns", [])
            rows = summary.get("rows", [])

            for r_idx, row_data in enumerate(rows):
                if r_idx + 1 >= len(table.rows): break
                
                for c_idx, col_key in enumerate(cols):
                    # 计算表格实际列索引 (从索引1开始，即第二列)
                    table_col_idx = c_idx + 1
                    
                    if table_col_idx >= len(table.columns): break
                    
                    cell = table.cell(r_idx + 1, table_col_idx)
                    text = row_data.get(col_key, "")

                    # 1. 获取 TextFrame 并清空 (关键步骤)
                    tf = cell.text_frame
                    tf.clear()
                    p = tf.paragraphs[0]

                    # 2. 添加 Run 并赋值
                    run = p.add_run()
                    run.text = text
                    
                    # 3. 计算样式逻辑
                    # 逻辑A: 投资逻辑字数多则变小
                    f_size = 9 if len(text) > 200 and col_key == "投资逻辑" else 10
                    
                    # 逻辑B: 判断是否为第二列 (index=1)，如果是则加粗
                    is_second_column = (table_col_idx == 1)

                    # 4. 应用样式
                    self._set_text_style(
                        run, 
                        size=f_size,
                        bold=is_second_column,       # <--- 这里控制加粗
                        font_name='Microsoft YaHei'  # 统一字体
                    )
                    
                    # 5. 设置居中 (仅第二列)
                    if is_second_column:
                        p.alignment = PP_ALIGN.CENTER
        
           

        except Exception as e:
            logging.error(f"摘要表格填充失败: {e}")

    def create_content_pages(self):
        """生成内容页"""
        # 清理旧幻灯片
        while len(self.prs.slides) > 2:
            rId = self.prs.slides._sldIdLst[2].rId
            self.prs.part.drop_rel(rId)
            del self.prs.slides._sldIdLst[2]

        layout = self.prs.slide_layouts[config.LAYOUT_IDX['content']]

        for i, content in  enumerate(self.data["content_slides"]):
            # 前6页尝试复用（如果有的话），否则新建
            if i < 6 and (2 + i) < len(self.prs.slides):
                slide = self.prs.slides[2 + i]
            else:
                slide = self.prs.slides.add_slide(layout)
            # 占位符
            if len(slide.placeholders) >= 2:
                title_ph = slide.placeholders[0]
                body_ph = slide.placeholders[10]
            else:
                # 如果占位符数量不足，尝试其他方法
                title_ph = None
                body_ph = None

            if title_ph and title_ph.has_text_frame:
                title_ph.text = content["title"]
                text_frame = title_ph.text_frame
                if text_frame.paragraphs and len(text_frame.paragraphs) > 0:
                    paragraph = text_frame.paragraphs[0]
                    if paragraph.runs and len(paragraph.runs) > 0:
                        run = paragraph.runs[0]
                        run.font.name = '华文细黑'
                        run.font.size = Pt(29.1)
                        run.font.bold = True
                        run.font.color.rgb = config.COLOR_DARK_BLUE

            # 正文
            if body_ph and body_ph.has_text_frame:
                # 添加项目符号
                bullets = content["bullets"]
                body_text = ""
                for bullet in bullets:
                    body_text += f"{bullet}\n"
                body_ph.text = body_text
                # 根据字数调整字体大小
                total_chars = len(body_text)
                font_size = Pt(14) if len(bullets) > 3 or total_chars > 200 else Pt(16)
                
                for p in body_ph.text_frame.paragraphs:
                    for run in p.runs:
                        run.font.size = font_size
                print(f"Slide {i+1} Body Text added")

            # 图片
            matched_image_path = self._find_matching_image(content["title"], match_len=2)
            if matched_image_path:
                self._fill_image(slide, matched_image_path)
                # 内容页总是添加注释 
                self._add_image_annotations(slide, matched_image_path)

            self._remove_all_picture_placeholders(slide)
            print("-" * 50)


    def create_image_slide(self, topic):
        """生成纯图页"""
        layout = self.prs.slide_layouts[config.LAYOUT_IDX['image_only']]
        slide = self.prs.slides.add_slide(layout)
        
        slide.placeholders[0].text = topic
        
        img_path = self._find_matching_image(topic, match_len=2)
        if img_path:
            self._fill_image(slide, img_path)
            
            # 仅资金流加注释
            if topic == "资金流":
                self._add_image_annotations(slide, img_path)
        
        self._remove_all_picture_placeholders(slide)

    def create_contact_page(self):
        """生成联系页"""
        layout = self.prs.slide_layouts[config.LAYOUT_IDX['contact']]
        slide = self.prs.slides.add_slide(layout)
        slide.placeholders[0].text = "联系我们"

        # 主文本框 (索引15)
        try:
            tf = slide.placeholders[15].text_frame
            tf.clear()
            p = tf.paragraphs[0]
            r = p.add_run()
            r.text = "www.fargowealth.com"
            r.font.underline = True
            
            p2 = tf.add_paragraph()
            p2.text = "资产管理服务由集团旗下公司绅士资本提供"
        except KeyError:
            pass

        # 地址列表
        for idx, text in config.CONTACT_ADDRESSES.items():
            try:
                tf = slide.placeholders[idx].text_frame
                tf.clear()
                lines = text.split("\n")
                for i, line in enumerate(lines):
                    p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                    p.text = line
                    for run in p.runs:
                        # 城市名加粗逻辑
                        if i == 0: run.font.bold = True
                        run.font.size = Pt(8)
            except KeyError:
                continue

    def create_disclaimer_pages(self):
        """生成免责页"""
        texts = config.DISCLAIMER_TEXTS.get(self.location, config.DISCLAIMER_TEXTS["default"])

        # 中文页
        if texts.get("cn"):
            layout = self.prs.slide_layouts[config.LAYOUT_IDX['disclaimer_cn']]
            slide = self.prs.slides.add_slide(layout)
            slide.placeholders[0].text = "免责声明"
            self._set_text_style(slide.placeholders[0].text_frame.paragraphs[0].runs[0], 
                                 size=29.1, bold=True, color=config.COLOR_DARK_BLUE)
            
            body = slide.placeholders[12]
            body.text = "\n".join(texts["cn"])
            for p in body.text_frame.paragraphs:
                for run in p.runs:
                    self._set_text_style(run, size=12)

        # 英文页 (非大陆)
        if self.location != "中国大陆" and texts.get("en"):
            layout = self.prs.slide_layouts[config.LAYOUT_IDX['disclaimer_en']]
            slide = self.prs.slides.add_slide(layout)
            slide.placeholders[0].text = "Disclaimer"
            self._set_text_style(slide.placeholders[0].text_frame.paragraphs[0].runs[0], 
                                 size=29.1, bold=True, color=config.COLOR_DARK_BLUE)
            
            body = slide.placeholders[12]
            body.text = "\n".join(texts["en"])
            for p in body.text_frame.paragraphs:
                for run in p.runs:
                    self._set_text_style(run, font_name='Arial', size=9)

    def run(self, output_filename):
        """执行全流程"""
        logging.info(f"开始生成 PPT - 地点: {self.location}")
        self.load_resources()
        
        self.create_cover()
        self.create_summary()
        self.create_content_pages()
        self.create_image_slide("个债精选")
        self.create_image_slide("个股精选")
        self.create_image_slide("资金流")
        self.create_contact_page()
        self.create_disclaimer_pages()
        
        # 保存
        os.makedirs(config.OUTPUT_DIR, exist_ok=True)
        out_path = os.path.join(config.OUTPUT_DIR, output_filename)
        
        if os.path.exists(out_path):
            try:
                os.remove(out_path)
            except PermissionError:
                logging.error("文件被占用，无法删除旧文件")
                return

        self.prs.save(out_path)
        logging.info(f"PPT 生成完成: {out_path} (共 {len(self.prs.slides)} 页)")


# ================= 对外接口函数 =================

def generate_ppt_from_json(json_path, template_path, output_filename, location_name, images_dir):
    """
    供 main.py 调用的简易接口函数
    """
    try:
        generator = PPTGenerator(template_path, json_path, images_dir, location_name)
        generator.run(output_filename)
        return True
    except Exception as e:
        logging.error(f"PPT 生成过程中发生错误: {e}")
        return False
# 示例调用
if __name__ == "__main__":
    # 配置
    IMAGES_DIR = "input_articles/20260116/images_20260116"
    DATA_JSON = "final_investment_report.json"
    
    # 交互输入
    print("请选择需要PPT的地点：\n1. 中国大陆\n2. 香港\n3. 新加坡")
    choice = input("请输入数字 (1/2/3): ").strip()
    
    LOCATION_MAP = {
        "1": ("中国大陆", "template/AI PPT v3.pptx"),
        "2": ("香港", "template/AI PPT v2.pptx"),
        "3": ("新加坡", "template/AI PPT v2.pptx")
    }
    
    if choice not in LOCATION_MAP:
        print("输入无效，程序退出")
        exit(1)
        
    loc_name, template_file = LOCATION_MAP[choice]
    output_file = f"AI_PPT_generated_{loc_name}.pptx"
    
    generate_ppt_from_json(DATA_JSON, template_file, output_file, loc_name, IMAGES_DIR)