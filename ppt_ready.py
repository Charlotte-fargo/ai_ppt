import json
import os
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import  PP_ALIGN
from datetime import datetime
import re
from PIL import Image

def get_image_dimensions(image_path):
    """获取图片尺寸"""
    img = Image.open(image_path)
    width, height = img.size
    img.close()
    return width, height

def calculate_fitted_size(img_width, img_height, max_width, max_height):
    """计算适合给定区域的图片大小，保持纵横比"""
    aspect = img_width / img_height if img_height > 0 else 1
    max_aspect = max_width / max_height if max_height > 0 else 1

    if aspect > max_aspect:
        # 图片更宽，适应宽度
        fitted_width = max_width
        fitted_height = max_width / aspect
    else:
        # 图片更高，适应高度
        fitted_height = max_height
        fitted_width = max_height * aspect

    return fitted_width, fitted_height

def find_picture_placeholder(slide, placeholder_idx=None):
    """查找图片占位符"""
    if placeholder_idx is not None:
        if placeholder_idx < len(slide.placeholders):
            ph = slide.placeholders[placeholder_idx]
            if ph.placeholder_format.type == 18:  # 图片占位符
                return ph
    else:
        for shape in slide.shapes:
            if shape.is_placeholder and shape.placeholder_format.type == 18:
                return shape
    return None

def remove_all_picture_placeholders(slide):
    """移除幻灯片中所有的图片占位符"""
    removed_count = 0
    # 需要小心，因为移除时列表会改变
    placeholders_to_remove = []
    for shape in slide.shapes:
        if shape.is_placeholder and shape.placeholder_format.type == 18:
            placeholders_to_remove.append(shape)

    for placeholder in placeholders_to_remove:
        try:
            sp = placeholder._element
            sp.getparent().remove(sp)
            print(f"✓ 删除图片占位符 (idx={placeholder.placeholder_format.idx})")
        except Exception as e:
            print(f"✗ 删除占位符失败: {e}")

    return removed_count

def clean_slide_after_insertion(slide):
    """
    清理幻灯片中未使用的占位符
    """
    print("清理幻灯片中的占位符...")

    # 移除空的图片占位符
    removed_count = remove_all_picture_placeholders(slide)

    print(f"共清理 {removed_count} 个占位符")
    return removed_count

def fill_image(slide, image_path, placeholder_idx=None, fill_type="fit"):
    """
    在幻灯片中插入图片并调整填充方式

    Args:
        slide: 幻灯片对象
        image_path: 图片路径
        placeholder_idx: 占位符索引（如果为None，则使用第一个图片占位符）
    """
    if not os.path.exists(image_path):
        print(f"图片文件不存在: {image_path}")
        return None

    
    # 获取图片尺寸
    img_width, img_height = get_image_dimensions(image_path)

    # 查找图片占位符
    placeholder = find_picture_placeholder(slide, placeholder_idx)

    # 如果找到了占位符
    if placeholder:
        # 计算适合占位符的大小
        fitted_width, fitted_height = calculate_fitted_size(img_width, img_height, placeholder.width, placeholder.height)

        # 居中放置
        left = placeholder.left + (placeholder.width - fitted_width) // 2
        top = placeholder.top + (placeholder.height - fitted_height) // 2

        picture = slide.shapes.add_picture(image_path, left, top, fitted_width, fitted_height)
        print(f"✓ 在占位符插入图片，大小: {fitted_width}x{fitted_height}")

        return picture


def find_matching_image(content_title, images_dir, match_length=3):
    """
    根据内容标题开头查找匹配的图片
    Args:
        content_title: 内容标题
        images_dir: 图片目录路径
        match_length: 匹配的字符数，默认3个字符
    Returns:
        匹配的图片路径或None
    """
    if not os.path.exists(images_dir):
        print(f"    警告: 图片目录不存在: {images_dir}")
        return None
    
    # 清理标题，移除特殊字符和空格
    cleaned_title = re.sub(r'[【】\[\]()（）,，.。、/\\|]', '', content_title)
    cleaned_title = cleaned_title.replace(' ', '').strip()
    parts = cleaned_title.split('_')
    print(f"    清理后的标题: '{parts[0]}'")
    
    # 取标题的前几个字符（根据match_length参数）
    if "债市" in parts[0]:
        title_start = "债券" 
    # elif "资金流" in cleaned_title:
    #     title_start = "个股投资观点"
    else:
        if len(parts[0]) >= match_length:
            title_start = parts[0][:match_length]
        else:
            title_start = parts[0]

    if not title_start:
        return None
    
    print(f"    查找匹配图片，标题开头: '{title_start}'")
    
    # 获取所有图片文件
    image_extensions = ['.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp']
    all_images = []
    
    for filename in os.listdir(images_dir):
        if any(filename.lower().endswith(ext) for ext in image_extensions):
            all_images.append(filename)
    
    if not all_images:
        print("    警告: 图片目录中没有图片文件")
        return None
    
    # 查找匹配的图片
    matched_images = []
    
    for filename in all_images:
        # 移除文件扩展名
        name_without_ext = os.path.splitext(filename)[0]
        print(f"    检查图片文件: {name_without_ext}{title_start}")
        # 检查是否以标题开头匹配
        if name_without_ext.startswith(title_start):
            matched_images.append(filename)
            print(f"    找到匹配图片: {filename}")    
            
            return os.path.join(images_dir, filename)
    return None
        
        # return None

def generate_cover_slide(prs, data, dark_blue):
    """生成封面幻灯片"""
    slide0 = prs.slides[0]
    doc = data["document"]

    # slide0.shapes[0].text = doc["date"]
    today_date = datetime.now().strftime("%Y-%m-%d")
    slide0.shapes[0].text = today_date
    slide0.shapes[1].text = doc["author"]
    slide0.shapes[2].text = doc["title"]

    # Format main title (shapes[2])
    text_frame = slide0.shapes[2].text_frame
    paragraph = text_frame.paragraphs[0]
    run = paragraph.runs[0]
    run.font.name = 'Microsoft YaHei'
    run.font.size = Pt(44)
    run.font.bold = True
    run.font.color.rgb = dark_blue

    # Format subtitle (shapes[1])
    text_frame = slide0.shapes[1].text_frame
    paragraph = text_frame.paragraphs[0]
    run = paragraph.runs[0]
    run.font.name = 'Microsoft YaHei'
    run.font.size = Pt(35)
    run.font.color.rgb = dark_blue


def generate_executive_summary_slide(prs, data, dark_blue):
    """生成执行摘要表格幻灯片"""
    slide1 = prs.slides[1]
    summary = data["executive_summary"]

    # 获取现有表格
    existing_table_shape = slide1.shapes[1]
    table = existing_table_shape.table

    # 检查现有表格的列数是否足够
    if len(table.columns) < 3:
        print("警告：现有表格列数不足3列，无法覆盖第二列和第三列")
    else:
        # 填充表头 - 只覆盖第二列和第三列
        for col_idx, col_name in enumerate(summary["columns"], start=1):  # 从第二列开始（索引1）
            if col_idx < len(table.columns):  # 确保列索引有效
                cell = table.cell(0, col_idx)
                # 清除原有内容
                cell.text = ""
                # 添加新内容
                text_frame = cell.text_frame
                paragraph = text_frame.paragraphs[0]
                # 设置段落水平居中
                run = paragraph.add_run()
                run.text = col_name
                # 格式化表头
                run.font.size = Pt(12)
                run.font.bold = True
                run.font.color.rgb = RGBColor(0, 0, 0)
                paragraph.alignment = PP_ALIGN.CENTER  # 水平居中

        # 填充数据行 - 只覆盖第二列和第三列
        for row_idx, row_data in enumerate(summary["rows"], start=1):  # 从第二行开始（索引1）
            # 确保行索引不超过表格行数
            if row_idx < len(table.rows):
                for col_idx, col_name in enumerate(summary["columns"]):
                    # 计算在表格中的实际列索引（从第二列开始）
                    table_col_idx = col_idx + 1
                    if table_col_idx < len(table.columns):  # 确保列索引有效
                        cell_text = row_data[col_name]
                        cell = table.cell(row_idx, table_col_idx)
                        # 清除原有内容
                        cell.text = ""
                        # 添加新内容
                        text_frame = cell.text_frame
                        paragraph = text_frame.paragraphs[0]
                        run = paragraph.add_run()
                        run.text = cell_text
                        # 格式化单元格
                        if col_name == "投资逻辑" and len(cell_text) > 200:
                            run.font.size = Pt(9)
                        else:
                            run.font.size = Pt(10)
                        run.font.color.rgb = RGBColor(0, 0, 0)
                        # 设置第二列（索引1）的文本居中
                        if table_col_idx == 1:
                            paragraph.alignment = PP_ALIGN.CENTER   # 设置第二列（索引1）的文本居中
            else:
                print(f"警告：数据行数超过表格行数，第{row_idx}行数据未填充")
                break

    # Format title for slide 1 (assuming shapes[0] is title)
    if len(slide1.shapes) > 0 and slide1.shapes[0].has_text_frame:
        text_frame = slide1.shapes[0].text_frame
        if text_frame.paragraphs:
            paragraph = text_frame.paragraphs[0]
            if paragraph.runs:
                run = paragraph.runs[0]
                run.font.name = '华文细黑'
                run.font.size = Pt(29.1)
                run.font.bold = True
                run.font.color.rgb = dark_blue

def add_image_annotations(slide, image_path):
    """
    根据图片文件名，在 PPT 幻灯片中添加“图表标题”和“资料来源”。
    文件名格式预期: 分类_标题_日期_来源.jpg
    """
    # ================= 配置区域 (在此调整位置和样式) =================
    # 1. 标题配置 (图片上方)
    TITLE_TOP = 3016459      # 距离顶部距离 (EMU)
    TITLE_LEFT_BASE = 3750684 # 左边距基准
    TITLE_WIDTH = 1616075
    TITLE_HEIGHT = 226581
    TITLE_FONT_SIZE = Pt(12)
    TITLE_FONT_NAME = '华文细黑'
    # 2. 资料来源配置 (图片下方)
    SOURCE_TOP = 6316663       # 距离顶部距离 (EMU)
    SOURCE_LEFT_BASE = 7315200 # 左边距基准
    SOURCE_WIDTH = 1285875
    SOURCE_HEIGHT = 266700
    SOURCE_FONT_SIZE = Pt(9)
    SOURCE_FONT_COLOR = RGBColor(100, 100, 100) # 深灰色
    # ==============================================================

    try:
        file_name = os.path.basename(image_path)
        name_without_ext = os.path.splitext(file_name)[0]
        parts = name_without_ext.split('_')

        # -------------------------------------------------------
        # A. 添加图表标题 (文件名第2部分)
        # -------------------------------------------------------
        if len(parts) >= 2:
            chart_title = parts[1]
            
            # 只有当标题不是 "NONE" 且不为空时才添加
            if chart_title and chart_title != "NONE":
                # 位置微调
                total_title = len(chart_title)
                move_left_amount_a = Pt(0)
                if total_title >12:
                    print(f"    {chart_title}标题较长({total_title}字符)，调整位置")
                    move_left_amount_a = Pt(40)
                elif total_title >20:
                    print(f"    {chart_title}标题较长({total_title}字符)，调整位置")
                    move_left_amount_a = Pt(202)
                t_left = TITLE_LEFT_BASE - move_left_amount_a
                
                # 创建/添加文本框
                title_box = slide.shapes.add_textbox(t_left, TITLE_TOP, TITLE_WIDTH, TITLE_HEIGHT)
                tf = title_box.text_frame
                tf.clear()
                
                p = tf.paragraphs[0]
                run = p.add_run()
                run.text = chart_title
                run.font.name = TITLE_FONT_NAME
                run.font.bold = True
                run.font.size = TITLE_FONT_SIZE
                run.font.color.rgb = RGBColor(0, 0, 0) # 黑色
                
                print(f"    [成功] 已添加标题: {chart_title}")

        # -------------------------------------------------------
        # B. 添加资料来源 (智能解析文件名末尾)
        # -------------------------------------------------------
        move_left_amount = 0
        source_text = ""

        # 解析来源文字逻辑
        if len(parts) >= 2:
            # 格式: 分类_标题_来源 (或更长) -> 取从索引3开始的所有部分

            source_text = parts[-1].replace('，', ' ')  # 只替换中文逗号，不要用join
            source_text = re.sub(r'\s+', ' ', source_text).strip()
            if len(source_text) > 13:
                print(f"    {source_text}来源文字较长({len(source_text)}字符)，调整位置")
                move_left_amount = Pt(40)
            else:
                move_left_amount = Pt(20)
        else:
            # 格式太短，直接用文件名
            source_text = name_without_ext
            move_left_amount = -Pt(15)
        
        # 清洗文字
        source_text = source_text.strip('_ ')
        
        # 写入 PPT
        if source_text:
            final_text = f"资料来源：{source_text}"
            
            s_left = SOURCE_LEFT_BASE - move_left_amount
            
            source_box = slide.shapes.add_textbox(s_left, SOURCE_TOP, SOURCE_WIDTH, SOURCE_HEIGHT)
            tf_source = source_box.text_frame
            # tf_source.clear()
            for paragraph in tf_source.paragraphs:
                    # 确保段落有内容，否则add_run
                    if not paragraph.runs:
                        paragraph.add_run()
                        paragraph.runs[0].text = final_text
                        
                    for run in paragraph.runs:
                        run.font.name = 'Microsoft YaHei'
                        run.font.size = Pt(9) # 10号字体
                        run.font.color.rgb = RGBColor(100, 100, 100) # 深灰色
                  
    
            
            print(f"    [成功] 已添加来源: {source_text}")

    except Exception as e:
        print(f"    [错误] 添加注释信息失败: {e}")
def generate_content_slides(prs, data, content_slide_layout, images_dir, dark_blue):
    """生成内容幻灯片"""
    # 先删除第3页及之后的所有现有幻灯片（如果需要清理旧内容）
    while len(prs.slides) > 2:
        rId = prs.slides._sldIdLst[2].rId
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[2]

    # 为每个内容项创建一个新幻灯片
    for i, content in enumerate(data["content_slides"]):
        # 为前6个内容项，如果有现成的幻灯片可以使用（索引2+i）
        if i < 6 and (2 + i) < len(prs.slides):
            slide = prs.slides[2 + i]
        else:
            # 超过6个内容项或者没有现成幻灯片，添加新幻灯片
            slide = prs.slides.add_slide(content_slide_layout)

        if len(slide.placeholders) >= 2:
            title_placeholder = slide.placeholders[0]
            body_placeholder = slide.placeholders[10]
        else:
            # 如果占位符数量不足，尝试其他方法
            title_placeholder = None
            body_placeholder = None

        # 设置标题
        if title_placeholder and title_placeholder.has_text_frame:
            title_placeholder.text = content["title"]
            # 格式化标题
            text_frame = title_placeholder.text_frame
            if text_frame.paragraphs and len(text_frame.paragraphs) > 0:
                paragraph = text_frame.paragraphs[0]
                if paragraph.runs and len(paragraph.runs) > 0:
                    run = paragraph.runs[0]
                    run.font.name = '华文细黑'
                    run.font.size = Pt(29.1)
                    run.font.bold = True
                    run.font.color.rgb = dark_blue

        # 设置正文内容
        if body_placeholder and body_placeholder.has_text_frame:
            # 添加项目符号
            bullets = content["bullets"]
            body_text = ""
            for bullet in bullets:
                body_text += f"{bullet}\n"
            body_placeholder.text = body_text
            # 根据字数调整字体大小
            total_chars = len(body_text)
            font_size = Pt(14) if len(bullets) > 3 or total_chars > 200 else Pt(16)

            # 遍历所有段落设置字体大小
            for paragraph in body_placeholder.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = font_size

            print(f"Slide {i+1} Body Text added")

        matched_image_path = find_matching_image(content["title"], images_dir, match_length=2)
        if matched_image_path:
            # 尝试在幻灯片中插入图片
            if fill_image(slide, matched_image_path):
                print(f"    成功插入图片: {os.path.basename(matched_image_path)}")
            else:
                print(f"    插入图片失败，幻灯片可能没有图片占位符")
            # try:
            #     # A. 提取资料来源文字
            #     file_name = os.path.basename(matched_image_path)
            #     name_without_ext = os.path.splitext(file_name)[0]
                
            #     parts = name_without_ext.split('_')
    
            add_image_annotations(slide, matched_image_path)
    
        else:
            print(f"    未找到匹配的图片")

            

        # 清理未使用的占位符
        clean_slide_after_insertion(slide)

        print("-" * 50)

def generate_image_slides(prs, content,image_slide_layout, images_dir, dark_blue):
    img_slide = prs.slides.add_slide(image_slide_layout)
    img_slide.placeholders[0].text = content
    matched_image_path = find_matching_image(content, images_dir, match_length=2)
    if matched_image_path:
        # 尝试在幻灯片中插入图片
        if fill_image(img_slide, matched_image_path):
            print(f"    成功插入图片: {os.path.basename(matched_image_path)}")
        else:
            print(f"    插入图片失败，幻灯片可能没有图片占位符")

    else:
        print(f"    未找到匹配的图片")
    if content == "资金流" and matched_image_path:
        add_image_annotations(img_slide, matched_image_path)
    # 清理未使用的占位符
    clean_slide_after_insertion(img_slide)

    print("-" * 50)


def generate_contact_slide(prs, contact_slide_layout2):
    """生成联系我们幻灯片"""
    contact_slide = prs.slides.add_slide(contact_slide_layout2)
    contact_slide.placeholders[0].text = "联系我们"

    # 左侧主文本框
    left_box = contact_slide.placeholders[15]
    tf = left_box.text_frame
    tf.clear()

    p1 = tf.paragraphs[0]

    p1.text = ""  # 先清空

    run = p1.add_run()
    run.text = "www.fargowealth.com"
    run.font.underline = True

    p2 = tf.add_paragraph()
    p2.text = "资产管理服务由集团旗下公司绅士资本提供"
    address_map = {
        10: "香港\n"
            "香港中环康乐广场8号\n"
            "交易广场二期26楼2606-2607室\n"
            "电话：+852 2956 9700",

        17: "新加坡\n"
            "新加坡滨海大道10号\n"
            "滨海湾金融中心2座 #16-05\n"
            "电话：+65 6509 0110",

        19: "Fargo Space\n"
            "香港尖沙咀海港城\n"
            "港威大厦5座33楼3301-3304室\n"
            "电话：+852 2439 9745",

        22: "北京\n"
            "北京市朝阳区东三环中路1号\n"
            "环球金融中心西楼1013室\n"
            "电话：+86 10 6507 8234",

        18: "上海\n"
            "上海市黄浦区太仓路233号\n"
            "新茂大厦2204-2205室\n"
            "电话：+86 21 6333 8131",

        20: "深圳\n"
            "深圳市前海深港合作区兴海大道\n"
            "3040号前海世茂大厦2402室\n"
            "电话：+86 755 2691 3468",
        21: "杭州\n"
            "杭州市上城区新业路228号\n"
            "来福士中心T2办公楼1702-1703室\n"
            "电话：+86 571 8805 8596"
    }

    for idx, text in address_map.items():
        ph = contact_slide.placeholders[idx]
        tf = ph.text_frame
        tf.clear()

        lines = text.split("\n")
        for i, line in enumerate(lines):
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()

            p.text = line

            # 城市名加粗（第一行，或空行后的第一行）
            if line.strip() in ["香港", "新加坡", "Fargo Space", "北京", "上海", "深圳", "杭州"]:
                for run in p.runs:
                    run.font.bold = True
                    run.font.size = Pt(8)
            else:
                for run in p.runs:
                    run.font.size = Pt(8)


def generate_disclaimer_slides(prs, disclaimer_slide_layout, disclaimer_slide_layout2, user_input, dark_blue):
    """生成免责声明幻灯片"""
    # 添加中文免责声明页面
    disclaimer_cn_slide = prs.slides.add_slide(disclaimer_slide_layout)
    # 设置标题
    if len(disclaimer_cn_slide.placeholders) > 0:
        disclaimer_cn_slide.placeholders[0].text = "免责声明"

        # 格式化标题
        text_frame = disclaimer_cn_slide.placeholders[0].text_frame
        if text_frame.paragraphs and len(text_frame.paragraphs) > 0:
            paragraph = text_frame.paragraphs[0]
            if paragraph.runs and len(paragraph.runs) > 0:
                run = paragraph.runs[0]
                run.font.name = '华文细黑'
                run.font.size = Pt(29.1)
                run.font.bold = True
                run.font.color.rgb = dark_blue

    # 设置中文免责声明正文
    if user_input == "香港":
        disclaimer_cn_bullets = [
            "绅士资本有限公司（'绅士资本'）获得香港证券及期货事务监察委员会（'证监会'）颁发的第 9 类受监管活动（资产管理）（中央编号：BIJ793）。",
            "在代表您行使投资酌情权时，绅士资本可能会不时从经纪商处收取现金回扣，以代表您将交易业务交给经纪商。",
            "绅士资本还可能从产品发行人那里获得非金钱利益，例如研究报告、市场分析数据、投资组合分析、培训和研讨会。",
            "绅士资本认为本文件的内容基于被认为可靠的信息来源。 投资和任何产生收入的工具的价值可升可跌。 投资涉及风险，包括本金损失。 不保证股息。 我们的预测是基于过去的表现。 过去的表现并不能保证未来的结果。 投资组合可能会遭受损失并获得收益。 未来回报无法保证，可能会出现本金损失。",
            "绅士资本相信通讯中的信息和内容的来源是可靠的，但它不能也不做任何明示或暗示的保证，并且不对任何信息或数据的准确性、有效性、及时性、适销性或完整性承担任何责任。 任何特定目的或用途，或信息或数据将没有错误。绅士资本不对任何人对本文所表达的任何陈述或意见的任何依赖承担任何责任。绅士资本或其任何关联公司、董事、高级职员或雇员均不对任何人因使用此信息而可能遭受的任何性质的任何损失或损害承担任何责任或任何形式的责任。 不得出于任何目的依赖本通讯中包含的信息或意见。",
            "本通讯中的内容仅适用于继续符合合格和/或认可/专业投资者定义的绅士资本的实际客户。",
            "本通讯的内容是严格保密的，仅供背景和信息之用。绅士资本未经审核内容的真实性、事实准确性或完整性。 内容并不声称是完全或完整的。",
            "未经绅士资本事先书面许可，不得以任何方式抄襲、复制、披露或出版本通讯的任何部分。",
            "本通讯不构成任何发行或出售的要约，或任何订阅或购买任何股份或任何其他利益的要约的招揽，也不构成任何司法管辖区或任何人的任何要约，也不构成任何要约的一部分，也不构成其发行或出售的任何要约的一部分，也不构成任何合同或承诺的基础或与之相关的依赖。"
        ]
    elif user_input == "中国大陆":
        disclaimer_cn_bullets = [
            "华港财富集团有限公司（以下统称“华港财富”），及其下属品牌华港财富不提供法务或税务咨询，本演示文本并不构成此类咨询。出版本演示文本的唯一目的是介绍相关信息，与任何读者的具体投资目标、财务状况或特定需求均无关。华港财富极力建议所有对本演示文本中阐述的产品或服务加以考虑的人士，接受恰当而独立的法务、税务及其它专业咨询。",
            "关于本文信息的精确性、完整性或可靠性，并不提供任何外在或内在的声明或保证，也无意对本资料中的课题进行完整的陈述或总结。虽然本演示文本中阐述的所有信息和意见均来自相信是非常可靠的来源，并且以良好诚信的态度整理而成，但对其精确性或完整性并不提供外在或内在的声明或保证。华港财富及其董事、雇员或代理人，不承担任何由于使用本材料中的全部或部分内容而造成任何损失或伤害的责任。",
            "某些产品和服务由于受到法律限制而无法在世界各地无限制地提供。本演示文本不构成销售要约，也不构成购买或出售任何证券或投资工具的要约邀请，不影响任何交易，也不决定任何形式的任何法律行为。本文任何内容均不应限制或妨碍任何具体报价要约中的特定条款。在任何禁止要约、要约邀请或销售的司法管辖权下，不提供任何关于任一产品的要约；也不向任何人提供这种要约、要约邀请或销售，如果这类做法非法的话。",
            "本材料中表达的任何意见均可能发生变动，不再另行通知。由于使用的假设和标准不同，这些意见可能与华港财富其它业务领域或部门表述的意见发生分歧或矛盾。华港财富没有任何对本文所包含的信息进行更新或保持其时效性的义务。",
            "本演示文本无意也不能用作以下用途：（1）规避《美国国内收入法典》（即美国税法）的惩罚，或者（2）向其他人宣传、营销或推荐任何与税务有关的事宜。",
            "除非华港财富许可，否则不得复制或分发本演示文本。",
            "本演示文本只供专业投资者参考使用，且仅限在香港地区分发和使用。",
            "©华港财富2025. 版权所有。繁体字标识和拼音注释均属于华港财富及其相关联公司注册与未注册的商标。保留所有权利。"
        ]
    else:
        disclaimer_cn_bullets = [
            "本手册的资料，仅供一般资讯用途。对该等资料，华港财富不会就任何错误、遗漏、或错误陈述或失实陈述(不论明示或默示)承担任何责任。对任何因使用或不当使用或依据本手册所记载的资料而引致或所涉及的损失，华港财富不承担任何义务、责任或法律责任。",
            "您有权对本手册上所有文字、照片、图片、标识、内容和其他信息（以下简称“信息”）进行保存、分析、修改、复制以供您个人使用，但未经华港财富书面同意，您不得以任何形式向任何第三方公开发表、传输本手册信息。本手册上显示的所有商标、标识及相关知识产权均为华港财富或其各自所有权所属人所有，未在本手册中明确授予他人的权利均由本公司或向本公司提供信息的第三方保留。华港财富保留不时自主编辑、修改、增加或删除本手册信息的权利。",
            "请注意本手册信息中有很大部分包括或含有第三方提供的信息，本手册并未对该第三方提供的信息作独立核实或确认。基于华港财富的业务活动，本公司经常对合作方以及其他第三方企业的信息存在保密义务，本公司的任何观点及信息披露均受到该保密义务的限制。本公司并不保证本手册观点及信息的真实性、准确性、完整性、时效性或者不存在侵权。",
            "本手册提供的各项信息并非任何形式的证券或其他资产和服务的买卖、发售或购买劝诱，或向您提供投资建议或任何具体建议。您不得将本手册信息作为业务、财务、投资、交易、法律、监管、税收或会计建议，或作为您本人、他人代表您本人、您的会计师或由您管理或受托的帐户进行任何投资决策的主要依据，您应当就任何计划进行的交易征询您的业务顾问、律师、税务及会计顾问的意见，华港财富不会对您使用或依赖本手册信息产生的任何后果负责。",
            "©华港财富2025. 版权所有。繁体字标识和拼音注释均属于华港财富及其相关联公司注册与未注册的商标。保留所有权利"
        ]

    if len(disclaimer_cn_slide.placeholders) > 1:
        body_placeholder = disclaimer_cn_slide.placeholders[12]
        body_text = ""
        for bullet in disclaimer_cn_bullets:
            body_text += f"{bullet}\n"
        body_placeholder.text = body_text

        # 设置小字体以适应内容
        text_frame = body_placeholder.text_frame
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.name = '华文细黑'
                run.font.size = Pt(12)

    # 添加英文免责声明页面
    if user_input == "中国大陆":
        pass  # 中国大陆用户不添加英文免责声明
    else:
        disclaimer_en_slide = prs.slides.add_slide(disclaimer_slide_layout2)

        # 设置标题
        if len(disclaimer_en_slide.placeholders) > 0:
            disclaimer_en_slide.placeholders[0].text = "Disclaimer"

            # 格式化标题
            text_frame = disclaimer_en_slide.placeholders[0].text_frame
            if text_frame.paragraphs and len(text_frame.paragraphs) > 0:
                paragraph = text_frame.paragraphs[0]
                if paragraph.runs and len(paragraph.runs) > 0:
                    run = paragraph.runs[0]
                    run.font.name = '华文细黑'
                    run.font.size = Pt(29.1)
                    run.font.bold = True
                    run.font.color.rgb = dark_blue

        # 设置英文免责声明正文
        if user_input == "香港":
            disclaimer_en_bullets = [
                "Gentleman Capital Limited (\"Gentleman Capital\") is licensed for Type 9 Regulated Activities (Asset Management) by the Securities and Futures Commission (\"SFC\") of Hong Kong (Central Entity No. BIJ793).",
                "When acting for you in the exercise of investment discretion, we may, from time to time, receive cash rebates from a broker in consideration of directing transaction business on behalf of you to the broker.",
                "We may also receive non-monetary benefits, such as research report, market analysis data, portfolio analysis, training and seminars from the product issuer.",
                "Gentleman Capital believes that the contents of this document are based on sources of information believed to be reliable. The value of investments and any income generated instruments may go down as well as up. Investing involves risk, including loss of principal. Dividends are not guaranteed. Our projection is based on past performance. Past performance does not guarantee future results. A portfolio could suffer losses as well as achieve gains. Future returns are not guaranteed and a loss of principal may occur.",
                "Gentleman Capital believes the source of the information and content in the communication to be reliable however it cannot and does not guarantee, either expressly or impliedly, and accepts no liability for the accuracy, validity, timeliness, merchantability or completeness of any information or data for any particular purpose or use or that the information or data will be free from error.",
                "Gentleman Capital does not undertake any responsibility for any reliance which is placed by any person on any statements or opinions which are expressed herein. Neither Gentleman Capital, nor any of its affiliates, directors, officers or employees will be liable or have any responsibility of any kind for any loss or damage of whatever nature that any person may incur resulting from the use of this information. No reliance may be placed for any purpose on the information or opinions contained in this communication.",
                "The content in this communication is intended only for actual clients of Gentleman Capital who continue to meet the definition of qualified and or accredited / professional investors.",
                "The content of this communication is strictly confidential and for background and information purposes only. The content has not been audited for veracity, factual accuracy or completeness by Gentleman Capital. The content does not purport to be full or complete.",
                "No part of this communication may be copied, reproduced, disclosed or published in any manner whatsoever without the prior written permission of Gentleman Capital.",
                "This communication does not constitute or form part of any offer to issue or sell, or any solicitation of any offer to subscribe or purchase, any shares or any other interests in any jurisdiction or to any person, nor shall it or the fact of its distribution form the basis of, or be relied on in connection with, any contract or commitment whatsoever."
            ]
        else:
            disclaimer_en_bullets = [
                "The information in this manual is for general information purposes only. Fargo Wealth shall not be held liable for any errors, omissions, or inaccuracies in such information, whether expressed or implied. Fargo Wealth disclaims any obligation, responsibility, or legal liability for any loss or damage arising from the use or misuse of, or reliance upon, the information contained in this manual.",
                "You are authorized to save, analyze, modify, and copy all the texts, photos, images, logos, content, and other information (hereinafter referred to as \"Information\") in this manual for your personal use. However, without the prior written consent of Fargo Wealth, you are not allowed to publicly disclose or transmit the information in this manual in any form to any third party. All trademarks, logos, and related intellectual property rights displayed in this manual are owned by Fargo Wealth or their respective owners, unless otherwise explicitly granted in this manual. Fargo Wealth reserves the right to independently edit, modify, add, or remove the information in this manual from time to time.",
                "Please note that a significant portion of the information in this manual includes or contains information provided by third parties, and this manual has not independently verified or confirmed the information provided by those third parties. Due to the business activities of Fargo Wealth, the company often has confidentiality obligations with respect to information from partners and other third-party entities. Any views and information disclosed by the company are subject to these confidentiality obligations. The company does not guarantee the truthfulness, accuracy, completeness, timeliness, or absence of infringement of the views and information in this manual.",
                "The information provided in this manual does not constitute any form of solicitation, offer, or inducement to buy, sell, or purchase securities or other assets and services, nor does it provide investment advice or any specific recommendations. You should not consider the information in this manual as business, financial, investment, trading, legal, regulatory, tax, or accounting advice, or as the primary basis for making any investment decisions for yourself, on behalf of others, by your accountant, or for accounts managed or entrusted to you. You should consult your business advisor, lawyer, tax advisor, and accountant for their opinions on any transactions you plan to undertake. Fargo Wealth will not be responsible for any consequences arising from your use of or reliance on the information in this manual.",
                "© Fargo Wealth 2025. All rights reserved. The traditional Chinese characters and Pinyin annotations are trademarks owned by Fargo Wealth and its affiliated companies, whether registered or unregistered. All rights reserved."
                ]

        if len(disclaimer_en_slide.placeholders) > 1:
            body_placeholder = disclaimer_en_slide.placeholders[12]
            body_text = ""
            for bullet in disclaimer_en_bullets:
                body_text += f"{bullet}\n"
            body_placeholder.text = body_text

            # 设置小字体以适应内容
            text_frame = body_placeholder.text_frame
            for paragraph in text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Arial'
                    run.font.size = Pt(9)

    print("联系我们和免责声明页面已添加")


def generate_ppt_from_json(data_json_path, template_path, OUTPUT, user_input,images_dir):
    dark_blue = RGBColor(0, 32, 96)

    prs = Presentation(template_path)

    print(f"模板共有 {len(prs.slide_layouts)} 个布局")
    for i, layout in enumerate(prs.slide_layouts):
        print(f"布局 {i}: 名称='{layout.name}'")

    # ========= 读取 JSON =========
    with open(data_json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    # ========= Slide 0：封面 =========
    generate_cover_slide(prs, data, dark_blue)

    # ========= Slide 1：Executive Summary 表格 =========
    generate_executive_summary_slide(prs, data, dark_blue)

    ## Content slides
    print(f"Number of slide layouts: {len(prs.slide_layouts)}")

    # 保存第三页的幻灯片版式（作为所有内容幻灯片的模板）
    content_slide_layout = prs.slide_layouts[9] 
    imge_slide_layout = prs.slide_layouts[10] 
    contact_slide_layout2 = prs.slide_layouts[2]  # 联系页布局是索引2
    disclaimer_slide_layout = prs.slide_layouts[4]  # cn免责声明页布局是索引2
    disclaimer_slide_layout2 = prs.slide_layouts[3]  # en免责声明页布局是索引4
    for ph in content_slide_layout.placeholders:
        print(
            ph.placeholder_format.idx,
            ph.placeholder_format.type,
            ph.name
        )

    generate_content_slides(prs, data, content_slide_layout, images_dir, dark_blue)
    generate_image_slides(prs, "个债精选", imge_slide_layout, images_dir, dark_blue)
    generate_image_slides(prs, "个股精选", imge_slide_layout, images_dir, dark_blue)
    generate_image_slides(prs, "资金流", imge_slide_layout, images_dir, dark_blue)

    # 4. Contact slide (if specified in user input)
    generate_contact_slide(prs, contact_slide_layout2)

    generate_disclaimer_slides(prs, disclaimer_slide_layout, disclaimer_slide_layout2, user_input, dark_blue)

    # ========= 保存PPT =========
    ai_generate_dir = "ai_generate"
    os.makedirs(ai_generate_dir, exist_ok=True)
    output_path = os.path.join(ai_generate_dir, OUTPUT)

    if os.path.exists(output_path):
        os.remove(output_path)
        print(f"已删除现有的文件: {output_path}")

    prs.save(output_path)
    print(f"PPT已生成并保存为: {output_path}")
    print(f"共创建了 {len(prs.slides)} 页幻灯片")


# Example usage
if __name__ == "__main__":
    images_dir = "input_articles/20260114/images_20260114"  # 图片目录路径
    user_input = input("请选择需要PPT的地点：\n1. 中国大陆\n2. 香港\n3. 新加坡\n请输入数字 (1/2/3): ")

    # 创建映射字典
    location_map = {
        "1": "中国大陆",
        "2": "香港", 
        "3": "新加坡"
    }

    # 验证输入
    if user_input not in location_map:
        print("输入无效，请输入：1、2 或 3")
        exit(1)

    # 获取对应的地点名称
    location_name = location_map[user_input]

    # 根据地点选择模板
    if location_name == "香港":
        TEMPLATE = "template/AI PPT v2.pptx"
    elif location_name == "中国大陆":
        TEMPLATE = "template/AI PPT v3.pptx"
    else:  # 新加坡
        TEMPLATE = "template/AI PPT v2.pptx"

    DATA_JSON = "final_investment_report.json"
    # 使用f-string正确格式化，将地点名称嵌入文件名
    OUTPUT = f"AI_PPT_generated_{location_name}.pptx"

    generate_ppt_from_json(DATA_JSON, TEMPLATE, OUTPUT, location_name,images_dir)
