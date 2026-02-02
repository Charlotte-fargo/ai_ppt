import streamlit as st
import os
import json
import requests
import time
import logging
import re

# --- 引入自定义模块 ---
# 确保 construct_json, ai_prompt, ppt_ready 都在同一目录下
from construct_json import json_main
from AI_prompt_ready import AIPromptRunner
from ppt_ready import PPTGenerator
# --- 引入配置文件 ---
import config

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# ================= 1. 后端逻辑函数 =================

def get_news_platform_token():
    """获取 News Platform Token (使用 config.py 配置)"""
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    data = {
        'client_id': config.NEWS_CLIENT_ID,
        'client_secret': config.NEWS_CLIENT_SECRET,
        'grant_type': 'client_credentials'
    }
    try:
        response = requests.post(config.NEWS_AUTH_URL, headers=headers, data=data, timeout=10)
        response.raise_for_status()
        return response.json().get('access_token')
    except Exception as e:
        logging.error(f"获取 News Token 失败: {e}")
        return None

def fetch_articles(token):
    """抓取文章列表"""
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    try:
        response = requests.get(config.NEWS_ARTICLE_URL, headers=headers, timeout=30)
        if response.status_code == 200:
            return response.json()
        elif response.status_code == 401:
            return "EXPIRED"
        return None
    except Exception as e:
        logging.error(f"文章抓取失败: {e}")
        return None

def save_temp_json(data, filename='articles.json'):
    """保存临时 JSON 数据"""
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        return True
    except Exception as e:
        logging.error(f"保存 {filename} 失败: {e}")
        return False

def choose_template(location_name, language="cn"):
    """根据地点和语言选择模板"""
    # 获取地点对应的模板映射
    location_templates = config.TEMPLATE_MAP.get(location_name, config.TEMPLATE_MAP["香港/Hong Kong"])
    
    # 根据语言选择模板
    template_path = location_templates.get(language, location_templates["cn"])
    
    # 检查文件是否存在，如果不存在则使用默认语言
    if not os.path.exists(template_path):
        template_path = location_templates["cn"]
    
    return template_path

def get_language(language):
    """根据 config 映射获取模板路径"""
    # 默认为香港模板
    rel_language = config.LANGUAGE_MAP.get(language, config.LANGUAGE_MAP["中文/Chinese"])
    # Streamlit 中直接使用相对路径通常没问题
    return rel_language
def clean_html_content(html_content):
    """
    清洗 HTML 文本：去除 img 标签和资料来源段落
    """
    if not isinstance(html_content, str):
        return html_content

    # 1. 匹配 < img ...> 标签 (包括跨行的属性)
    # re.DOTALL 让 . 也能匹配换行符
    img_pattern = re.compile(r'<img[^>]+>', re.IGNORECASE | re.DOTALL)
    
    # 2. 匹配包含 "资料来源" 的 <p> 段落
    # 逻辑：匹配以 <p (或 <p >) 开始，中间内容包含 "资料来源"，直到 </p > 结束
    # 这样可以连带删除包裹它的 <span ...> 等标签
    source_pattern = re.compile(r'<p[^>]*>.*?资料来源.*?</p >', re.IGNORECASE | re.DOTALL)

    # 执行替换
    content = img_pattern.sub('', html_content)     # 删图片
    content = source_pattern.sub('', content)       # 删资料来源
    
    return content
def process_single_file(file_path, save_path):
    """
    读取单个文件，清洗数据，并保存
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        # 构建清洗后的字典结构，只保留需要的字段
        cleaned_data = {
            "titles": data.get("titles", {}),
            "summaries": data.get("summaries", {}),
            "contents": {} # 稍后填充
        }

        # 处理 contents
        raw_contents = data.get("contents", {})
        if raw_contents:
            for lang, html_text in raw_contents.items():
                cleaned_data["contents"][lang] = clean_html_content(html_text)

        # 保存到输出文件夹
        with open(save_path, 'w', encoding='utf-8') as f:
            json.dump(cleaned_data, f, ensure_ascii=False, indent=2)
            
        print(f"[成功] 已清洗: {os.path.basename(file_path)}")

    except json.JSONDecodeError:
        print(f"[跳过] 文件格式错误 (非标准JSON): {os.path.basename(file_path)}")
    except Exception as e:
        print(f"[错误] 处理 {os.path.basename(file_path)} 时出错: {str(e)}")

def batch_process(INPUT_FOLDER, OUTPUT_FOLDER):
    # 1. 检查输入文件夹是否存在
    if not os.path.exists(INPUT_FOLDER):
        print(f"错误：找不到输入文件夹 '{INPUT_FOLDER}'，请先创建并放入 JSON 文件。")
        return

    # 2. 如果输出文件夹不存在，自动创建
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)
        print(f"已创建输出文件夹: {OUTPUT_FOLDER}")

    # 3. 获取所有 JSON 文件
    files = [f for f in os.listdir(INPUT_FOLDER) if f.lower().endswith('.json')]
    
    if not files:
        print(f"在 '{INPUT_FOLDER}' 中没有找到 .json 文件。")
        return

    print(f"开始处理，共发现 {len(files)} 个文件...\n")

    # 4. 循环处理
    for filename in files:
        input_path = os.path.join(INPUT_FOLDER, filename)
        output_path = os.path.join(OUTPUT_FOLDER, filename)
        process_single_file(input_path, output_path)

    print(f"\n全部完成！清洗后的文件在 '{OUTPUT_FOLDER}' 文件夹中。")
# ================= 2. 密码验证逻辑 =================

def check_password():
    """简单的密码保护"""
    if st.session_state.get('password_correct', False):
        return True

    st.header("🔒 请登录")
    password_input = st.text_input("请输入访问密码", type="password")
    
    if st.button("登录"):
        # 这里硬编码密码，您可以根据需要修改
        if password_input == config.APP_PASSWORD:  
            st.session_state['password_correct'] = True
            st.rerun()
        else:
            st.error("密码错误，请重试")
            
    return False

# ================= 3. Streamlit 主界面 =================

def main_app():
    # --- 0. 页面样式优化 (CSS) ---
    st.markdown("""
        <style>
        .stButton>button {
            height: 3em;
            font-size: 20px;
            font-weight: bold;
        }
        .reportview-container .main .block-container{
            padding-top: 2rem;
        }
        </style>
    """, unsafe_allow_html=True)

    # --- 1. 标题区域 ---
    col_logo, col_title = st.columns([1, 5], gap="medium", vertical_alignment="center")
    
    with col_logo:
        # use_container_width=True 让图片自动填满这 1 份的宽度，不用手动设 width
        st.image("logo.png", width='stretch') 
        
    with col_title:
        # 使用 markdown 的 # 号，并去除默认的 margin (空白)，让它和 Logo 贴得更紧
        st.markdown(
            """
            <h1 style='margin-bottom: 0px; margin-top: 0px;'>CIO Office 自动化报告系统</h1>
            <p style='font-size: 16px; color: gray; margin-top: -5px;'>Automated Investment Report Generator</p>
            """, 
            unsafe_allow_html=True
        )

    st.markdown("---")
    # 2. 设置区 (使用两列布局，解决“乱”的问题)
    with st.container():
        st.subheader("1. 设置 / Settings")
        c1, c2 = st.columns(2)
        
        with c1:
            location_name = st.radio(
                "📍 目标地点 / Destination:",
                ("中国大陆/China", "香港/Hong Kong", "新加坡/Singapore"),
                index=0,
                horizontal=True
            )
            
        with c2:
            language = st.radio(
                "🗣️ 目标语言 / Language:",
                ("中文/Chinese", "英文/English"),
                index=0,
                horizontal=True
            )

    st.markdown("---")

    # 3. 执行区
    st.subheader("2. 执行 / Execute")
    
    # 一个醒目的大按钮
    start_btn = st.button("🚀 开始生成 PPT / Start Generation", type="primary", use_container_width=True)
    
    if start_btn:
        # --- 这里改回了你想要的简单进度条模式 ---
        status_text = st.empty()
        progress_bar = st.progress(0)
        
        try:
            # === Step 1 ===
            status_text.markdown("**Step 1/4:** 正在连接 News Platform 获取数据... (Connecting...)")
            progress_bar.progress(10)
            
            token = get_news_platform_token()
            if not token:
                st.error("❌ 无法获取 News Token")
                return

            articles = fetch_articles(token)
            if not articles or articles == "EXPIRED":
                st.error("❌ 文章列表为空或 Token 失效")
                return
            
            save_temp_json(articles, 'articles.json')
            progress_bar.progress(30)

            # === Step 2 ===
            status_text.markdown("**Step 2/4:** 正在下载图片并整理素材... (Downloading images...)")
            # json_main logic...
            articles_dir, images_dir = json_main("articles.json")
            
            if not articles_dir or not os.path.exists(articles_dir):
                st.error("❌ 文件处理失败")
                return
            progress_bar.progress(50)

            # === Step 3 ===
            status_text.markdown("**Step 3/4:** AI 正在撰写报告，请稍候... (AI Writing...)")
            
            language_code = get_language(language)
            print(f"Init AIPromptRunner with language={language_code}")
            articles_dir_cleaned = "cleaned_articles"
            os.makedirs(articles_dir_cleaned, exist_ok=True)
            batch_process(articles_dir, config.CLEANED_DIR)
            runner = AIPromptRunner(language=language_code)
            final_json_data = runner.run(specific_folder=config.CLEANED_DIR)
            
            if not final_json_data:
                st.error("❌ AI 生成失败")
                return
            
            progress_bar.progress(80)

            # === Step 4 ===
            status_text.markdown(f"**Step 4/4:** 正在渲染 {location_name} 版 PPT... (Rendering PPT...)")
            
            template_path = choose_template(location_name, language_code)
            output_filename = f"AI_PPT_generated_{location_name}_{language_code}.pptx"
            final_output_path = os.path.join(config.OUTPUT_DIR, output_filename)
            
            os.makedirs(config.OUTPUT_DIR, exist_ok=True)
            
            generator = PPTGenerator(final_json_data, template_path, images_dir, location_name, language=language_code)
            success = generator.run(final_output_path)
            
            if success:
                progress_bar.progress(100)
                status_text.success("✅ PPT 生成完成！(Generation Complete)")
                
                # 生成成功后的下载按钮
                real_file_path = os.path.join(config.OUTPUT_DIR, output_filename)
                
                if os.path.exists(real_file_path):
                    with open(real_file_path, "rb") as file:
                        st.download_button(
                            label=f"📥 点击下载: {output_filename}",
                            data=file,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True,
                            type="primary"
                        )
                else:
                    st.error("❌ 文件生成路径异常")
            else:
                st.error("❌ PPT 生成过程中发生错误")

        except Exception as e:
            st.error(f"❌ 发生异常: {str(e)}")
            logging.exception("运行出错")

# 入口保持不变
if __name__ == "__main__":
    st.set_page_config(page_title="EasyView Report", page_icon="📊", layout="centered")
    if check_password():
        main_app()
