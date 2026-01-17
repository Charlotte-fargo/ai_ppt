import streamlit as st
import os
import json
import requests
import time
import logging

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

def get_template_path(location_name):
    """根据 config 映射获取模板路径"""
    # 默认为香港模板
    rel_path = config.TEMPLATE_MAP.get(location_name, config.TEMPLATE_MAP["香港"])
    # Streamlit 中直接使用相对路径通常没问题
    return os.path.abspath(rel_path)

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
    # 1. 标题和 Logo
    # st.image("logo.png", width=200) # 如有 logo 可解开注释
    st.title("EasyView 自动化报告系统")
    st.markdown("---")

    # 2. 设置区
    st.header("1. 设置")
    
    # 地点选择 (使用 config 中的 Key)
    location_name = st.radio(
        "请选择 PPT 目标地点:",
        ("中国大陆", "香港", "新加坡"),
        index=0,
        horizontal=True
    )
    
    st.sidebar.success("✅ 已登录")
    
    # 3. 执行区
    st.header("2. 执行")
    
    if st.button("🚀 开始生成 PPT", type="primary", use_container_width=True):
        
        # 初始化进度条
        status_text = st.empty()
        progress_bar = st.progress(0)
        
        try:
            # --- 阶段 1: 抓取文章 ---
            status_text.text("Step 1/4: 连接 News Platform 获取数据...")
            progress_bar.progress(10)
            
            token = get_news_platform_token()
            if not token:
                st.error("无法获取 News Token，请检查 config.py")
                return

            articles = fetch_articles(token)
            if not articles or articles == "EXPIRED":
                st.error("文章列表为空或 Token 失效")
                return
            
            save_temp_json(articles, 'articles.json')
            progress_bar.progress(30)

            # --- 阶段 2: 处理素材 ---
            status_text.text("Step 2/4: 下载图片并整理素材...")
            # json_main 返回处理后的文章目录路径
            articles_dir, images_dir = json_main("articles.json")
            
            if not articles_dir or not os.path.exists(articles_dir):
                st.error("文件处理失败，无法生成文章目录")
                return
            progress_bar.progress(50)

            # --- 阶段 3: AI 生成 ---
            status_text.text("Step 3/4: AI 正在撰写报告 (需约 1-2 分钟)...")
            
            # 实例化 Runner (自动从 config 读取 Token)
            runner = AIPromptRunner()
            # 运行 AI 任务
            final_json_data = runner.run(specific_folder=articles_dir)
            
            if not final_json_data:
                st.error("AI 生成失败，请查看后台日志")
                return
            
            # Runner 默认保存为 final_investment_report.json
            report_path = final_json_data
            progress_bar.progress(80)

            # --- 阶段 4: 生成 PPT ---
            status_text.text(f"Step 4/4: 正在渲染 {location_name} 版 PPT...")
            
            template_path = get_template_path(location_name)
            output_filename = f"AI_PPT_generated_{location_name}.pptx"
            
            # 确保输出目录存在 (使用 config.OUTPUT_DIR)
            os.makedirs(config.OUTPUT_DIR, exist_ok=True)
            generator = PPTGenerator(final_json_data, template_path, images_dir, location_name)
        
            # 假设 run 方法接收输出路径
            success = generator.run(final_output_path)
            
            if success:
                progress_bar.progress(100)
                status_text.success("✅ PPT 生成完成！")
                
                # 构建下载路径
                real_file_path = os.path.join(config.OUTPUT_DIR, output_filename)
                
                if os.path.exists(real_file_path):
                    with open(real_file_path, "rb") as file:
                        st.download_button(
                            label=f"📥 点击下载: {output_filename}",
                            data=file,
                            file_name=output_filename,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            use_container_width=True
                        )
                else:
                    st.error(f"错误：找不到生成的文件 {real_file_path}")
            else:
                st.error("PPT 生成过程中发生错误")

        except Exception as e:
            st.error(f"发生未捕获异常: {str(e)}")
            logging.exception("运行出错")

# ================= 程序入口 =================

if __name__ == "__main__":
    # 配置页面属性
    st.set_page_config(
        page_title="EasyView 报告生成器",
        page_icon="📊",
        layout="centered"
    )
    
    # 检查 config.py 是否配置
    if "在此处填入" in config.API_TOKEN:
        st.warning("⚠️ 警告: config.py 中的 API_TOKEN 尚未配置！")

    # 密码验证通过后才显示主程序
    if check_password():
        main_app()
