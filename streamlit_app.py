import requests
import json
import streamlit as st
import os
import time
from construct_json import json_main
from ppt_ready import generate_ppt_from_json
from ai_prompt import (load_json_files_from_folder,
                       prepare_context_from_files,
                       run_ai_job,
                       poll_result,
                       extract_final_json, get_access_token_b)

# ================= 1. 配置区域 (基于你提供的 curl) =================
AUTH_URL_a = "https://auth.easyview.xyz/realms/Easyview-News-Platform-Realm/protocol/openid-connect/token"
ARTICLE_URL_a = "https://news-platform.easyview.xyz/api/v1/channel/cio/articles"
CLIENT_ID_a = "cio-backend"
CLIENT_SECRET_a = "4cbb1527-bcc4-42ae-a7ec-691359f3e119"

AUTH_URL_b = "https://auth-v2.easyview.xyz/realms/evhk/protocol/openid-connect/token"
API_BASE_URL_b = "https://api-v2.easyview.xyz/v3/ai"
CLIENT_ID_b = "cioinsight-api-client"
CLIENT_SECRET_b = "b02fe9e7-36e6-4c81-a389-9399184eda9b"

location_map = {
    "1": "中国大陆",
    "2": "香港",
    "3": "新加坡"
}

import sys
import os

def get_resource_path(relative_path):
    """
    获取资源的绝对路径。
    用于解决打包成 exe 后，程序无法找到内部文件的问题。
    """
    if hasattr(sys, '_MEIPASS'):
        # PyInstaller 会把文件解压到 sys._MEIPASS 指向的临时目录
        return os.path.join(sys._MEIPASS, relative_path)
    
    # 正常开发环境下，就使用当前目录
    return os.path.join(os.path.abspath("."), relative_path)

def select_location():
    """
    与用户交互选择地区（中国大陆/香港/新加坡）。

    Returns:
        str | None:
            - 返回地区中文名（如 "香港" / "中国大陆" / "新加坡"）
            - 若输入不合法则返回 None（并打印提示）

    Side Effects:
        - 使用 input() 读取控制台输入
        - 使用 print() 输出提示

    Notes:
        - 地区映射来自全局变量 location_map
        - 调用方通常需要对 None 做重试或直接退出处理
    """
    user_input = input("请选择需要PPT的地点：\n1. 中国大陆\n2. 香港\n3. 新加坡\n请输入数字 (1/2/3): ")
    if user_input not in location_map:
        print("输入无效，请输入：1、2 或 3")
        return None
    return location_map[user_input]


def choose_template(location_name):
    """
    根据地区选择 PPT 模板路径。

    Args:
        location_name (str): 地区名称（"香港" / "中国大陆" / "新加坡"）

    Returns:
        str: 模板文件路径
            - 香港 -> template/AI PPT v2.pptx
            - 中国大陆 -> template/AI PPT v3.pptx
            - 其他 -> template/AI PPT v2.pptx（默认）

    Notes:
        - 该函数只负责返回路径，不校验文件是否存在
        - 若你后续新增地区，可在此扩展映射规则
    """
    if location_name == "香港":
        TEMPLATE = get_resource_path(os.path.join("template", "AI PPT v2.pptx"))
        return TEMPLATE
    if location_name == "中国大陆":
        TEMPLATE = get_resource_path(os.path.join("template", "AI PPT v3.pptx"))
        return TEMPLATE
    return get_resource_path(os.path.join("template", "AI PPT v2.pptx"))

def get_access_token():
    """
    向 News Platform 的 OIDC/Keycloak Token Endpoint 申请 access_token（client_credentials 模式）。

    This function is used for Service-to-Service authentication:
    - It posts form data to AUTH_URL_a
    - Obtains an access_token used as `Authorization: Bearer <token>` for subsequent API calls

    Returns:
        str | None:
            - 成功：返回 access_token 字符串
            - 失败：返回 None（并打印错误与可能的服务器返回内容）

    Side Effects:
        - 发起 HTTP POST 请求到 AUTH_URL_a
        - 使用 print() 输出申请进度与错误信息（便于命令行运行时排查）

    Dependencies / Globals:
        - requests
        - AUTH_URL_a, CLIENT_ID_a, CLIENT_SECRET_a（建议改为环境变量读取，避免明文密钥）
        - 固定使用 grant_type = "client_credentials"

    Error Handling:
        - response.raise_for_status() 会在 4xx/5xx 时抛出异常
        - 捕获所有异常并打印；若 response 已存在，则输出 response.text 方便定位

    Security Notes:
        - 不要打印 client_secret 或完整 token
        - 若该文件会提交到仓库，请移除硬编码的 CLIENT_SECRET_a
    """
    headers = {'Content-Type': 'application/x-www-form-urlencoded'}
    data = {
        'client_id': CLIENT_ID_a,
        'client_secret': CLIENT_SECRET_a,
        'grant_type': 'client_credentials'
    }
    try:
        print(" 正在申请 Token...")
        response = requests.post(AUTH_URL_a, headers=headers, data=data)
        response.raise_for_status()
        token_info = response.json()
        token = token_info.get('access_token')
        print(f" Token 获取成功! (有效期: {token_info.get('expires_in')} 秒)")
        return token
    except Exception as e:
        print(f"获取 Token 失败: {e}")
        if 'response' in locals():
            print(f"服务器返回: {response.text}")
        return None


def fetch_articles(token):
    """
    使用 News Platform access_token 拉取 CIO 频道文章列表。

    Args:
        token (str):
            通过 get_access_token() 获取的 access_token。
            将以 `Authorization: Bearer <token>` 的形式携带。

    Returns:
        dict | str | None:
            - 成功（HTTP 200）：返回 response.json()（通常包含 articles 列表等字段）
            - Token 失效（HTTP 401）：返回字符串 "EXPIRED"（用于上层触发重新取 token）
            - 其他错误：返回 None

    Side Effects:
        - 发起 HTTP GET 请求到 ARTICLE_URL_a
        - 使用 print() 输出运行日志与部分响应预览（便于排查）

    Dependencies / Globals:
        - requests
        - ARTICLE_URL_a（文章列表接口地址）
        - token 由外部传入

    Notes / Pitfalls:
        - 当前未设置 timeout，网络抖动时可能阻塞；建议加 timeout=10
        - `response.text[:200]` 仅用于调试，生产环境可关闭以避免日志过大/敏感信息泄露
        - 返回值包含三种类型（dict/str/None），调用方需显式分支处理
    """
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    try:
        print(" 正在获取文章列表...")
        response = requests.get(ARTICLE_URL_a, headers=headers)
        if response.status_code == 200:
            print(" 成功获取文章数据！")
            print("数据预览:", response.text[:200] + "...")
            return response.json()
        elif response.status_code == 401:
            print("Token 失效了")
            return "EXPIRED"
        else:
            print(f"请求文章出错: {response.status_code}")
            return None
    except Exception as e:
        print(f" 连接错误: {e}")
        return None


def save_json_file(path, data):
    try:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        print(f"Saved JSON to {path}")
    except Exception as e:
        print(f"保存 JSON 失败: {e}")


def prepare_files_and_context():
    articles_dir, images_dir = json_main("articles.json")
    print("--- 步骤 1: 读取本地文件 ---")
    uploaded_files = load_json_files_from_folder(articles_dir)
    if len(uploaded_files) == 0:
        print("没有读取到文件，程序结束。请检查 FOLDER_PATH 路径设置。")
        return None, None, None
    context = prepare_context_from_files(uploaded_files)
    print(f"\n已合并文档内容，总字符数: {len(context)}")
    return context, articles_dir, images_dir


def run_ai_pipeline(context):
    print("\n--- 步骤 2: 获取 API Token ---")
    token_b = get_access_token_b(CLIENT_ID_b, CLIENT_SECRET_b)
    if not token_b:
        print("获取 AI 服务 Token 失败，跳过 AI 生成。")
        return None
    print("\n--- 步骤 3: 调用 AI 生成报告 ---")
    job_id = run_ai_job(token_b, context, API_BASE_URL_b)
    if not job_id:
        print("AI 任务提交失败。")
        return None
    result = poll_result(token_b, job_id, API_BASE_URL_b)
    if not result:
        print("AI 任务无结果。")
        return None
    final_report = extract_final_json(result)
    final_report_path = "final_investment_report.json"
    if final_report:
        print("\n\n====== 完美清洗后的 JSON 报告 ======\n")
        print(json.dumps(final_report, ensure_ascii=False, indent=4))
        save_json_file(final_report_path, final_report)
        print(f"\n报告已保存为 {final_report_path}")
        return final_report_path
    print("提取数据失败，请检查日志。")
    return None



# ================= 密码验证函数 =================
def check_password():
    """如果不正确，只显示登录框；如果正确，才显示主程序"""
    
    # 1. 如果已经在会话中登录过，直接放行
    if st.session_state.get('password_correct', False):
        return True

    # 2. 显示登录输入框
    st.header("🔒 请登录")
    password_input = st.text_input("请输入访问密码", type="password")
    
    # 3. 验证逻辑
    if st.button("登录"):
        # 这里设置您的密码，比如 "888888"
        # 更安全的方式是读取环境变量，但在代码里写死也可以
        if password_input == "123456":  
            st.session_state['password_correct'] = True
            st.rerun()  # 重新刷新页面，进入主程序
        else:
            st.error("密码错误，请重试")
            
    return False

# ================= 主程序封装 =================



def main_app():
    # 1. 页面配置 (设置网页标题和图标)
    st.set_page_config(
        page_title="EasyView 报告生成器",
        page_icon="📱",
        layout="centered"
    )

    # 2. 标题和 Logo
    st.image("logo.png", width=200) # 如果没有 logo.png 请注释掉这行
    st.title("EasyView 自动化报告系统")
    st.markdown("---")

    # 3. 手机端控制区
    st.header("1. 设置")
    
    # 手机上用大按钮更好按
    location_name = st.radio(
        "请选择 PPT 目标地点:",
        ("中国大陆", "香港", "新加坡"),
        index=0,
        horizontal=True
    )
    st.sidebar.success("已登录") # 侧边栏显示状态
    # 4. 运行按钮
    st.header("2. 执行")
    # 使用 type="primary" 让按钮变色，更显眼
    if st.button("🚀 开始生成 PPT", type="primary", use_container_width=True):
        
        # 进度条和状态文字
        status_text = st.empty()
        progress_bar = st.progress(0)
        
        try:
            # --- 阶段 1: 准备 ---
            status_text.text("正在连接服务器获取数据...")
            progress_bar.progress(10)
            token = get_access_token()
            
            if token:
                articles = fetch_articles(token)
                progress_bar.progress(30)
                save_json_file('articles.json', articles)
            else:
                st.error("无法连接到数据源")
                return

            # --- 阶段 2: 处理 ---
            status_text.text("正在下载图片并处理上下文...")
            context, articles_dir, images_dir = prepare_files_and_context()
            progress_bar.progress(50)

            # --- 阶段 3: AI 生成 ---
            status_text.text("AI 正在撰写报告 (这可能需要一分钟)...")
            final_report_path = run_ai_pipeline(context)
            progress_bar.progress(80)

            # --- 阶段 4: 生成 PPT ---
            status_text.text(f"正在生成 {location_name} 版本的 PPT...")
            TEMPLATE = choose_template(location_name)
            OUTPUT = f"AI_PPT_generated_{location_name}.pptx"
            
            generate_ppt_from_json(final_report_path, TEMPLATE, OUTPUT, location_name, images_dir)
            
            progress_bar.progress(100)
            status_text.success("✅ PPT 生成完成！")

        
            real_file_path = os.path.join("ai_generate", OUTPUT) 

            # 4. 检查文件是否存在（防止报错）
            if os.path.exists(real_file_path):
                with open(real_file_path, "rb") as file:
                    st.download_button(
                        label=f"📥 点击下载: {OUTPUT}",
                        data=file,
                        file_name=OUTPUT,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True
                    )
            else:
                st.error(f"错误：找不到生成的文件。请检查 {real_file_path} 是否存在。")

        except Exception as e:
            st.error(f"发生错误: {str(e)}")
            # 打印详细错误方便调试
            # st.exception(e)

if __name__ == "__main__":
    # main()
    st.set_page_config(page_title="EasyView", page_icon="🔒")
    
    # 先检查密码，通过了才运行 main_app
    if check_password():
        main_app()
