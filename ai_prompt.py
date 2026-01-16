import requests
import json
import re
import time
import uuid
import os
import glob  # 新增库：用于查找文件
import base64
# ================= 配置区域 =================

# 文件夹路径 (可以是绝对路径，也可以是相对路径)

FOLDER_PATH = "input_articles/20260109/articles_20260109/" 

# 1. AI 服务的认证地址 (注意：是 auth-v2 和 evhk)
AUTH_URL = "https://auth-v2.easyview.xyz/realms/evhk/protocol/openid-connect/token"

# 2. AI 接口地址
API_BASE_URL = "https://api-v2.easyview.xyz/v3/ai"

# 3. AI 服务的专用凭证 
CLIENT_ID = "cioinsight-api-client"
CLIENT_SECRET = "b02fe9e7-36e6-4c81-a389-9399184eda9b"

# ================= 1. 数据处理部分 =================

def load_json_files_from_folder(folder_path):
    """
    读取指定文件夹下的所有 .json 文件
    返回一个字典：{'文件名': JSON对象, ...}
    """
    data_dict = {}
    
    # 检查文件夹是否存在
    if not os.path.exists(folder_path):
        print(f" 错误：找不到文件夹 '{folder_path}'")
        return {}

    # 查找所有 .json 文件
    # os.path.join 确保路径拼接在不同系统(Windows/Mac)都正确
    json_pattern = os.path.join(folder_path, "*.json")
    file_list = glob.glob(json_pattern)
    
    print(f" 在 '{folder_path}' 下找到了 {len(file_list)} 个 JSON 文件。")

    for file_path in file_list:
        try:
            # 获取文件名 (例如 "9575_债市.json")
            file_name = os.path.basename(file_path)
            
            with open(file_path, 'r', encoding='utf-8') as f:
                content = json.load(f)
                data_dict[file_name] = content
                print(f"  - 已读取: {file_name}")
        except Exception as e:
            print(f"  -  读取失败 {file_name}: {e}")
            
    return data_dict

def clean_html(raw_html):
    """清除 HTML 标签，保留纯文本"""
    if not raw_html:
        return ""
    cleanr = re.compile('<.*?>')
    text = re.sub(cleanr, '', raw_html)
    return text.strip()

def prepare_context_from_files(files_data):
    """将多个文件的内容合并成一个上下文文本"""
    context_str = "以下是各资产类别的原始分析报告内容：\n\n"
    
    if not files_data:
        return ""

    for filename, content_json in files_data.items():
        # 提取标题，做一些容错处理
        titles = content_json.get("titles", {})
        title = titles.get("zh_CN", "未知标题") if isinstance(titles, dict) else "未知标题"
        
        # 提取HTML内容
        contents = content_json.get("contents", {})
        html_content = contents.get("zh_CN", "") if isinstance(contents, dict) else ""
        
        # 清洗HTML
        pure_text = clean_html(html_content)
        
        context_str += f"--- 文档开始: {filename} (标题: {title}) ---\n"
        context_str += pure_text + "\n"
        context_str += f"--- 文档结束 ---\n\n"
        
    return context_str

# ================= 2. API 调用部分 =================

def get_access_token_b(CLIENT_ID, CLIENT_SECRET):
    payload = {
        'grant_type': 'client_credentials',
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET
    }
    try:
        resp = requests.post(AUTH_URL, data=payload)
        resp.raise_for_status()
        return resp.json().get('access_token')
    except Exception as e:
        print(f" 认证失败: {e}")
        return None

def run_ai_job(token, context_text, API_BASE_URL):
    # 1. 基础检查
    if not context_text:
        print("⚠️ 没有提取到任何文本内容，取消 AI 任务。")
        return None

    url = f"{API_BASE_URL}/job"
    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {token}'
    }
    
    # 2. Prompt 设计 (指令)
    system_prompt = """
    你是一个专业的中文首席投资官助理。你需要阅读提供的金融市场分析文档，并生成一份标准化的投资观点报告。
    
    任务要求：
    1. 生成7种资产的投资观点（中港股市、美股、欧股、日股、债市、黄金、原油）。如果提供的文档中缺少某种资产，请根据你的知识库合理推断或标记为"暂无数据"。
    2. 中港股市和黄金的投资逻辑中文字数必须在80字左右。
    3. 美股，欧股投资逻辑的字数控制在55字以内。
    4. 原油，日股和债市的投资逻辑的字数控制在50字以内。
    5. 以下生成的每一个bullet point字数控制在50字左右，三个bullet point总字数需要在150字左右。
    
    硬性写作要求：
    - 标题格式为“资产类别名称：xxxxx”
    - 观点内容不超过三句 bullet point。
    - 每一句观点的格式为“小标题：xxxx”。
    - 语言专业、简练。
    - 标题需要抓住核心结论，点明关键驱动因素。

    最后，请仅输出一个纯净的 JSON 格式，不要包含Markdown标记（如 ```json）。JSON结构如下：
    {
      "document": { "title": "环球市场投资观点", "author":"CIO Office", "date": "..." },
      "executive_summary": { 
          "columns": ["资产类别", "投资逻辑"], 
          "rows": [ {"资产类别": "...", "投资逻辑": "..."} ] 
      },
      "content_slides": [ 
          { "title": "...", "bullets": ["...", "..."] } 
      ]
    }
    """

    # 3. 将 指令 + 内容 合并，并进行 Base64 编码
    #    这样做的好处是避免指令过长导致 parameter 字段放不下
    final_content = f"{system_prompt}\n\n========== 以下是分析文档内容 ==========\n{context_text}"
    
    try:
        encoded_text = base64.b64encode(final_content.encode('utf-8')).decode('utf-8')
    except Exception as e:
        print(f" 编码错误: {e}")
        return None

    # 4. 构建 Payload
    payload = {
        "type": "callLlm", 
        "metadata": {
            "tenantId": "GOLDHORSE",
            "clientId": "CIO",
            "userId": "script_runner",
            "priority": 1,
            "custom": {}
        },
        "input": {
            "parameter": {
                "model_name": "gemini-3-pro-preview",
                # 指引 AI 去读文件
                "prompt": "请详细阅读附带的文件资源（resource），文件中包含了身份设定、具体指令以及需要分析的金融文档内容。请严格按照文件中的 JSON 格式要求输出结果。"
            },
            "resource": [
                {
                    # 给虚拟文件起个名字
                    "name": "instruction_and_context.txt",
                    "resource": {
                        "dataStatus": "NEW",
                        "mimeType": "text/plain",
                        "encoding": "BASE64",
                        "data": encoded_text
                    }
                }
            ]
        },
        "callback": []
    }

    # 5. 打印调试信息并发送 (修正了这里的变量名)
    print(f"\n🚀 [DEBUG] 正在提交任务，文本长度: {len(context_text)} 字符")
    
    try:
        resp = requests.post(url, headers=headers, json=payload)
        
        # 打印一下返回值，万一报错能看到原因
        if resp.status_code != 200:
            print(f"❌ 提交失败 (状态码 {resp.status_code}): {resp.text}")
            return None

        response_data = resp.json()
        
        # 优先拿 uuid，没有才拿 id
        # job_id = response_data.get("uuid") or response_data.get("id")
        job_id = response_data.get("id")
        
        # 如果万一没有 id，再拿 uuid 做备选（虽然这种情况很少见）
        if not job_id:
            job_id = response_data.get("uuid")

        print(f"✅ 任务提交成功! 任务 ID: {job_id}")
        return job_id
        

    except Exception as e:
        print(f"❌ 请求异常: {e}")
        return None


# def poll_result(token, job_id,API_BASE_URL):
#     url = f"{API_BASE_URL}/job/JOB_ID/{job_id}"
#     headers = {'Authorization': f'Bearer {token}'}
    
#     print(" AI 正在生成报告 (可能需要 30-60 秒)...")
#     for _ in range(30):
#         resp = requests.get(url, headers=headers)
#         if resp.status_code == 200:
#             data = resp.json()
#             status = data.get("status")
#             if status in ["SUCCESS", "COMPLETED"]:
#                 return data
#             if status == "FAILED":
#                 print(" 任务处理失败")
#                 return None
#         print(".", end="", flush=True)
#         time.sleep(3)
#     return None
def poll_result(token, job_id, API_BASE_URL):
    # ✅ 确保 URL 拼接正确 (参考之前的文档，必须包含 /JOB_ID/)
    url = f"{API_BASE_URL}/job/JOB_ID/{job_id}"
    
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }

    print(f"🔍 开始轮询任务结果: {url}")

    # 循环 60 次，每次等待 3 秒，最长等待 3 分钟
    for i in range(60):
        try:
            resp = requests.get(url, headers=headers)
            
            if resp.status_code == 200:
                data = resp.json()
                status = data.get("status")
                
                # 打印当前状态，方便观察进度
                print(f"   [第 {i+1} 次查询] 状态: {status}")

                if status in ["SUCCESS", "COMPLETED"]:
                    print("✅ AI 任务成功！")
                    return data
                
                if status == "FAILED" or status == "ERROR":
                    print("\n❌ 任务处理失败！服务端返回信息如下：")
                    # 👇 这一行最关键！把它打印出来发给我
                    print(data) 
                    break
            else:
                print(f"   ⚠️ 查询接口报错: {resp.status_code} - {resp.text}")
                
        except Exception as e:
            print(f"   ⚠️ 请求异常: {e}")

        time.sleep(3) # 等待 3 秒再查

    print("\n⏰ 等待超时，AI 还没跑完。")
    return None
def extract_final_json(api_response):
    """
    专门针对 LangGraph 日志格式进行清洗
    1. 遍历日志，找到 type 为 'JOB_ENDED' 的条目
    2. 提取其中的 data.content
    3. 清洗 Markdown 标记并解析为 JSON
    """
    raw_content = None

    # ---------------- Step 1: 从繁杂的日志中定位核心内容 ----------------
    
    # 情况 A: API 直接返回了列表 (你截图中的情况)
    if isinstance(api_response, list):
        print("🔍 检测到执行日志列表，正在寻找最终结果...")
        for event in api_response:
            # 找到任务结束的标志
            if isinstance(event, dict) and event.get("type") == "JOB_ENDED":
                data = event.get("data", {})
                # 内容通常在 content 或 output 字段
                raw_content = data.get("content") or data.get("output")
                if raw_content:
                    print(" 成功定位到 JOB_ENDED 数据！")
                    break
    
    # 情况 B: API 返回的是一个包含 result 的大字典
    elif isinstance(api_response, dict):
        if "output" in api_response:
             raw_content = api_response["output"]
             # 如果 output 里面还有一层 text
             if isinstance(raw_content, dict):
                 raw_content = raw_content.get("text", "") or raw_content.get("content", "")
        elif "result" in api_response:
             raw_content = api_response["result"]

    if not raw_content:
        print("警告：在返回的数据中没找到 'JOB_ENDED' 或有效的内容字段。")
        # 调试用：只打印数据的 Keys，不打印内容，避免刷屏
        if isinstance(api_response, dict): print(f"Keys: {api_response.keys()}")
        return None

    # ---------------- Step 2: 字符串清洗 (去除 Markdown) ----------------
    
    # 确保是字符串
    if not isinstance(raw_content, str):
        # 如果已经是字典，直接返回
        if isinstance(raw_content, (dict, list)):
            return raw_content
        raw_content = str(raw_content)

    # 去掉 ```json 和 ``` 标记
    clean_text = re.sub(r'```json\s*', '', raw_content)
    clean_text = re.sub(r'```\s*', '', clean_text)
    
    # 有时候开头会有 "Answer:" 或类似前缀，尝试找到第一个 {
    start_index = clean_text.find('{')
    end_index = clean_text.rfind('}')
    if start_index != -1 and end_index != -1:
        clean_text = clean_text[start_index : end_index + 1]

    # ---------------- Step 3: JSON 解析 ----------------
    
    try:
        # 解析字符串为 Python 字典
        final_json = json.loads(clean_text)
        return final_json
    except json.JSONDecodeError as e:
        print(f"解析 JSON 失败: {e}")
        print("原始文本片段:", clean_text[:200]) # 只打印前200字调试
        return None
# ================= 3. 主程序 =================

if __name__ == "__main__":
    # 1. 从文件夹读取所有文件
    print("--- 步骤 1: 读取本地文件 ---")
    uploaded_files = load_json_files_from_folder(FOLDER_PATH)
    
    if len(uploaded_files) == 0:
        print("没有读取到文件，程序结束。请检查 FOLDER_PATH 路径设置。")
        exit()

    # 2. 准备上下文文本
    context = prepare_context_from_files(uploaded_files)
    print(f"\n已合并文档内容，总字符数: {len(context)}")

    # 3. 获取 Token
    print("\n--- 步骤 2: 获取 API Token ---")
    token = get_access_token_b(CLIENT_ID, CLIENT_SECRET)
    
    # 4. 运行 AI 任务
    if token:
        print("\n--- 步骤 3: 调用 AI 生成报告 ---")
        job_id = run_ai_job(token, context,API_BASE_URL)
        
        if job_id:
            result = poll_result(token, job_id,API_BASE_URL)
            if result:
                # 提取 AI 的回答内容
                final_report = extract_final_json(result)
                if final_report:
                    print("\n\n====== 完美清洗后的 JSON 报告 ======\n")
                # indent=4 让它漂亮地格式化打印
                    print(json.dumps(final_report, ensure_ascii=False, indent=4))
                    
                    # 保存
                    with open("final_investment_report.json", "w", encoding="utf-8") as f:
                        json.dump(final_report, f, ensure_ascii=False, indent=4)
                    print("\n报告已保存为 'final_investment_report.json'")
                else:
                    print("提取数据失败，请检查日志。")
    
