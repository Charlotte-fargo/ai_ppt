import os
import glob
import re
import json
import base64
import requests
import time
import logging
import config  # 引入配置文件

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class AIPromptRunner:
    def __init__(self, language=None):
        # 从配置中读取参数
        self.api_base = config.API_BASE_URL
        # self.token = config.API_TOKEN
        # self.input_dir = config.INPUT_BASE_DIR
        self.AUTH_URL = config.AUTH_URL
        self.CLIENT_ID = config.CLIENT_ID
        self.CLIENT_SECRET = config.CLIENT_SECRET
        self.model_name = config.AI_MODEL_NAME
        self.metadata = config.API_METADATA
        if language == "en":
            self.prompt = config.AI_INSTRUCTION_PROMPT_en
            self.AI_system_prompt = config.AI_SYSTEM_PROMPT_en
        else:
            self.prompt = config.AI_INSTRUCTION_PROMPT_cn
            self.AI_system_prompt = config.AI_SYSTEM_PROMPT_cn

        
        # 运行时状态
        self.context_text = ""

    # ================= 1. 数据准备 =================
   
    def load_files(self, folder_path=None):
        """读取文件夹下的所有 JSON 文件并合并"""
        target_dir = folder_path if folder_path else self.input_dir
        
        if not os.path.exists(target_dir):
            logging.error(f"文件夹不存在: {target_dir}")
            return False

        files = glob.glob(os.path.join(target_dir, "*.json"))
        logging.info(f"在 '{target_dir}' 下找到 {len(files)} 个文件")

        if not files:
            return False

        merged_content = "以下是各资产类别的原始分析报告内容：\n\n"
        
        for file_path in files:
            try:
                filename = os.path.basename(file_path)
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                # 尝试提取标题和正文 (兼容不同爬虫格式)
                # 假设格式：{'titles': {'zh_CN': '...'}, 'contents': {'zh_CN': '...'}}
                titles = data.get("titles", {})
                title = titles.get("zh_CN", "未知标题") if isinstance(titles, dict) else "未知标题"
                
                contents = data.get("contents", {})
                raw_html = contents.get("zh_CN", "") if isinstance(contents, dict) else ""
                
                # 简单清洗 HTML
                clean_text = re.sub(r'<.*?>', '', raw_html).strip()
                
                merged_content += f"--- 文档开始: {filename} (标题: {title}) ---\n"
                merged_content += clean_text + "\n"
                merged_content += f"--- 文档结束 ---\n\n"
                
                logging.info(f"已读取: {filename}")
            except Exception as e:
                logging.warning(f"读取失败 {file_path}: {e}")

        self.context_text = merged_content
        return True

    # ================= 2. 任务提交 =================

   def _prepare_payload(self):
        """构建 LLM 调用 Payload（不使用附件）"""
        if not self.context_text:
            return None

        # 1. 拼接 System Prompt 和 Context
        full_prompt = f"{self.AI_system_prompt}\n\n{self.context_text}"
        
        
        # 2. 组装 Payload
        payload = {
            "type": "callLlm",
            "metadata": self.metadata,
            "input": {
                "parameter": {
                    "model_name": self.model_name,
                    "prompt": full_prompt  # 直接使用拼接后的完整 prompt
                }
                # 移除了 resource 部分，因为我们不需要附件
            },
            "callback": []
        }
        return payload

    def submit_job(self):
        """提交 AI 任务"""
        url = f"{self.api_base}/job"
        headers = {
            'Authorization': f'Bearer {self.token}',
            'Content-Type': 'application/json'
        }
        
        payload = self._prepare_payload()
        if not payload:
            return None

        logging.info(f"正在提交任务... (文本长度: {len(self.context_text)})")
        
        try:
            resp = requests.post(url, headers=headers, json=payload)
            if resp.status_code != 200:
                logging.error(f"提交失败: {resp.status_code} - {resp.text}")
                return None
            
            data = resp.json()
            # 优先获取 ID (Int)，其次 UUID
            job_id = data.get("id") or data.get("uuid")
            logging.info(f"任务提交成功，ID: {job_id}")
            return job_id
        except Exception as e:
            logging.error(f"请求异常: {e}")
            return None

    # ================= 3. 结果轮询与清洗 =================

    def poll_job(self, job_id, max_retries=60):
        """轮询任务状态"""
        url = f"{self.api_base}/job/JOB_ID/{job_id}"
        headers = {'Authorization': f'Bearer {self.token}'}
        
        logging.info(f"开始轮询结果: {url}")
        
        for i in range(max_retries):
            try:
                resp = requests.get(url, headers=headers)
                if resp.status_code == 200:
                    data = resp.json()
                    status = data.get("status")
                    logging.info(f"[第 {i+1} 次] 状态: {status}")
                    
                    if status in ["SUCCESS", "COMPLETED"]:
                        return data
                    if status in ["FAILED", "ERROR"]:
                        logging.error(f"任务失败: {data}")
                        return None
                else:
                    logging.warning(f"查询报错: {resp.status_code}")
            except Exception as e:
                logging.warning(f"轮询异常: {e}")
            
            time.sleep(3)
        
        logging.error("等待超时")
        return None

    def _extract_json_content(self, api_response):
        """从 API 响应中提取并清洗 JSON"""
        raw_content = None

        # 1. 寻找内容字段 (兼容 List 和 Dict 返回)
        if isinstance(api_response, list):
            for event in api_response:
                if event.get("type") == "JOB_ENDED":
                    data = event.get("data", {})
                    raw_content = data.get("content") or data.get("output")
                    break
        elif isinstance(api_response, dict):
            # 兼容不同的 output 结构
            out = api_response.get("output") or api_response.get("result")
            if isinstance(out, dict):
                raw_content = out.get("text") or out.get("content")
            else:
                raw_content = out

        if not raw_content:
            logging.error("未找到有效的输出内容")
            return None

        # 2. 清洗 Markdown
        text = str(raw_content)
        text = re.sub(r'```json\s*', '', text)
        text = re.sub(r'```\s*', '', text)
        
        # 3. 截取 JSON 部分
        start = text.find('{')
        end = text.rfind('}')
        if start != -1 and end != -1:
            text = text[start : end + 1]

        # 4. 解析
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            logging.error("JSON 解析失败，返回原始内容")
            logging.debug(text[:500])
            return None

    def save_report(self, json_data, output_file="final_investment_report.json"):
        """保存最终结果"""
        try:
            with open(output_file, "w", encoding="utf-8") as f:
                json.dump(json_data, f, ensure_ascii=False, indent=4)
            logging.info(f"报告已保存: {output_file}")
            return True
        except Exception as e:
            logging.error(f"保存失败: {e}")
            return False

    def get_access_token_b(self):
        payload = {
            'grant_type': 'client_credentials',
            'client_id': self.CLIENT_ID,
            'client_secret': self.CLIENT_SECRET
        }
        try:
            resp = requests.post(self.AUTH_URL, data=payload)
            resp.raise_for_status()
            return resp.json().get('access_token')
        except Exception as e:
            print(f" 认证失败: {e}")
            return None
    # ================= 4. 主流程入口 =================

    def run(self, specific_folder=None):
        """执行全流程"""
        # 1. 读取文件
        if not self.load_files(specific_folder):
            logging.error("文件加载失败，流程终止")
            return None
        token = self.get_access_token_b()
        self.token = token
        # 2. 提交任务
        job_id = self.submit_job()
        if not job_id:
            return None
        
        # 3. 轮询结果
        result_raw = self.poll_job(job_id)
        if not result_raw:
            return None
        
        # 4. 提取清洗
        final_json = self._extract_json_content(result_raw)
        
        # 5. 保存
        if final_json:
            self.save_report(final_json)
            return final_json
        return None

# ================= 测试入口 =================

if __name__ == "__main__":
    # 如果想测试特定的文件夹，可以在这里指定
    TEST_FOLDER = "input_articles/20260116/articles_20260116/" 
    # TEST_FOLDER = None # 使用 config.py 里的默认配置

    runner = AIPromptRunner()
    runner.run(TEST_FOLDER)
