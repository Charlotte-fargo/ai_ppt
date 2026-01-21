# config.py
from pptx.dml.color import RGBColor
import requests
import os
# ==============================================================================
# Web 界面访问密码
APP_PASSWORD = "123456"
# 0. News Platform 配置 (用于抓取文章)
# ==============================================================================
NEWS_AUTH_URL = "https://auth.easyview.xyz/realms/Easyview-News-Platform-Realm/protocol/openid-connect/token"
NEWS_ARTICLE_URL = "https://news-platform.easyview.xyz/api/v1/channel/cio/articles"
NEWS_CLIENT_ID = "cio-backend"
NEWS_CLIENT_SECRET = "4cbb1527-bcc4-42ae-a7ec-691359f3e119"
# 1. AI API 与 认证配置
# ==============================================================================

AUTH_URL = "https://auth-v2.easyview.xyz/realms/evhk/protocol/openid-connect/token"
API_BASE_URL = "https://api-v2.easyview.xyz/v3/ai"
# AI 服务的专用凭据
CLIENT_ID = "cioinsight-api-client"
CLIENT_SECRET = "b02fe9e7-36e6-4c81-a389-9399184eda9b"
# AI 模型名称
AI_MODEL_NAME = "gemini-3-pro-preview"

# 请求元数据 (Metadata)
API_METADATA = {
    "tenantId": "GOLDHORSE",
    "clientId": "CIO",
    "userId": "script_runner",
    "priority": 1,
    "custom": {}
}
# 获取访问令牌的函数
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
API_TOKEN = get_access_token_b(CLIENT_ID, CLIENT_SECRET)
# ==============================================================================
# 2. AI 提示词 (Prompt) 配置 - 决定报告质量的核心
# ==============================================================================

# 系统指令 (System Prompt)：定义 AI 的身份、任务目标和输出格式
AI_SYSTEM_PROMPT_cn = """
你是一个专业的中文首席投资官助理。你需要阅读提供的金融市场分析文档，并生成一份标准化的投资观点报告。

任务要求：
    1. 生成7种资产的投资观点,资产类别包括中港股市、美股、欧股、日股、债市、黄金、原油这7类，请不要改一个字。如果提供的文档中缺少某种资产，请根据你的知识库合理推断或标记为"暂无数据"。投资逻辑中文字数必须在70字。以下生成的每一个bullet point字数必须在83字左右，三个bullet point总共的字数必须在230字以内,不要解释思考过程，直接给出最终结论。

硬性写作要求：
- 标题格式为“资产类别名称：xxxxx”
- 观点内容不超过三句 bullet point。
- 每一句观点的格式为“小标题：xxxx”。
- 标题需要抓住核心结论，点明关键驱动因素，句子里的因果逻辑之间要有逗号，不一定一句全无停顿。
- 观点内容需有理有据，避免空洞表述。
- 标题字数控制在12-15字以内。


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
注意：不要进行通过逐步推理,不要解释思考过程，直接给出最终结论。
确保输出的 JSON 结构严格符合要求，避免任何格式错误。
每生成一个bullet point后，请检查字数是否符合要求。
检查content_slides中的每个title的开头需要是资产类别中的，一个字都不能改。
"""

        # === 英文 Prompt ===
AI_SYSTEM_PROMPT_en = """
        You are an English professional assistant to a Chief Investment Officer. Read the provided financial market analysis documents and generate a standardized investment outlook report. You do not need to show your analysis process, just output the final JSON result.

        Task Requirements:
        1. Generate investment views for 7 asset classes: HK/China Equities, US Equities, European Equities, Japan Equities, Fixed Income, Gold, and Crude Oil. Do not change any Asset classes, keep. If a specific asset class is missing in the documents, infer reasonably from your knowledge base or mark it as "No Data Available". Strictly follow the output format below.Asset Title: Must be maximum 6 words. Format: "Asset Class: [Core View Summary]".For Asset:HK/China Equities,must be 5 words including HK/China Equities. Investment Rationale (Summary Logic): Must be approximately 22 words. This should be a high-level concise summary.AND Each bullet point must be 26 words not explain the thought process; provide the final conclusion directly. 

        Writing Requirements:
        - Title format: "Asset Class Name: [Core View Summary]"
        - Content must be maximum 3 bullet points.
        - Each bullet point format: "Sub-title: [Detail]".
        - Total word count per asset: around 60 words.
        - Tone: Professional, concise, financial English.

        Finally, output ONLY pure JSON. Do not include Markdown tags (like ```json). JSON structure:
        {
          "document": { "title": "Global Investment Outlook", "author":"CIO Office", "date": " " },
          "executive_summary": { 
              "columns": ["Asset Class", "Investment Logic"], 
              "rows": [ {"Asset Class": "...", "Investment Logic": "..."} ] 
          },
          "content_slides": [ 
              { "title": "...", "bullets": ["...", "..."] } 
          ]
        }
        Note:Do not Chain of Thought , do not explain the thought process; provide the final conclusion directly. 
        Ensure the output JSON structure strictly adheres to the requirements, avoiding any formatting errors.
        After generating each bullet point, check if the character count meets the requirements.
        Check that each title in content_slides starts with one of the asset class names, without any alterations.
        """
# 指引指令 (Instruction Prompt)：放在 Parameter 中，指引 AI 去读取附件
AI_INSTRUCTION_PROMPT_cn = "请详细阅读附带的文件资源（resource），文件中包含了身份设定、具体指令以及需要分析的金融文档内容。请严格按照文件中的 JSON 格式要求输出结果。"
AI_INSTRUCTION_PROMPT_en = "Please carefully read the attached file resources (resource), which contain identity settings, specific instructions, and the content of financial documents to be analyzed. Please strictly follow the JSON format requirements in the file to output the results."
# ==============================================================================
# 3. 文件路径与目录配置
# ==============================================================================
# 输出 PPT 目录
BASE_DIR = "/tmp" 
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
TEMP_JSON = os.path.join(BASE_DIR, "articles.json")
CLEANED_DIR = os.path.join(BASE_DIR, "cleaned_files")
IMAGES_DIR = os.path.join(BASE_DIR, "images")
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(CLEANED_DIR, exist_ok=True)
os.makedirs(IMAGES_DIR, exist_ok=True)
os.makedirs(BASE_DIR, exist_ok=True)


# PPT 模板路径映射 (根据用户选择的地点，自动匹配模板文件)
TEMPLATE_MAP = {
    "香港/Hong Kong": {
        "cn": "template/AI PPT v2.pptx",  # 中文版
        "en": "template/AI PPT v4.pptx"  # 英文版
    },
    "中国大陆/China": {
        "cn": "template/AI PPT v3.pptx",  # 中文版
        "en": "template/AI PPT v5.pptx"  # 英文版
    },
    "新加坡/Singapore": {
        "cn": "template/AI PPT v2.pptx",  # 中文版
        "en": "template/AI PPT v4.pptx"  # 英文版
    }
}
LANGUAGE_MAP = {
    "中文/Chinese": "cn",
    "英文/English": "en"
}

# ==============================================================================
# 4. PPT 视觉样式配置 (颜色、字体、布局)
# ==============================================================================

# 常用颜色定义
COLOR_DARK_BLUE = RGBColor(0, 32, 96)
COLOR_BLACK = RGBColor(0, 0, 0)
COLOR_GRAY = RGBColor(100, 100, 100)
COLOR_LIGHT_BLUE = RGBColor(60, 109, 148)
# 幻灯片布局索引 (Slide Layout Index)
# 注意：这些索引对应你 PPT 母版中的位置，如果母版改了，这里要改
LAYOUT_IDX = {
    'contact': 2,       # 联系我们页
    'disclaimer_en': 3, # 英文免责声明
    'disclaimer_cn': 4, # 中文免责声明
    'content': 9,       # 正文页
    'image_only': 10    # 纯图页 (如资金流)
}

# 图片注释坐标 (单位: EMU)
# 用于在图片周围添加"标题"和"资料来源"
ANNOTATION_CONFIG = {
    'title': {
        'top': 3016459,
        'left_base': 3750684,
        'width': 1616075,
        'height': 226581,
        'font_name': '华文细黑',
        'size': 11
    },
    'source': {
        'top': 6316663,
        'left_base': 7315200, 
        'width': 1285875,
        'height': 266700,
        'font_name': 'Microsoft YaHei',
        'size': 9
    },
    'contact_info': {
        'left': 381600,
        'top': 1065600,
        'width': 5090200,
        'height': 882000,
        'font_name': 'Microsoft YaHei',
        'size': 12
    }
}

# ==============================================================================
# 5. 内容常量 (地址、免责声明)
# ==============================================================================

# 联系地址映射 (Placeholder Index -> 地址文本)
CONTACT_ADDRESSES = {
    10: "香港\n香港中环康乐广场8号\n交易广场二期26楼2606-2607室\n电话：+852 2956 9700",
    17: "新加坡\n新加坡滨海大道10号\n滨海湾金融中心2座 #16-05\n电话：+65 6509 0110",
    19: "Fargo Space\n香港尖沙咀海港城\n港威大厦5座33楼3301-3304室\n电话：+852 2439 9745",
    22: "北京\n北京市朝阳区东三环中路1号\n环球金融中心西楼1013室\n电话：+86 10 6507 8234",
    18: "上海\n上海市黄浦区太仓路233号\n新茂大厦2204-2205室\n电话：+86 21 6333 8131",
    20: "深圳\n深圳市前海深港合作区兴海大道\n3040号前海世茂大厦2402室\n电话：+86 755 2691 3468",
    21: "杭州\n杭州市上城区新业路228号\n来福士中心T2办公楼1702-1703室\n电话：+86 571 8805 8596"
}
CONTACT_ADDRESSES_en = {
    10: "Hong Kong\n2606-2607, 26/F, Two Exchange Square, 8 Connaught Place, Central, Hong Kong\nTel: +852 2956 9700",
    17: "Singapore\n10 Marina Boulevard #16-05, Marina Bay Financial Centre Tower 2, Singapore 018983\nTel: +65 6509 0110",
    19: "Fargo Space\n3301-3304, 33/F, Tower 5, The Gateway, Harbour City, Tsim Sha Tsui, Hong Kong\nTel: +852 2439 9745",
    22: "Beijing\n1013, 10/F, West Tower, World Financial Center, No.1 East 3rd Ring Middle Road, Chaoyang District, Beijing\nTel: +86 10 6507 8234",
    18: "Shanghai\n2204-2205, 22/F, The Platinum, No.233 Taicang Road, Huangpu District, Shanghai\nTel: +86 21 6333 8131",
    20: "Shenzhen\n2402, 24/F, Qianhai Shimao Tower, No.3040 Xinghai Avenue, Nanshan District, Shenzhen\nTel: +86 755 2691 3468",
    21: "Hangzhou\n1702-1703, 17/F, Tower 2 Raffles City, No.228 Xinye Road, Shangcheng District, Hangzhou\nTel: +86 571 8805 8596"
}

# 免责声明文案库
DISCLAIMER_TEXTS = {
    "香港": {
        "cn": [
            "绅士资本有限公司（'绅士资本'）获得香港证券及期货事务监察委员会（'证监会'）颁发的第 9 类受监管活动（资产管理）（中央编号：BIJ793）。",
            "在代表您行使投资酌情权时，绅士资本可能会不时从经纪商处收取现金回扣，以代表您将交易业务交给经纪商。",
            "绅士资本还可能从产品发行人那里获得非金钱利益，例如研究报告、市场分析数据、投资组合分析、培训和研讨会。",
            "绅士资本认为本文件的内容基于被认为可靠的信息来源。 投资和任何产生收入的工具的价值可升可跌。 投资涉及风险，包括本金损失。 不保证股息。 我们的预测是基于过去的表现。 过去的表现并不能保证未来的结果。 投资组合可能会遭受损失并获得收益。 未来回报无法保证，可能会出现本金损失。",
            "绅士资本相信通讯中的信息和内容的来源是可靠的，但它不能也不做任何明示或暗示的保证，并且不对任何信息或数据的准确性、有效性、及时性、适销性或完整性承担任何责任。 任何特定目的或用途，或信息或数据将没有错误。绅士资本不对任何人对本文所表达的任何陈述或意见的任何依赖承担任何责任。绅士资本或其任何关联公司、董事、高级职员或雇员均不对任何人因使用此信息而可能遭受的任何性质的任何损失或损害承担任何责任或任何形式的责任。 不得出于任何目的依赖本通讯中包含的信息或意见。",
            "本通讯中的内容仅适用于继续符合合格和/或认可/专业投资者定义的绅士资本的实际客户。",
            "本通讯的内容是严格保密的，仅供背景和信息之用。绅士资本未经审核内容的真实性、事实准确性或完整性。 内容并不声称是完全或完整的。",
            "未经绅士资本事先书面许可，不得以任何方式抄襲、复制、披露或出版本通讯的任何部分。",
            "本通讯不构成任何发行或出售的要约，或任何订阅或购买任何股份或任何其他利益的要约的招揽，也不构成任何司法管辖区或任何人的任何要约，也不构成任何要约的一部分，也不构成其发行或出售的任何要约的一部分，也不构成任何合同或承诺的基础或与之相关的依赖。"
        ],
        "en": [
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
    },
    "中国大陆": {
        "cn": [
            "华港财富集团有限公司（以下统称“华港财富”），及其下属品牌华港财富不提供法务或税务咨询，本演示文本并不构成此类咨询。出版本演示文本的唯一目的是介绍相关信息，与任何读者的具体投资目标、财务状况或特定需求均无关。华港财富极力建议所有对本演示文本中阐述的产品或服务加以考虑的人士，接受恰当而独立的法务、税务及其它专业咨询。",
            "关于本文信息的精确性、完整性或可靠性，并不提供任何外在或内在的声明或保证，也无意对本资料中的课题进行完整的陈述或总结。虽然本演示文本中阐述的所有信息和意见均来自相信是非常可靠的来源，并且以良好诚信的态度整理而成，但对其精确性或完整性并不提供外在或内在的声明或保证。华港财富及其董事、雇员或代理人，不承担任何由于使用本材料中的全部或部分内容而造成任何损失或伤害的责任。",
            "某些产品和服务由于受到法律限制而无法在世界各地无限制地提供。本演示文本不构成销售要约，也不构成购买或出售任何证券或投资工具的要约邀请，不影响任何交易，也不决定任何形式的任何法律行为。本文任何内容均不应限制或妨碍任何具体报价要约中的特定条款。在任何禁止要约、要约邀请或销售的司法管辖权下，不提供任何关于任一产品的要约；也不向任何人提供这种要约、要约邀请或销售，如果这类做法非法的话。",
            "本材料中表达的任何意见均可能发生变动，不再另行通知。由于使用的假设和标准不同，这些意见可能与华港财富其它业务领域或部门表述的意见发生分歧或矛盾。华港财富没有任何对本文所包含的信息进行更新或保持其时效性的义务。",
            "本演示文本无意也不能用作以下用途：（1）规避《美国国内收入法典》（即美国税法）的惩罚，或者（2）向其他人宣传、营销或推荐任何与税务有关的事宜。",
            "除非华港财富许可，否则不得复制或分发本演示文本。",
            "本演示文本只供专业投资者参考使用，且仅限在香港地区分发和使用。",
            "©华港财富2025. 版权所有。繁体字标识和拼音注释均属于华港财富及其相关联公司注册与未注册的商标。保留所有权利。"
        ],
        "en": []
    },
    "default": {
        "cn": [
            "本手册的资料，仅供一般资讯用途。对该等资料，华港财富不会就任何错误、遗漏、或错误陈述或失实陈述(不论明示或默示)承担任何责任。对任何因使用或不当使用或依据本手册所记载的资料而引致或所涉及的损失，华港财富不承担任何义务、责任或法律责任。",
            "您有权对本手册上所有文字、照片、图片、标识、内容和其他信息（以下简称“信息”）进行保存、分析、修改、复制以供您个人使用，但未经华港财富书面同意，您不得以任何形式向任何第三方公开发表、传输本手册信息。本手册上显示的所有商标、标识及相关知识产权均为华港财富或其各自所有权所属人所有，未在本手册中明确授予他人的权利均由本公司或向本公司提供信息的第三方保留。华港财富保留不时自主编辑、修改、增加或删除本手册信息的权利。",
            "请注意本手册信息中有很大部分包括或含有第三方提供的信息，本手册并未对该第三方提供的信息作独立核实或确认。基于华港财富的业务活动，本公司经常对合作方以及其他第三方企业的信息存在保密义务，本公司的任何观点及信息披露均受到该保密义务的限制。本公司并不保证本手册观点及信息的真实性、准确性、完整性、时效性或者不存在侵权。",
            "本手册提供的各项信息并非任何形式的证券或其他资产和服务的买卖、发售或购买劝诱，或向您提供投资建议或任何具体建议。您不得将本手册信息作为业务、财务、投资、交易、法律、监管、税收或会计建议，或作为您本人、他人代表您本人、您的会计师或由您管理或受托的帐户进行任何投资决策的主要依据，您应当就任何计划进行的交易征询您的业务顾问、律师、税务及会计顾问的意见，华港财富不会对您使用或依赖本手册信息产生的任何后果负责。",
            "©华港财富2025. 版权所有。繁体字标识和拼音注释均属于华港财富及其相关联公司注册与未注册的商标。保留所有权利"
        ],
        "en": [
            "The information in this manual is for general information purposes only. Fargo Wealth shall not be held liable for any errors, omissions, or inaccuracies in such information, whether expressed or implied. Fargo Wealth disclaims any obligation, responsibility, or legal liability for any loss or damage arising from the use or misuse of, or reliance upon, the information contained in this manual.",
            "You are authorized to save, analyze, modify, and copy all the texts, photos, images, logos, content, and other information (hereinafter referred to as \"Information\") in this manual for your personal use. However, without the prior written consent of Fargo Wealth, you are not allowed to publicly disclose or transmit the information in this manual in any form to any third party. All trademarks, logos, and related intellectual property rights displayed in this manual are owned by Fargo Wealth or their respective owners, unless otherwise explicitly granted in this manual. Fargo Wealth reserves the right to independently edit, modify, add, or remove the information in this manual from time to time.",
            "Please note that a significant portion of the information in this manual includes or contains information provided by third parties, and this manual has not independently verified or confirmed the information provided by those third parties. Due to the business activities of Fargo Wealth, the company often has confidentiality obligations with respect to information from partners and other third-party entities. Any views and information disclosed by the company are subject to these confidentiality obligations. The company does not guarantee the truthfulness, accuracy, completeness, timeliness, or absence of infringement of the views and information in this manual.",
            "The information provided in this manual does not constitute any form of solicitation, offer, or inducement to buy, sell, or purchase securities or other assets and services, nor does it provide investment advice or any specific recommendations. You should not consider the information in this manual as business, financial, investment, trading, legal, regulatory, tax, or accounting advice, or as the primary basis for making any investment decisions for yourself, on behalf of others, by your accountant, or for accounts managed or entrusted to you. You should consult your business advisor, lawyer, tax advisor, and accountant for their opinions on any transactions you plan to undertake. Fargo Wealth will not be responsible for any consequences arising from your use of or reliance on the information in this manual.",
            "© Fargo Wealth 2025. All rights reserved. The traditional Chinese characters and Pinyin annotations are trademarks owned by Fargo Wealth and its affiliated companies, whether registered or unregistered. All rights reserved."
        ]
    }
}
