import json
import os
import re
from bs4 import BeautifulSoup
import requests
import shutil
from datetime import datetime
from urllib.parse import urlparse
import time
import sys
print(f"Python 版本: {sys.version}")
print(f"Python 路径: {sys.executable}")
"""
This script processes JSON articles to filter and download only the latest dated articles along with their first images.

It reads articles from 'test.json', filters to keep only those with the most recent publishTime date,
saves them in dated subfolders under 'input_articles', and downloads the first image from each article's content.

本脚本处理JSON文章，筛选并下载最新日期的文章及其第一张图片。
从'test.json'读取文章，筛选出具有最新publishTime日期的文章，
将它们保存在'input_articles'下的日期子文件夹中，并下载每篇文章内容中的第一张图片。

Functions:
- load_articles(file_path): Loads and returns articles from the specified JSON file.
- filter_latest_articles(articles): Filters the list to include only articles from the latest publish date.
- extract_first_image_url(html_content): Extracts the first image URL from HTML content using regex.
- download_image(img_url, save_path): Downloads an image from the URL and saves it to the specified path.
- get_file_extension(url): Determines the file extension from the image URL.
- process_article(article, idx, output_dir, articles_dir, images_dir): Processes a single article by saving its JSON and downloading its image.
- main(): Orchestrates the entire process: loads data, filters, sets up directories, and processes articles.

函数：
- load_articles(file_path): 从指定的JSON文件加载并返回文章。
- filter_latest_articles(articles): 筛选列表以仅包含最新发布日期的文章。
- extract_first_image_url(html_content): 使用正则表达式从HTML内容中提取第一个图片URL。
- download_image(img_url, save_path): 从URL下载图片并保存到指定路径。
- get_file_extension(url): 从图片URL确定文件扩展名。
- process_article(article, idx, output_dir, articles_dir, images_dir): 处理单个文章：保存其JSON并下载图片。
- main(): 编排整个过程：加载数据、筛选、设置目录并处理文章。
"""
def load_articles(file_path):
    """读取原始 JSON 文件并返回文章列表"""
    with open(file_path, "r", encoding="utf-8") as f:
        data = json.load(f)
    return data.get("articles", [])

def filter_latest_articles(articles):
    """过滤出最新的日期的文章"""
    valid_articles = []
    for article in articles:
        publish_time = article.get("metadata", {}).get("audit", {}).get("publishTime", "")
        if publish_time:
            try:
                dt = datetime.strptime(publish_time, "%Y-%m-%dT%H:%M:%SZ")
                valid_articles.append((dt, article))
            except (ValueError, TypeError):
                pass

    if valid_articles:
        latest_dt = max(dt for dt, _ in valid_articles)
        latest_articles = [article for dt, article in valid_articles if dt.date() == latest_dt.date()]
    else:
        latest_articles = []
    return latest_articles


def extract_first_image_url(html_content):
    """从HTML内容中提取第一张图片的URL"""
    try:
        from bs4 import BeautifulSoup
        soup = BeautifulSoup(html_content, 'html.parser')
        
        # 查找第一个img标签
        img_tag = soup.find('img')
        if img_tag and img_tag.get('src'):
            return img_tag.get('src')
        
        # 如果没有找到img标签，尝试查找其他可能的图片标签
        # 例如，有些文章可能使用div的背景图片
        div_with_bg = soup.find(style=re.compile(r'background.*?url'))
        if div_with_bg:
            # 提取url
            import re
            match = re.search(r'url\(["\']?(.*?)["\']?\)', div_with_bg.get('style', ''))
            if match:
                return match.group(1)
        
        return None
    except Exception as e:
        print(f"提取图片URL时出错: {e}")
        return None
def download_image(img_url, save_path):
    """下载图片并保存到指定路径"""
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        
        response = requests.get(img_url, headers=headers, timeout=10)
        response.raise_for_status()
        
        with open(save_path, 'wb') as f:
            f.write(response.content)
        
        print(f"    图片下载成功: {save_path}")
        return True
    except Exception as e:
        print(f"    图片下载失败: {e}")
        return False

def get_file_extension(url):
    """从URL获取文件扩展名"""
    path = urlparse(url).path
    # 获取扩展名
    ext = os.path.splitext(path)[1]
    # 如果没有扩展名或扩展名太长，使用默认值
    if not ext or len(ext) > 10:
        return '.jpg'
    return ext

def extract_first_data_source(html_content):
    """提取第一个出现的资料来源"""
    try:
        soup = BeautifulSoup(html_content, 'html.parser')
        for p in soup.find_all('p'):
            text = p.get_text() # 不先 strip，保留内部结构
            if '资料来源' in text:
                # 核心修复：使用 split 切割字符串
                # parts[-1] 表示取“资料来源”后面的那一部分
                parts = text.split('资料来源')
                
                if len(parts) > 1:
                    # 取最后一部分（即冒号和来源名）
                    target_part = parts[-1]
                    
                    # 清理冒号（中文和英文）以及前后的空白
                    source = target_part.replace('：', '').replace(':', '').strip()
                    
                    # 二次清洗：如果来源后面紧接着还有换行符（虽然HTML里通常在末尾），只取第一行
                    # 这一步是为了防止“资料来源”出现在段落中间的情况
                    if '\n' in source:
                        source = source.split('\n')[0].strip()
                    
                    # 验证长度：真正的资料来源通常很短（例如 "Bloomberg"），如果提取出太长的文字说明切分错了
                    if source and len(source) < 50:
                        return source
        return None
    except Exception as e:
        print(f"提取出错: {e}")
        return None
def remove_unpaired_brackets(text):
    stack = []
    indices_to_remove = set()

    # 第一遍，标记未配对的括号
    for i, char in enumerate(text):
        if char == '('or char == '（':
            stack.append(i)
        elif char == ')'or char == '）':
            if stack:
                stack.pop()
            else:
                indices_to_remove.add(i)
    # 栈中剩余的左括号也是未配对的
    indices_to_remove.update(stack)

    # 构建新的字符串，跳过要删除的索引
    new_text = ''.join([char for i, char in enumerate(text) if i not in indices_to_remove])
    return new_text
def extract_chart_title(html_content):
    """提取图表标题 (仅提取冒号后面的内容)"""
    try:
        import re
        # 1. 前面的 '图表...[：:]' 不放在括号里，作为匹配条件但不提取。
        # 2. \s* 允许冒号后面可能有空格。
        # 3. ([^<>\n]+) 是唯一的捕获组，只提取冒号后面的文字。
        pattern = r'图表\s*[0-9一二三四五六七八九十]+\s*[：:]\s*([^<>\n]+)'
        
        match = re.search(pattern, html_content)
        if match:
            # match.group(1) 只对应上面的 ([^<>\n]+) 部分
            title = match.group(1).strip()
            
            # 再次清洗，防止包含多余的 HTML 符号或空格
            title = re.sub(r'<[^>]+>', '', title)
            title = remove_unpaired_brackets(title)
            return title.strip()
            
    except Exception as e:
        print(f"提取图表标题出错: {e}")
    return None

def process_article_by_category(articles, output_dir, articles_dir, images_dir):
    """按分类处理文章：每个分类只保存最新的文章"""
    
    # 1. 按分类分组文章
    categories_dict = {}
    
    for idx, article in enumerate(articles):
        # 获取文章分类
        cio_tags = article.get("metadata", {}).get("classifications", {}).get("tagNames", {}).get("cio", [])
        
        if cio_tags:
            # 优先取带 'cio_category_' 的标签
            categories = [tag.replace("cio_category_", "") for tag in cio_tags if "cio_category_" in tag]
            if not categories:
                categories = cio_tags
            category_name = categories[0] if categories else "未分类"
        else:
            category_name = "未分类"
        
        # 获取发布时间
        publish_time = article.get("metadata", {}).get("audit", {}).get("publishTime", "")
        
        # 按分类分组
        if category_name not in categories_dict:
            categories_dict[category_name] = []
        
        categories_dict[category_name].append({
            "article": article,
            "publish_time": publish_time,
            "original_index": idx
        })
    
    print(f"共发现 {len(categories_dict)} 个分类")
    
    # 2. 每个分类选择最新的一篇文章
    selected_articles = []
    
    for category_name, articles_list in categories_dict.items():
        # 按发布时间排序（最新的在前）
        articles_list.sort(
            key=lambda x: x["publish_time"] if x["publish_time"] else "", 
            reverse=True
        )
        
        # 取最新的文章
        if articles_list:
            selected_article = articles_list[0]
            selected_articles.append({
                "article": selected_article["article"],
                "category": category_name,
                "publish_time": selected_article["publish_time"],
                "original_index": selected_article["original_index"]
            })
            print(f"分类 '{category_name}': 选择了第 {selected_article['original_index']+1} 篇文章（最新）")
    
    print(f"总共选择了 {len(selected_articles)} 篇文章（每个分类一篇）")
    print()
    
    # 3. 处理每篇选中的文章
    for idx, item in enumerate(selected_articles):
        article = item["article"]
        category_name = item["category"]
        publish_time = item["publish_time"]
        
        # 取标题
        title = article.get("titles", {}).get("zh_CN", f"未命名文章_{idx}")
        
        print(f"处理文章 {idx+1}: {title[:50]}..." if len(title) > 50 else f"处理文章 {idx+1}: {title}")
        print(f"  分类: {category_name}")
        
        # 格式化发布时间
        if publish_time:
            try:
                dt = datetime.strptime(publish_time, "%Y-%m-%dT%H:%M:%SZ")
                formatted_time = dt.strftime("%Y%m%d")
            except (ValueError, TypeError):
                formatted_time = "无日期"
        else:
            formatted_time = "无日期"
        
        print(f"  发布时间: {formatted_time}")
        
        # 清理文件名
        safe_category = re.sub(r'[<>:"/\\|?*]', '_', category_name)
        
        # 生成文件名：类别_发布时间.json
        file_name = f"{safe_category}_{formatted_time}.json"
        file_path = os.path.join(articles_dir, file_name)
        
        # 保存完整文章
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(article, f, ensure_ascii=False, indent=2)
        
        print(f"  文章已保存: {file_name}")
        
        # 提取并下载第一张图片
        html_content = article.get("contents", {}).get("zh_CN", "")
        # print(html_content)
        if html_content:
            img_url = extract_first_image_url(html_content)
            
            if img_url:
                print(f"  发现图片: {img_url[:80]}..." if len(img_url) > 80 else f"  发现图片: {img_url}")
                
                # 1. 获取扩展名
                img_ext = get_file_extension(img_url)
                
                # 2. 提取并清洗标题
                raw_title = extract_chart_title(html_content)
                if raw_title:
                    # 清洗标题中的非法字符 (Windows文件名不支持 \ / : * ? " < > |)
                    safe_title = re.sub(r'[<>:"/\\|?*]', '_', raw_title)
                    print(f"  提取到标题: {safe_title}")
                else:
                    safe_title = "无标题" # 给一个默认值，防止 NoneType 报错
                    
                # 3. 提取资料来源
                data_source = extract_first_data_source(html_content)
                print(f"  资料来源: {data_source}" if data_source else "  未找到资料来源")

                # 4. 生成文件名逻辑
                # 情况 A: 特殊分类 - 个股投资观点更新 (强制来源 bloomberg，标题 NONE)
                if safe_category == "个股投资观点更新":
                    safe_category =="资金流"
                    safe_data_source = "Bloomberg"
                    final_title = "NONE"
                    img_file_name = f"{"资金流"}_{final_title}_{safe_data_source}{img_ext}"

                # 情况 B: 特殊分类 - 精选类 (只保留分类名)
                elif safe_category in ["个股精选", "个债精选"]:
                    # 注意：后续代码有 while os.path.exists 检测，重复时会自动变为 个股精选_1.jpg
                    img_file_name = f"{safe_category}{img_ext}"

                # 情况 C: 普通分类 (包含 标题 和 来源)
                else:
                    # 清洗来源中的非法字符
                    if data_source:
                        safe_data_source = re.sub(r'[<>:"/\\|?*]', '_', data_source)
                        img_file_name = f"{safe_category}_{safe_title}_{safe_data_source}{img_ext}"
                    else:
                        # 只有标题，没有来源
                        img_file_name = f"{safe_category}_{safe_title}{img_ext}"

                # 创建图片文件名（使用分类名而不是索引）
        
                img_file_path = os.path.join(images_dir, img_file_name)
                
                
                # 下载图片
                download_image(img_url, img_file_path)
                
                # 在JSON文件中记录图片路径
                if os.path.exists(img_file_path):
                    article["local_image_path"] = os.path.relpath(img_file_path, output_dir)
                    with open(file_path, "w", encoding="utf-8") as f:
                        json.dump(article, f, ensure_ascii=False, indent=2)
                else:
                    print("  未发现图片")
            else:
                print("  无HTML内容")
        

        
        # 避免请求过快，添加短暂延迟
        time.sleep(0.5)
    
    return selected_articles
def json_main(json_path):
    # 1. 读取原始 JSON
    articles = load_articles(json_path)
    
    # 2. 过滤出最新的日期的文章
    latest_articles = filter_latest_articles(articles)
    
    if not latest_articles:
        print("没有找到有效的文章")
        return

    # 检查并删除现有的input_articles文件夹
    if os.path.exists("input_articles"):
        shutil.rmtree("input_articles")
        print("已删除现有的input_articles文件夹")

    # 获取最新日期
    publish_time = latest_articles[0].get("metadata", {}).get("audit", {}).get("publishTime", "")
    if publish_time:
        try:
            dt = datetime.strptime(publish_time, "%Y-%m-%dT%H:%M:%SZ")
            date_str = dt.strftime("%Y%m%d")
        except (ValueError, TypeError):
            date_str = "unknown"
    else:
        date_str = "unknown"

    # 3. 输出目录，以日期命名
    output_dir = os.path.join("input_articles", date_str)
    articles_dir = os.path.join(output_dir, f"articles_{date_str}")
    images_dir = os.path.join(output_dir, f"images_{date_str}")
    os.makedirs(articles_dir, exist_ok=True)
    os.makedirs(images_dir, exist_ok=True)
    
    selected_articles = process_article_by_category(
        articles, output_dir, articles_dir, images_dir
    )
    print("处理完成！")
    print(f"共处理了 {len(selected_articles)} 个分类的文章")
    print(f"输出目录: {output_dir}")
    print(f"文章保存目录: {articles_dir}")
    print(f"图片保存目录: {images_dir}")
    # 显示分类统计
    categories = [item["category"] for item in selected_articles]
    print(f"\n处理的分类列表: {', '.join(categories)}")
    return articles_dir, images_dir

if __name__ == "__main__":
    json_main("articles.json")