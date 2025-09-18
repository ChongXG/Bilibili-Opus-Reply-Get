import requests
import json
import time
import random
import os
from datetime import datetime
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import List, Dict, Optional

class BilibiliCommentExporter:
    def __init__(self):
        self.session = requests.Session()
        # 更完整的请求头，模拟真实浏览器
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Referer': 'https://www.bilibili.com/',
            'Origin': 'https://www.bilibili.com',
            'Accept': 'application/json, text/plain, */*',
            'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8',
            'Accept-Encoding': 'gzip, deflate, br',
            'Connection': 'keep-alive',
            'Sec-Fetch-Dest': 'empty',
            'Sec-Fetch-Mode': 'cors',
            'Sec-Fetch-Site': 'same-site',
            'DNT': '1',
            'Sec-GPC': '1'
        })
    
    def get_dynamic_comments(self, dynamic_id: str, max_pages: int = 10) -> List[Dict]:
        comments = []
        next_page = 1  # 从第一页开始
        page = 1
        
        while page <= max_pages:
            print(f"正在获取第 {page} 页评论...")
            
            # 构建API URL - 注意这里type=11表示动态
            api_url = f"https://api.bilibili.com/x/v2/reply/main?jsonp=jsonp&next={next_page}&type=11&oid={dynamic_id}&mode=3&plat=1"
            
            try:
                # 添加随机延迟，避免请求过于频繁
                time.sleep(random.uniform(1.0, 2.5))
                
                response = self.session.get(api_url, timeout=15)
                
                # 检查HTTP状态码
                if response.status_code != 200:
                    print(f"HTTP错误: {response.status_code}")
                    print(f"响应内容: {response.text[:200]}...")
                    break
                
                # 检查响应内容是否为JSON
                if not response.text.strip().startswith('{'):
                    print(f"响应不是JSON格式: {response.text[:100]}...")
                    break
                
                # 尝试解析JSON
                try:
                    data = response.json()
                except json.JSONDecodeError as e:
                    print(f"JSON解析错误: {e}")
                    print(f"响应内容: {response.text[:200]}...")
                    break
                
                if data['code'] != 0:
                    print(f"API返回错误: {data['message']} (代码: {data['code']})")
                    break
                
                # 检查数据是否存在
                if 'data' not in data:
                    print("API返回的数据中没有data字段")
                    break
                
                if 'replies' not in data['data']:
                    print("API返回的数据中没有replies字段")
                    print("可能是动态没有评论或API结构已变化")
                    break
                
                # 解析评论数据
                replies = data['data']['replies']
                if not replies:
                    print("没有更多评论了")
                    break
                
                for reply in replies:
                    comment = {
                        'rpid': reply.get('rpid', ''),  # 评论ID
                        'mid': reply.get('mid', ''),  # 用户ID
                        'uname': reply.get('member', {}).get('uname', ''),  # 用户名
                        'message': reply.get('content', {}).get('message', ''),  # 评论内容
                        'like': reply.get('like', 0),  # 点赞数
                        'ctime': datetime.fromtimestamp(reply.get('ctime', 0)).strftime('%Y-%m-%d %H:%M:%S'),  # 评论时间
                        'root': reply.get('root', ''),  # 根评论ID
                        'parent': reply.get('parent', '')  # 父评论ID
                    }
                    comments.append(comment)
                
                # 检查是否有下一页
                if 'cursor' not in data['data'] or data['data']['cursor'].get('is_end', True):
                    print("已到达最后一页")
                    break
                
                next_page = data['data']['cursor'].get('next', next_page + 1)
                page += 1
                
            except requests.exceptions.RequestException as e:
                print(f"网络请求错误: {e}")
                break
            except KeyError as e:
                print(f"数据键错误: {e}，可能是API返回结构变化")
                if 'data' in locals():
                    print(f"响应数据: {json.dumps(data, ensure_ascii=False, indent=2)[:500]}...")
                break
            except Exception as e:
                print(f"获取评论时出错: {e}")
                break
        
        return comments
    
    def export_to_json(self, comments: List[Dict], filepath: str):
        """导出评论到JSON文件"""
        try:
            # 确保目录存在
            os.makedirs(os.path.dirname(filepath), exist_ok=True)
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(comments, f, ensure_ascii=False, indent=2)
            print(f"评论已导出到 {filepath}")
            return True
        except Exception as e:
            print(f"导出到JSON失败: {e}")
            return False
    
    def export_to_excel(self, comments: List[Dict], filepath: str):
        """导出评论到Excel文件"""
        try:
            # 确保目录存在
            os.makedirs(os.path.dirname(filepath), exist_ok=True)
            df = pd.DataFrame(comments)
            df.to_excel(filepath, index=False)
            print(f"评论已导出到 {filepath}")
            return True
        except Exception as e:
            print(f"导出到Excel失败: {e}")
            return False
    
    def export_to_csv(self, comments: List[Dict], filepath: str):
        """导出评论到CSV文件"""
        try:
            # 确保目录存在
            os.makedirs(os.path.dirname(filepath), exist_ok=True)
            df = pd.DataFrame(comments)
            df.to_csv(filepath, index=False, encoding='utf-8-sig')
            print(f"评论已导出到 {filepath}")
            return True
        except Exception as e:
            print(f"导出到CSV失败: {e}")
            return False

def select_save_path():
    """使用弹窗选择保存路径"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    
    # 选择保存目录
    save_path = filedialog.askdirectory(
        title="选择保存目录",
        initialdir=os.getcwd()
    )
    
    # 如果用户取消了选择，使用当前目录
    if not save_path:
        save_path = os.getcwd()
        print("未选择目录，使用当前目录")
    
    return save_path

def select_save_file(export_format):
    """使用弹窗选择保存文件"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口
    
    # 根据格式设置文件类型
    if export_format == "excel":
        filetypes = [("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        default_extension = ".xlsx"
    elif export_format == "csv":
        filetypes = [("CSV文件", "*.csv"), ("所有文件", "*.*")]
        default_extension = ".csv"
    else:  # json
        filetypes = [("JSON文件", "*.json"), ("所有文件", "*.*")]
        default_extension = ".json"
    
    # 生成默认文件名
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    default_filename = f"bilibili_comments_{timestamp}{default_extension}"
    
    # 选择保存文件
    filepath = filedialog.asksaveasfilename(
        title="保存评论文件",
        defaultextension=default_extension,
        initialfile=default_filename,
        filetypes=filetypes
    )
    
    return filepath

def main():
    exporter = BilibiliCommentExporter()
    
    # 从用户输入获取动态ID
    dynamic_id = input("请输入B站动态OID（从开发者工具(F12)中获取）: ").strip()
    
    if not dynamic_id:
        print("错误: 未提供动态OID")
        return
    
    # 获取用户指定的页数
    try:
        max_pages_input = input("请输入要获取的页数（默认5页，最多999页）: ").strip()
        max_pages = int(max_pages_input) if max_pages_input else 5
        max_pages = min(max_pages, 999)  # 限制最大页数为999
    except ValueError:
        max_pages = 5
        print("输入无效，使用默认值5页")
    
    # 获取评论
    print("开始获取评论，请稍候...")
    comments = exporter.get_dynamic_comments(dynamic_id, max_pages)
    
    if not comments:
        print("未获取到任何评论")
        return
    
    print(f"共获取到 {len(comments)} 条评论")
    
    # 导出选项
    export_option = input("请选择导出格式 (1: JSON, 2: Excel, 3: CSV, 按Enter默认JSON): ").strip()
    
    # 使用弹窗选择保存路径
    print("请选择保存位置...")
    if export_option == "2":
        filepath = select_save_file("excel")
        if not filepath:
            print("未选择文件，取消导出")
            return
        success = exporter.export_to_excel(comments, filepath)
    elif export_option == "3":
        filepath = select_save_file("csv")
        if not filepath:
            print("未选择文件，取消导出")
            return
        success = exporter.export_to_csv(comments, filepath)
    else:
        filepath = select_save_file("json")
        if not filepath:
            print("未选择文件，取消导出")
            return
        success = exporter.export_to_json(comments, filepath)
    
    # 显示导出结果
    if success:
        print(f"评论已成功导出到: {filepath}")
        
        # 显示前几条评论
        print("\n前5条评论预览:")
        for i, comment in enumerate(comments[:5]):
            print(f"{i+1}. [{comment['ctime']}] {comment['uname']}: {comment['message'][:50]}...")
    else:
        print("导出失败")

if __name__ == "__main__":
    main()