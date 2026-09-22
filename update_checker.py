import requests
import json
import os
import time       # 【谨慎改造】引入时间模块用于休眠
import random     # 【谨慎改造】引入随机模块模拟人类行为
from pathlib import Path

class UpdateChecker:
    """检查国家法律法规数据库的更新"""
    
    def __init__(self, db_path=None):
        self.base_url = "https://flk.npc.gov.cn"
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
            'Content-Type': 'application/json',
            'Accept': 'application/json, text/plain, */*'
        })
        self.db_path = db_path or Path(__file__).parent / "legal_database.db"

    def get_enum_data(self):
        """获取分类枚举数据，用于确定分类编码"""
        url = f"{self.base_url}/law-search/search/enumData"
        resp = self.session.get(url, timeout=30)
        resp.raise_for_status()
        return resp.json()

    def search_laws(self, keyword="", category_code=None, page_num=1, page_size=20):
        """搜索法规，用于获取最新的法规列表"""
        url = f"{self.base_url}/law-search/search/list"
        payload = {
            "searchRange": 1,
            "searchType": 2,
            "searchContent": keyword,
            "orderByParam": {"order": "-1", "sort": "gbrq"},
            "flfgCodeId": [category_code] if category_code else [],
            "pageNum": page_num,
            "pageSize": page_size,
            "sxx": []  # 可添加状态过滤，如 [3] 表示有效
        }
        
        # 【谨慎改造】加入异常退避机制
        try:
            resp = self.session.post(url, json=payload, timeout=30)
            
            if resp.status_code == 429:
                print("🚨 触发官方限流（429），立刻休息 30 秒，本轮更新暂停！")
                time.sleep(30)
                raise Exception("触发限流，自动退出")
            elif resp.status_code == 403:
                print("🚨 访问被拒绝（403），可能被短暂封禁，立刻停止！")
                raise Exception("访问被拒绝，停止抓取")
                
            resp.raise_for_status()
            return resp.json()
            
        except requests.exceptions.RequestException as e:
            print(f"网络请求出错: {e}")
            raise

    def get_law_detail(self, bbbs_id):
        """获取法规详情，包括下载链接"""
        url = f"{self.base_url}/law-search/search/flfgDetails?bbbs={bbbs_id}"
        resp = self.session.get(url, timeout=30)
        resp.raise_for_status()
        return resp.json()

    def get_existing_bbbs_set(self):
        """从本地数据库中获取已存在的法规bbbs标识集合"""
        import sqlite3
        if not self.db_path.exists():
            return set()
        conn = sqlite3.connect(str(self.db_path))
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT bbbs FROM laws")
            existing = {row[0] for row in cursor.fetchall()}
        except sqlite3.OperationalError:
            existing = set()
        finally:
            conn.close()
        return existing

    def check_for_updates(self, category="法律", max_pages=1, status_filter=None):
        """
        检查指定分类下是否有新法规
        返回新增的法规列表
        """
        # 【谨慎改造】强制上限：不管前端传什么，最多只抓 1 页
        max_pages = min(max_pages, 1)

        # 1. 获取分类编码（加入容错处理）
        category_code = None
        try:
            enum_data = self.get_enum_data()
            for item in enum_data.get("data", {}).get("flfgCode", []):
                if item.get("name") == category or category in item.get("name", ""):
                    category_code = item.get("codeId")
                    break
        except Exception as e:
            print(f"获取分类枚举失败，将跳过分类过滤。错误信息: {e}")

        if not category_code:
            print(f"警告：未找到分类 [{category}] 的精确编码，将默认搜索全部分类。")

        # 2. 获取本地已有的法规标识
        existing_bbbs = self.get_existing_bbbs_set()

        # 3. 分页搜索，找出新法规
        new_laws = []
        for page in range(1, max_pages + 1):
            # 【谨慎改造】模拟人类阅读，随机休息 3 到 7 秒
            sleep_time = random.uniform(3.0, 7.0)
            print(f"正在检查第 {page} 页，休息 {sleep_time:.1f} 秒以减轻服务器压力...")
            time.sleep(sleep_time)
            
            result = self.search_laws(
                keyword="证券",  # 【谨慎改造】锁定关键词，只搜证券相关，避免下载大量无关文件
                category_code=category_code,
                page_num=page,
                page_size=50
            )
            rows = result.get("rows", [])
            if not rows:
                break
            for row in rows:
                bbbs = row.get("bbbs")
                if bbbs and bbbs not in existing_bbbs:
                    new_laws.append(row)
            if len(rows) < 50:
                break

        # 4. 获取新法规的下载链接
        for law in new_laws:
            # 【谨慎改造】获取详情前也加上一点微小的随机延迟
            time.sleep(random.uniform(1.0, 3.0))
            try:
                detail = self.get_law_detail(law["bbbs"])
                data = detail.get("data", {})
                oss = data.get("ossFile", {})
                law["word_url"] = oss.get("ossWordPath")
                law["pdf_url"] = oss.get("ossPdfPath")
            except Exception:
                law["word_url"] = None
                law["pdf_url"] = None

        return new_laws

    def download_law(self, law, output_dir="downloaded"):
        """下载单个法规的Word文件"""
        os.makedirs(output_dir, exist_ok=True)
        word_path = law.get("word_url")
        if not word_path:
            return None
        # 官方OSS文件路径需拼接基础URL
        url = f"https://wb.flk.npc.gov.cn/{word_path}"
        filename = f"{law.get('title', law['bbbs'])}.docx"
        filepath = os.path.join(output_dir, filename)
        resp = self.session.get(url, timeout=60, stream=True)
        resp.raise_for_status()
        with open(filepath, 'wb') as f:
            for chunk in resp.iter_content(chunk_size=8192):
                f.write(chunk)
        return filepath