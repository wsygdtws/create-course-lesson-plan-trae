# -*- coding: utf-8 -*-
"""百度网盘文件下载工具（通过 PCS API 或 MCP）"""
import requests
import urllib.parse
import os

def download_file(access_token, remote_path, local_path, timeout=60):
    """通过百度 PCS API 下载单个文件。
    access_token: 百度网盘 access_token
    remote_path: 网盘中的文件路径，如 "/学习/AI/xxx.pdf"
    local_path: 本地保存路径
    返回 True/False
    """
    encoded_path = urllib.parse.quote(remote_path, safe='')
    url = f"https://d.pcs.baidu.com/rest/2.0/pcs/file?method=download&access_token={access_token}&path={encoded_path}"
    try:
        resp = requests.get(url, stream=True, timeout=timeout)
        if resp.status_code == 200:
            os.makedirs(os.path.dirname(local_path), exist_ok=True)
            with open(local_path, 'wb') as f:
                for chunk in resp.iter_content(chunk_size=8192):
                    f.write(chunk)
            print(f"  下载成功: {remote_path} -> {local_path}")
            return True
        else:
            print(f"  下载失败 [{resp.status_code}]: {remote_path}")
            return False
    except Exception as e:
        print(f"  下载异常: {remote_path} - {e}")
        return False

def list_files(access_token, dir_path, timeout=30):
    """列出网盘目录下的文件列表。
    返回 list[dict]，每个 dict 含 path, filename, size, isdir
    """
    encoded_path = urllib.parse.quote(dir_path, safe='')
    url = f"https://pan.baidu.com/rest/2.0/xpan/file?method=list&dir={encoded_path}&access_token={access_token}"
    try:
        resp = requests.get(url, timeout=timeout)
        data = resp.json()
        if 'list' in data:
            return [{
                'path': f['path'],
                'filename': f['server_filename'],
                'size': f.get('size', 0),
                'isdir': f.get('isdir', 0) == 1
            } for f in data['list']]
        return []
    except Exception as e:
        print(f"  列目录失败: {dir_path} - {e}")
        return []

if __name__ == '__main__':
    print("此模块需配合 access_token 使用。")
    print("也可通过 Trae 的 mcp_baidu-netdisk MCP 工具直接下载，无需手动获取 token。")
    print("用法: from baidu_download_tool import download_file, list_files")
