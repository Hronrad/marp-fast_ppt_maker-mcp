import os
import sys
import subprocess
import shutil
from typing import Literal
from mcp.server.fastmcp import FastMCP

# 初始化
sys.stderr.write("[STARTUP] Server is initiating...\n")
mcp = FastMCP("Marp-PPT-Agent")

# --- 辅助函数：寻找浏览器 ---
def find_browser_path():
    """在 macOS 上自动寻找 Chrome 或 Edge"""
    paths = [
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
        "/Applications/Brave Browser.app/Contents/MacOS/Brave Browser"
    ]
    for p in paths:
        if os.path.exists(p):
            return p
    return None

# --- 辅助函数：寻找 Marp ---
def find_marp_executable():
    """优先寻找本地 node_modules，其次寻找全局安装"""
    # 1. 检查当前项目下的 node_modules
    local_marp = os.path.abspath(os.path.join(os.getcwd(), "node_modules", ".bin", "marp"))
    if os.path.exists(local_marp):
        return local_marp
    
    # 2. 检查全局路径 (shutil.which)
    global_marp = shutil.which("marp")
    if global_marp:
        return global_marp
        
    return None

@mcp.tool()
def create_presentation(title: str, content: str, theme: str = "default") -> str:
    # 1. 环境自检
    marp_bin = find_marp_executable()
    if not marp_bin:
        return "❌ 严重错误: 找不到 marp 可执行文件。请运行 'npm install @marp-team/marp-cli' 安装。"
        
    browser_path = find_browser_path()
    if not browser_path:
        return "❌ 严重错误: 找不到 Google Chrome 或 Edge 浏览器。Marp 无法生成 PPT。"

    # 2. 准备文件
    base_dir = os.path.abspath(os.getcwd())
    output_dir = os.path.join(base_dir, "output_slides")
    os.makedirs(output_dir, exist_ok=True)
    
    md_file = os.path.join(output_dir, f"{title}.md")
    pptx_file = os.path.join(output_dir, f"{title}.pptx")
    
    # 写入文件
    header = f"---\nmarp: true\ntheme: {theme}\n---\n\n"
    with open(md_file, "w", encoding="utf-8") as f:
        f.write(header + content)

    # 3. 准备环境变量 (注入 CHROME_PATH)
    env = os.environ.copy()
    env["CHROME_PATH"] = browser_path
    # 修复 PATH，确保 node 能被找到
    env["PATH"] = "/usr/local/bin:/opt/homebrew/bin:" + env.get("PATH", "")

    # 4. 构建命令
    cmd = [marp_bin, md_file, "-o", pptx_file, "--allow-local-files"]

    # 打印调试信息
    sys.stderr.write(f"DEBUG: 使用 Marp: {marp_bin}\n")
    sys.stderr.write(f"DEBUG: 使用 Browser: {browser_path}\n")
    sys.stderr.write(f"DEBUG: 执行命令: {' '.join(cmd)}\n")
    sys.stderr.flush()

    try:
        # 5. 执行命令 (关键：stdin=DEVNULL)
        result = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            stdin=subprocess.DEVNULL, # <--- 核心修复：防止死锁
            env=env,
            text=True,
            timeout=45, # 给足时间
            check=True
        )
        
        return f"✅ 成功！已生成: {pptx_file}"

    except subprocess.TimeoutExpired as e:
        # 捕获超时时的输出
        stdout_log = e.stdout if e.stdout else "无"
        stderr_log = e.stderr if e.stderr else "无"
        sys.stderr.write(f"ERROR: 超时详情 - stdout: {stdout_log}\n")
        sys.stderr.write(f"ERROR: 超时详情 - stderr: {stderr_log}\n")
        return f"❌ 超时错误 (45s)。Marp 可能卡在启动浏览器上。\n调试日志:\nSTDOUT: {stdout_log}\nSTDERR: {stderr_log}"
        
    except subprocess.CalledProcessError as e:
        return f"❌ Marp 报错 (Code {e.returncode}):\n{e.stderr}"
    except Exception as e:
        return f"❌ 未知 Python 错误: {str(e)}"

if __name__ == "__main__":
    mcp.run()