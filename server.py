import os
import sys
import shutil
import subprocess
import re
import math
from typing import Literal
from mcp.server.fastmcp import FastMCP

# 初始化
mcp = FastMCP("Marp-PPT-Agent")
max_chars_per_slide = 1200 # PPT 每页字符上限（可调整）

# --- 1. 基础设施 ---
def find_browser_path():
    paths = [
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
        "/Applications/Brave Browser.app/Contents/MacOS/Brave Browser"
    ]
    for p in paths:
        if os.path.exists(p):
            return p
    return None

def find_marp_executable():
    local_marp = os.path.abspath(os.path.join(os.getcwd(), "node_modules", ".bin", "marp"))
    if os.path.exists(local_marp):
        return local_marp
    return shutil.which("marp")

# --- 2. 核心算法：严格防溢出切分 (Strict Overflow Prevention) ---

# --- 2. 核心算法：动态层级 + 视觉权重切分 ---

def _get_target_heading_levels(text: str):
    """
    预扫描文本，找出文中存在的最高级和次高级标题。
    返回一个集合，例如 {2, 3} 代表只在 ## 和 ### 处强制切分。
    """
    levels = set()
    # 扫描文中所有的标题层级
    for line in text.split('\n'):
        # 匹配标准Markdown标题
        match = re.match(r'^(#{1,6})\s', line)
        if match:
            levels.add(len(match.group(1)))
    
    if not levels:
        return {1, 2} # 默认兜底

    sorted_levels = sorted(list(levels))
    
    # 策略：只锁定最高级和次高级
    # 例如：文中只有 H2, H3, H4 -> 锁定 {2, 3}，H4 不强制分页
    target_levels = set(sorted_levels[:2])
    
    sys.stderr.write(f"DEBUG: 动态层级检测结果: H{target_levels}\n")
    return target_levels


class MarkdownSplitter:
    def __init__(self, target_levels, max_cost=1200): 
        self.new_lines = []
        self.current_cost = 0
        self.max_cost = max_cost
        self.target_levels = target_levels
        self.in_code_block = False
        self.in_math_block = False 
        self.page_count = 0
        self.cost_log = []
        
        self.img_pattern = re.compile(r'!\[.*?\]\(.*?\)') 
        self.math_pattern = re.compile(r'\$\$') 
        self.list_pattern = re.compile(r'^(\d+\.|-|\*)\s')

    def _get_visual_cost(self, line: str) -> int:
        s_line = line.strip()
        if not s_line: return 50 
        if self.img_pattern.search(line): return 320
        if self.math_pattern.search(line): return 160
        
        header_match = re.match(r'^(#{1,6})\s', line)
        if header_match:
            level = len(header_match.group(1))
            multiplier = 2.8 - 0.3 * level
            cost = int(multiplier * math.ceil(len(s_line) / 12) * 50)
            return cost

        if self.list_pattern.match(line): return len(line) + 30
        return len(line)

    def safe_add_break(self):
        while self.new_lines and self.new_lines[-1].strip() == "":
            self.new_lines.pop()
        
        if self.new_lines and self.new_lines[-1].strip() == "---":
            self.current_cost = 0
            return 
        
        if self.current_cost <= 0:
            return
            
        self.page_count += 1
        log_msg = f"【page {self.page_count}】Cost: {self.current_cost}"
        self.cost_log.append(log_msg)
        sys.stderr.write(log_msg + "\n")
        
        self.new_lines.append("\n---\n\n")
        self.current_cost = 0 

    def _get_block_cost(self, lines, start_idx):
        """前瞻计算整个代码块/公式块的总Cost"""
        cost = 0
        stripped_start = lines[start_idx].strip()
        is_math = stripped_start == '$$'
        is_code = stripped_start.startswith('```')
        
        for j in range(start_idx, len(lines)):
            cost += self._get_visual_cost(lines[j])
            if j > start_idx:
                stripped_curr = lines[j].strip()
                # 遇到闭合标签，停止前瞻
                if is_math and stripped_curr == '$$':
                    break
                if is_code and stripped_curr.startswith('```'):
                    break
        return cost

    def process(self, text):
        lines = text.split('\n')
        
        for i, line in enumerate(lines):
            line_cost = self._get_visual_cost(line)
            stripped_line = line.strip()
            
            # 检测是否即将进入代码块/公式块
            starts_code = (not self.in_code_block) and stripped_line.startswith('```')
            starts_math = (not self.in_math_block) and stripped_line == '$$'
            
            # --- 块级前瞻预判 (Block Lookahead) ---
            if starts_code or starts_math:
                block_cost = self._get_block_cost(lines, i)
                # 如果加上整个块会溢出，立刻在块前面切分！
                if (self.current_cost + block_cost) > self.max_cost and self.current_cost > 0:
                    self.safe_add_break()
            
            # 1. 状态机：代码块保护
            if stripped_line.startswith('```'):
                self.in_code_block = not self.in_code_block
                self.new_lines.append(line)
                self.current_cost += line_cost
                continue

            # 2. 状态机：数学公式块保护
            if stripped_line == '$$':
                self.in_math_block = not self.in_math_block
                self.new_lines.append(line)
                self.current_cost += line_cost
                continue

            # 如果在代码块或公式块内部，绝对不切分，直接追加
            if self.in_code_block or self.in_math_block:
                self.new_lines.append(line)
                self.current_cost += line_cost
                continue

            # --- 逻辑 A: 动态标题强制切分 ---
            header_match = re.match(r'^(#{1,6})\s', line)
            if header_match:
                level = len(header_match.group(1))
                if level in self.target_levels:
                    if self.current_cost > 0 and i > 0: 
                        self.safe_add_break()
                    self.new_lines.append(line)
                    self.current_cost += line_cost 
                    continue

            # --- 逻辑 B: 严格预判切分 ---
            if (self.current_cost + line_cost) > self.max_cost:
                if self.current_cost > 0:
                    self.safe_add_break()
                    self.new_lines.append(line)
                    self.current_cost = line_cost
                    continue
            
            self.new_lines.append(line)
            self.current_cost += line_cost
            
        if self.current_cost > 0:
            self.page_count += 1
            log_msg = f"【page {self.page_count}】Cost: {self.current_cost}"
            self.cost_log.append(log_msg)
            sys.stderr.write(log_msg + "\n")
        
        return "\n".join(self.new_lines)

def _smart_split_markdown(text: str):
    # 0. 预检
    if text.count('\n---') > 3:
        return text, []
    
    # 1. 动态获取文档结构
    target_levels = _get_target_heading_levels(text)
    
    # 2. 传入层级和权重阈值
    splitter = MarkdownSplitter(target_levels=target_levels, max_cost=max_chars_per_slide)
    result = splitter.process(text)
    return result, splitter.cost_log

# --- 3. MCP 工具定义 ---
@mcp.tool()
def create_presentation(
    title: str,
    content: str,
    theme: Literal["default", "gaia", "uncover"] = "default",
    style_class: str = "lead",
    auto_split: bool = True 
) -> str:
    """
    将 Markdown 内容转换为 PPTX 和 PDF。
    """
    
    # --- 环境检查 ---
    marp_bin = find_marp_executable()
    if not marp_bin:
        return "❌ 错误: 找不到 Marp。"
    browser_path = find_browser_path()
    if not browser_path:
        return "❌ 错误: 找不到浏览器。"

# --- 逻辑层 ---
    
    # 核心修正2：在所有处理开始前，先剥离原始的 frontmatter
    # 防止其自身的 '---' 触发分页，也防止其内容被算作第一页的 Cost
    final_content = content.strip()
    if final_content.startswith('---'):
        parts = final_content.split('---', 2)
        if len(parts) >= 3:
            final_content = parts[2].strip()

    # 在剥离头部后，再判断文中是否还有手动分页符
    has_manual_breaks = "\n---" in final_content
    cost_log = []
    
    if auto_split or not has_manual_breaks:
        sys.stderr.write("DEBUG: 启动智能切分..\n")
        final_content, cost_log = _smart_split_markdown(final_content)

    # 注入新的 Header
    header = f"---\nmarp: true\ntheme: {theme}\nclass: {style_class}\npaginate: true\n---\n\n"
    full_markdown = header + final_content


    # --- IO层 ---
    base_dir = os.path.abspath(os.getcwd())
    output_dir = os.path.join(base_dir, "output_slides")
    os.makedirs(output_dir, exist_ok=True)
    
    md_file = os.path.join(output_dir, f"{title}.md")
    pptx_file = os.path.join(output_dir, f"{title}.pptx")
    pdf_file = os.path.join(output_dir, f"{title}.pdf")
    
    with open(md_file, "w", encoding="utf-8") as f:
        f.write(full_markdown)

    # --- 执行层 ---
    env = os.environ.copy()
    env["CHROME_PATH"] = browser_path
    env["PATH"] = "/usr/local/bin:/opt/homebrew/bin:" + env.get("PATH", "")

    results = []
    
    def run_marp(output_path, format_flag):
        cmd = [marp_bin, md_file, "-o", output_path, "--allow-local-files"]
        try:
            subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                stdin=subprocess.DEVNULL,
                env=env,
                text=True,
                check=True,
                timeout=120
            )
            return True, None
        except Exception as e:
            err_msg = e.stderr if isinstance(e, subprocess.CalledProcessError) else str(e)
            return False, err_msg

    # 1. PPTX
    ok, err = run_marp(pptx_file, "PPTX")
    results.append(f"✅ PPTX: {pptx_file}" if ok else f"❌ PPTX 失败: {err}")

    # 2. PDF
    ok, err = run_marp(pdf_file, "PDF")
    results.append(f"✅ PDF:  {pdf_file}" if ok else f"❌ PDF 失败: {err}")

    # 3. 添加 cost 信息到结果
    final_result = "\n".join(results)
    if cost_log:
        final_result += "\n\n📊 转换详情:\n" + "\n".join(cost_log)
    
    return final_result

if __name__ == "__main__":
    mcp.run()