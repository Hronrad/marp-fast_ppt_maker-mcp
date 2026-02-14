import os
import sys
import shutil
import asyncio
import re
from typing import Literal
from mcp.server.fastmcp import FastMCP
from playwright.async_api import async_playwright

mcp = FastMCP("Marp-PPT-Agent")

def find_browser_path():
    paths = [
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"
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

class EngineSplitter:
    def __init__(self, slide_usable_height=620):
        self.usable_height = slide_usable_height
        
    def _get_target_heading_levels(self, text: str):
        import re
        levels = set()
        for line in text.split('\n'):
            match = re.match(r'^(#{1,6})\s', line)
            if match:
                levels.add(len(match.group(1)))
        if not levels:
            return {1, 2}
        return set(sorted(list(levels))[:2])

    def _safe_chunk_text(self, text: str):
        """带智能层级追踪的切分算法"""
        import re
        lines = text.split('\n')
        chunks = []
        current_chunk = []
        current_context = []
        list_hierarchy = {} # 记录当前激活的列表层级: {缩进量: 文本}
        in_code = False
        in_math = False
        in_table = False
        table_header = ""

        def close_chunk():
            nonlocal current_chunk, current_context
            if current_chunk:
                chunks.append({"type": "text", "text": "\n".join(current_chunk), "context": list(current_context), "header": None})
                current_chunk = []

        for line in lines:
            stripped = line.strip()
            if stripped.startswith('```'):
                in_code = not in_code
            if line.count('$$') % 2 != 0:
                in_math = not in_math

            if in_code or in_math:
                if not current_chunk:
                    current_context = [list_hierarchy[k] for k in sorted(list_hierarchy.keys())]
                current_chunk.append(line)
                continue

            is_table_row = bool(re.match(r'^\|.*\|$', stripped))

            if not in_table:
                if is_table_row:
                    current_chunk.append(line)
                    if len(current_chunk) >= 2:
                        prev = current_chunk[-2].strip()
                        curr = current_chunk[-1].strip()
                        if re.match(r'^\|.*\|$', prev) and re.match(r'^\|[\s\-\|:]+\|$', curr):
                            in_table = True
                            table_header = current_chunk[-2] + "\n" + current_chunk[-1]
                            pre_table = current_chunk[:-2]
                            current_chunk = []
                            if pre_table:
                                chunks.append({"type": "text", "text": "\n".join(pre_table), "context": list(current_context), "header": None})
                            chunks.append({"type": "table_header", "text": table_header, "context": [], "header": table_header})
                            list_hierarchy = {} # 进入表格，清除列表层级记忆
                else:
                    is_list_item = bool(re.match(r'^([ \t]*)([\-\*\+]|\d+\.)\s', line))
                    is_heading = bool(re.match(r'^(#{1,6})\s', line))
                    is_blank = (stripped == '')

                    if is_blank:
                        close_chunk()
                    elif is_heading:
                        close_chunk()
                        list_hierarchy = {}
                        current_context = []
                        current_chunk = [line]
                    elif is_list_item:
                        close_chunk()
                        match = re.match(r'^([ \t]*)([\-\*\+]|\d+\.)\s', line)
                        indent = len(match.group(1).replace('\t', '    '))

                        # 清除同级或更深层级的记忆，仅保留父级
                        keys_to_remove = [k for k in list_hierarchy.keys() if k >= indent]
                        for k in keys_to_remove:
                            del list_hierarchy[k]

                        current_context = [list_hierarchy[k] for k in sorted(list_hierarchy.keys())]
                        list_hierarchy[indent] = line
                        current_chunk = [line]
                    else:
                        match = re.match(r'^([ \t]+)', line)
                        # 如果是无缩进的普通段落，说明列表已结束
                        if not match and not current_chunk:
                            list_hierarchy = {}
                        if not current_chunk:
                            current_context = [list_hierarchy[k] for k in sorted(list_hierarchy.keys())]
                        current_chunk.append(line)
            else:
                if is_table_row:
                    chunks.append({"type": "table_row", "text": line, "header": table_header, "context": []})
                else:
                    in_table = False
                    table_header = ""
                    list_hierarchy = {}
                    if stripped != '':
                        is_list_item = bool(re.match(r'^([ \t]*)([\-\*\+]|\d+\.)\s', line))
                        is_heading = bool(re.match(r'^(#{1,6})\s', line))
                        if is_list_item:
                            match = re.match(r'^([ \t]*)([\-\*\+]|\d+\.)\s', line)
                            indent = len(match.group(1).replace('\t', '    '))
                            current_context = []
                            list_hierarchy[indent] = line
                            current_chunk = [line]
                        elif is_heading:
                            current_context = []
                            current_chunk = [line]
                        else:
                            current_context = []
                            current_chunk.append(line)

        close_chunk()
        return chunks

    async def process(self, text: str, theme: str, marp_bin: str, env: dict):
        import sys, os, asyncio, re
        sys.stderr.write("DEBUG: [Two-Pass] 阶段1 - 构建探针 DOM...\n")
        
        target_levels = self._get_target_heading_levels(text)
        chunks = self._safe_chunk_text(text)
        
        probe_md_lines = [
            "---",
            "marp: true",
            f"theme: {theme}",
            "---",
            "<style>section { height: auto !important; min-height: 720px !important; padding-bottom: 50px !important; }</style>\n"
        ]
        
        for idx, chunk in enumerate(chunks):
            c_type = chunk["type"]
            c_text = chunk["text"]
            probe = f'<div class="m-probe" data-idx="{idx}" style="height:0;margin:0;padding:0;visibility:hidden;"></div>'
            
            if c_type == "table_row":
                last_pipe = c_text.rfind('|')
                row = c_text[:last_pipe] + probe + c_text[last_pipe:] if last_pipe != -1 else c_text + probe
                probe_md_lines.append(row)
            elif c_type == "table_header":
                lines = c_text.split('\n')
                last_pipe = lines[0].rfind('|')
                if last_pipe != -1:
                    lines[0] = lines[0][:last_pipe] + probe + lines[0][last_pipe:]
                probe_md_lines.append("\n".join(lines))
            else:
                probe_md_lines.append(c_text + f"\n{probe}\n")
                
        probe_md = "\n".join(probe_md_lines)
        
        base_dir = os.path.abspath(os.getcwd())
        output_dir = os.path.join(base_dir, "output_slides")
        os.makedirs(output_dir, exist_ok=True)
        probe_md_file = os.path.join(output_dir, "probe_temp.md")
        probe_html_file = os.path.join(output_dir, "probe_temp.html")
        
        with open(probe_md_file, "w", encoding="utf-8") as f:
            f.write(probe_md)
            
        proc = await asyncio.create_subprocess_exec(
            marp_bin, probe_md_file, "-o", probe_html_file, "--html", "--allow-local-files",
            stdout=asyncio.subprocess.PIPE, stderr=asyncio.subprocess.PIPE, stdin=asyncio.subprocess.DEVNULL, env=env
        )
        await proc.communicate()
        
        sys.stderr.write("DEBUG: [Two-Pass] 阶段2 - Chromium 物理测距...\n")
        from playwright.async_api import async_playwright
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True, executable_path=env.get("CHROME_PATH"))
            page = await browser.new_page()
            await page.goto(f"file://{probe_html_file}")
            await page.wait_for_timeout(1000) 
            
            js_code = """
            () => {
                const probes = Array.from(document.querySelectorAll('.m-probe'));
                const rootTop = document.querySelector('section').getBoundingClientRect().top;
                return probes.map(p => {
                    const rect = p.getBoundingClientRect();
                    return { idx: parseInt(p.dataset.idx), y: rect.bottom - rootTop };
                });
            }
            """
            probe_data = await page.evaluate(js_code)
            await browser.close()

        sys.stderr.write("DEBUG: [Two-Pass] 阶段3 - 物理边界与上下文智能注入...\n")
        final_lines = []
        current_page_baseline = 0
        
        for data in probe_data:
            idx = data["idx"]
            y_pos = data["y"]
            chunk = chunks[idx]
            
            is_target_heading = False
            if chunk["type"] == "text":
                match = re.match(r'^(#{1,6})\s', chunk["text"])
                if match and len(match.group(1)) in target_levels:
                    is_target_heading = True
            
            is_overflow = (y_pos - current_page_baseline) > self.usable_height
            is_first_on_page = (idx == 0) or (current_page_baseline == probe_data[idx-1]["y"])
            
            if is_overflow or (is_target_heading and not is_first_on_page):
                if idx > 0:
                    final_lines.append("\n---\n")
                    current_page_baseline = probe_data[idx-1]["y"] 
                    
                    if chunk.get("type") == "table_row":
                        final_lines.append(chunk["header"])
                        
                    # --- 核心修复：列表跨页智能注入父级上下文 ---
                    if chunk.get("context"):
                        for ctx_line in chunk["context"]:
                            final_lines.append(ctx_line)
                        
            final_lines.append(chunk["text"])

        try:
            os.remove(probe_md_file)
            os.remove(probe_html_file)
        except:
            pass

        return "\n".join(final_lines)
    
    
# 核心修复4：将 MCP 接口设定为 async def
@mcp.tool()
async def create_presentation(
    title: str,
    content: str,
    theme: Literal["default", "gaia", "uncover"] = "default",
    style_class: str = "lead",
    auto_split: bool = True 
) -> str:
    """将 Markdown 内容转换为 PPTX 和 PDF。采用双重渲染确保物理高度精准。"""
    marp_bin = find_marp_executable()
    if not marp_bin:
        return "❌ 错误: 找不到 Marp。"
    browser_path = find_browser_path()
    if not browser_path:
        return "❌ 错误: 找不到浏览器。"

    env = os.environ.copy()
    env["CHROME_PATH"] = browser_path
    env["PATH"] = "/usr/local/bin:/opt/homebrew/bin:" + env.get("PATH", "")

    # --- 逻辑层 ---
    final_content = content.strip()
    
    # 1. 剥离可能自带的 Frontmatter
    if final_content.startswith('---'):
        parts = final_content.split('---', 2)
        if len(parts) >= 3:
            final_content = parts[2].strip()

    has_manual_breaks = "\n---" in final_content
    
    if auto_split or not has_manual_breaks:
        # 核心修正：强制清除文中所有已存在的 --- 分页符，避免干扰排版引擎
        final_content = re.sub(r'^\s*---\s*$', '', final_content, flags=re.MULTILINE)
        
        # 自动规范化公式格式：将 \( \) 替换为 $ $，将 \[ \] 替换为 $$ $$
        final_content = re.sub(r'\\\((.*?)\\\)', r'$\1$', final_content)
        final_content = re.sub(r'\\\[(.*?)\\\]', r'$$\1$$', final_content, flags=re.DOTALL)
        
        # 删除多余的空白行（将3个以上的换行压缩为2个），保持排版紧密
        final_content = re.sub(r'\n{3,}', '\n\n', final_content).strip()
        
        # 触发 Two-Pass 物理引擎
        splitter = EngineSplitter(slide_usable_height=620)
        final_content = await splitter.process(final_content, theme, marp_bin, env)

    header = f"---\nmarp: true\ntheme: {theme}\nclass: {style_class}\npaginate: true\n---\n\n"
    full_markdown = header + final_content

    base_dir = os.path.abspath(os.getcwd())
    output_dir = os.path.join(base_dir, "output_slides")
    os.makedirs(output_dir, exist_ok=True)
    
    md_file = os.path.join(output_dir, f"{title}.md")
    pptx_file = os.path.join(output_dir, f"{title}.pptx")
    pdf_file = os.path.join(output_dir, f"{title}.pdf")
    
    with open(md_file, "w", encoding="utf-8") as f:
        f.write(full_markdown)

    results = []
    
    # 核心修复6：使用纯异步的方式调用最终生成命令
    async def run_marp_async(output_path, format_flag):
        cmd = [marp_bin, md_file, "-o", output_path, "--allow-local-files"]
        try:
            proc = await asyncio.create_subprocess_exec(
                *cmd,
                stdout=asyncio.subprocess.PIPE,
                stderr=asyncio.subprocess.PIPE,
                stdin=asyncio.subprocess.DEVNULL,
                env=env
            )
            stdout, stderr = await asyncio.wait_for(proc.communicate(), timeout=120)
            if proc.returncode != 0:
                return False, stderr.decode()
            return True, None
        except asyncio.TimeoutError:
            return False, "命令执行超时"
        except Exception as e:
            return False, str(e)

    ok, err = await run_marp_async(pptx_file, "PPTX")
    results.append(f"✅ PPTX: {pptx_file}" if ok else f"❌ PPTX 失败: {err}")

    ok, err = await run_marp_async(pdf_file, "PDF")
    results.append(f"✅ PDF:  {pdf_file}" if ok else f"❌ PDF 失败: {err}")

    return "\n".join(results)

if __name__ == "__main__":
    mcp.run()