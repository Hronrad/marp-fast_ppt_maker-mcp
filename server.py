import os
import sys
import shutil
import asyncio
import re
from mcp.server.fastmcp import FastMCP
from playwright.async_api import async_playwright

mcp = FastMCP("Marp-fast PPT maker-Agent")

def find_browser_path():
    """Cross-platform browser path detection."""
    if sys.platform == "win32":
        paths = [
            os.path.expandvars(r"%ProgramFiles%\Google\Chrome\Application\chrome.exe"),
            os.path.expandvars(r"%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"),
            os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
            os.path.expandvars(r"%ProgramFiles%\Microsoft\Edge\Application\msedge.exe"),
            os.path.expandvars(r"%ProgramFiles(x86)%\Microsoft\Edge\Application\msedge.exe"),
        ]
    elif sys.platform == "darwin":
        paths = [
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"
        ]
    else:  # Linux and others
        paths = [
            "/usr/bin/google-chrome",
            "/usr/bin/google-chrome-stable",
            "/usr/bin/microsoft-edge-stable",
            "/usr/bin/microsoft-edge"
        ]
        
    for p in paths:
        if os.path.exists(p):
            return p
    return None

def find_marp_executable():
    """Cross-platform Marp executable detection."""
    local_marp_dir = os.path.abspath(os.path.join(os.getcwd(), "node_modules", ".bin"))
    # shutil.which auto-adds .cmd/.exe on Windows
    local_marp = shutil.which("marp", path=local_marp_dir)
    if local_marp:
        return local_marp
    return shutil.which("marp")

class EngineSplitter:
    def __init__(self, slide_usable_height=620):
        self.usable_height = slide_usable_height
        
    def _get_target_heading_levels(self, text: str, split_levels: int):
            levels = set()
            for line in text.split('\n'):
                match = re.match(r'^(#{1,6})\s', line)
                if match:
                    levels.add(len(match.group(1)))
            if not levels:
                return {1, 2}
            # Core change: use split_levels for top-N heading levels
            return set(sorted(list(levels))[:split_levels])

    def _safe_chunk_text(self, text: str):
        lines = text.split('\n')
        chunks = []
        current_chunk = []
        current_context = []
        list_hierarchy = {}
        in_code = False
        in_math = False
        in_table = False
        table_header = ""
        
        # Core fix: remember blank-line state before a chunk
        pending_blank = False

        def close_chunk(force_type="text", header=None):
            nonlocal current_chunk, current_context, pending_blank
            if current_chunk:
                chunks.append({
                    "type": force_type, 
                    "text": "\n".join(current_chunk), 
                    "context": list(current_context), 
                    "header": header,
                    "blank_before": pending_blank
                })
                current_chunk = []
                pending_blank = False

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
                                chunks.append({
                                    "type": "text", 
                                    "text": "\n".join(pre_table), 
                                    "context": list(current_context), 
                                    "header": None,
                                    "blank_before": pending_blank
                                })
                                pending_blank = False
                            chunks.append({
                                "type": "table_header", 
                                "text": table_header, 
                                "context": [], 
                                "header": table_header,
                                "blank_before": pending_blank
                            })
                            pending_blank = False
                            list_hierarchy = {}
                else:
                    is_list_item = bool(re.match(r'^([ \t]*)([\-\*\+]|\d+\.)\s', line))
                    is_heading = bool(re.match(r'^(#{1,6})\s', line))
                    is_blank = (stripped == '')

                    if is_blank:
                        close_chunk()
                        # Record blank-line state for the next chunk to claim
                        pending_blank = True
                    elif is_heading:
                        close_chunk()
                        list_hierarchy = {}
                        current_context = []
                        current_chunk = [line]
                    elif is_list_item:
                        close_chunk()
                        match = re.match(r'^([ \t]*)([\-\*\+]|\d+\.)\s', line)
                        indent = len(match.group(1).replace('\t', '    '))

                        keys_to_remove = [k for k in list_hierarchy.keys() if k >= indent]
                        for k in keys_to_remove:
                            del list_hierarchy[k]

                        current_context = [list_hierarchy[k] for k in sorted(list_hierarchy.keys())]
                        list_hierarchy[indent] = line
                        current_chunk = [line]
                    else:
                        match = re.match(r'^([ \t]+)', line)
                        if not match and not current_chunk:
                            list_hierarchy = {}
                        if not current_chunk:
                            current_context = [list_hierarchy[k] for k in sorted(list_hierarchy.keys())]
                        current_chunk.append(line)
            else:
                if is_table_row:
                    chunks.append({
                        "type": "table_row", 
                        "text": line, 
                        "header": table_header, 
                        "context": [],
                        "blank_before": False
                    })
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
                    else:
                        pending_blank = True

        close_chunk()
        return chunks

    async def process(self, text: str, theme: str, marp_bin: str, env: dict, heading_split_levels: int = 2):
        sys.stderr.write("DEBUG: [Two-Pass] Phase 1 - Build probe DOM...\n")
        
        target_levels = self._get_target_heading_levels(text, heading_split_levels)
        chunks = self._safe_chunk_text(text)
        
        probe_md_lines = [
            "---",
            "marp: true",
            f"theme: {theme}",
            "---",
            "<style>section { height: auto !important; overflow: visible !important; }</style>\n"
        ]
        
        for idx, chunk in enumerate(chunks):
            # Core fix: restore paragraph spacing when rebuilding the DOM
            if chunk.get("blank_before") and idx > 0:
                probe_md_lines.append("") 
                
            c_type = chunk["type"]
            c_text = chunk["text"]
            probe = f'<span class="m-probe" data-idx="{idx}" style="font-size:0; line-height:0; margin:0; padding:0; visibility:hidden;"></span>'
            
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
                if c_text.strip().endswith('```') or c_text.strip().endswith('$$'):
                    probe_md_lines.append(c_text + f"\n{probe}\n")
                else:
                    probe_md_lines.append(c_text + probe)
                
        probe_md = "\n".join(probe_md_lines)
        
        base_dir = os.path.abspath(os.getcwd())
        output_dir = os.path.join(base_dir, "output_slides")
        os.makedirs(output_dir, exist_ok=True)
        probe_md_file = os.path.join(output_dir, "probe_temp.md")
        probe_html_file = os.path.join(output_dir, "probe_temp.html")
        
        with open(probe_md_file, "w", encoding="utf-8") as f:
            f.write(probe_md)
            
        cmd = [marp_bin, probe_md_file, "-o", probe_html_file, "--html", "--allow-local-files"]
        themes_dir = os.path.join(base_dir, "themes")
        if os.path.exists(themes_dir):
            cmd.extend(["--theme-set", themes_dir])

        proc = await asyncio.create_subprocess_exec(
            *cmd,
            stdout=asyncio.subprocess.PIPE, stderr=asyncio.subprocess.PIPE, stdin=asyncio.subprocess.DEVNULL, env=env
        )
        await proc.communicate()
        
        sys.stderr.write("DEBUG: [Two-Pass] Phase 2 - Chromium physical measurement...\n")

        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=True, executable_path=env.get("CHROME_PATH"))
            page = await browser.new_page()
            await page.goto(f"file://{probe_html_file}")
            await page.wait_for_timeout(1000) 
            
            js_code = """
            () => {
                const section = document.querySelector('section');
                const style = window.getComputedStyle(section);
                const pt = parseFloat(style.paddingTop) || 0;
                const pb = parseFloat(style.paddingBottom) || 0;
                const usableHeight = 720 - pt - pb;

                const probes = Array.from(document.querySelectorAll('.m-probe'));
                const contentTop = section.getBoundingClientRect().top + pt;

                return {
                    usableHeight: usableHeight,
                    probes: probes.map(p => {
                        let target = p.parentElement;
                        const blockTags = ['li', 'p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'tr', 'div', 'blockquote', 'pre'];
                        while (target && !blockTags.includes(target.tagName.toLowerCase()) && target.tagName.toLowerCase() !== 'section') {
                            target = target.parentElement;
                        }
                        if (!target || target.tagName.toLowerCase() === 'section') {
                            target = p.parentElement || p;
                        }
                        
                        const rect = target.getBoundingClientRect();
                        const tStyle = window.getComputedStyle(target);
                        const mb = parseFloat(tStyle.marginBottom) || 0;
                        
                        return { idx: parseInt(p.dataset.idx), y: rect.bottom + mb - contentTop };
                    })
                };
            }
            """
            result = await page.evaluate(js_code)
            await browser.close()

        usable_height = result["usableHeight"]
        probe_data = result["probes"]
        safe_usable_height = usable_height - 30

        sys.stderr.write("DEBUG: [Two-Pass] Phase 3 - Physical boundary measurement...\n")
        final_lines = []
        current_page_baseline = 0
        
        for data in probe_data:
            idx = data["idx"]
            y_pos = data["y"]
            chunk = chunks[idx]
            
            is_target_heading = False
            if chunk["type"] == "text":
                match = re.match(r'^ {0,3}(#{1,6})\s', chunk["text"])
                if match and len(match.group(1)) in target_levels:
                    is_target_heading = True
            
            is_overflow = (y_pos - current_page_baseline) > safe_usable_height
            is_first_on_page = (idx == 0) or (current_page_baseline == probe_data[idx-1]["y"])
            
            if (is_overflow or is_target_heading) and not is_first_on_page:
                if idx > 0:
                    final_lines.append("\n---\n")
                    current_page_baseline = probe_data[idx-1]["y"] 
                    
                    if chunk.get("type") == "table_row":
                        final_lines.append(chunk["header"])
                        
                    if chunk.get("context"):
                        for ctx_line in chunk["context"]:
                            final_lines.append(ctx_line)
            else:
                # Core fix: preserve blank paragraph spacing when no page break
                if chunk.get("blank_before") and len(final_lines) > 0 and final_lines[-1] != "\n---\n":
                    final_lines.append("")
                        
            final_lines.append(chunk["text"])

        try:
            os.remove(probe_md_file)
            os.remove(probe_html_file)
        except:
            pass

        return "\n".join(final_lines)

@mcp.tool()
async def create_presentation(
    title: str,
    content: str,
    theme: str = "default",
    style_class: str = "",
    auto_split: bool = True,
    generate_pptx: bool = True,
    heading_split_levels: int = 2
) -> str:
    """
    One-click tool to convert Markdown into PPT-style slides. It auto-splits content and
    uses Marp to generate PPTX and PDF. The two-pass rendering engine ensures accurate
    physical layout and smart splitting for better readability. Ideal for fast generation
    of professional decks such as team reports, academic talks, or content sharing.

    Parameters:
    - generate_pptx: Whether to generate the PPTX file. Default is True.

    [LLM Theme Guide] Choose the best theme based on the content:
    - "default": Small font, clean black-on-white, best compatibility.
    - "gaia": Medium font, warm tone, low contrast. Good for humanities, art/design,
        eco/lifestyle topics.
    - "uncover": Large font, minimalist, high contrast. Good for product launches,
        TED-style talks, creative pitches, and image-heavy decks.
    - "academic": Medium font with red titles. Note: right-aligned; use only when needed.
    - "beam": Small font, Beamer-like. Good for academic content, but less compatible
        with long titles and complex elements.
    - "rose-pine-dawn": Small font, light background, gentle style.
    - "rose-pine-moon": Small font, dark background, elegant for dark themes.
    - "rose-pine-dawn-modern": Medium font, adds a modern card-style title on top of
        rose-pine-dawn.

    [style_class Guide] Optional class to tweak layout:
    - "": Default style, suitable for most cases.
    - "lead": Centered title layout similar to uncover.
    - "invert": Inverted colors for dark presentations.

    [heading_split_levels Guide] Controls heading-driven page breaks (default: 2):
    - 2 (default): The top two heading levels (e.g., H1 and H2) trigger new pages.
    - 1: Only the top-level headings (e.g., H1) trigger page breaks.
    - 3 or more: For very deep documents where each subsection is long.
    """
    marp_bin = find_marp_executable()
    if not marp_bin:
        return "❌ Error: Marp not found."
    browser_path = find_browser_path()
    if not browser_path:
        return "❌ Error: Browser not found."

    env = os.environ.copy()
    env["CHROME_PATH"] = browser_path
    # Only append extra bins on non-Windows to avoid breaking PATH
    if sys.platform != "win32":
        env["PATH"] = "/usr/local/bin:/opt/homebrew/bin:" + env.get("PATH", "")

    # --- Logic ---
    final_content = content.strip()
    
    # 1. Strip any existing frontmatter
    if final_content.startswith('---'):
        parts = final_content.split('---', 2)
        if len(parts) >= 3:
            final_content = parts[2].strip()

    has_manual_breaks = "\n---" in final_content
    
    if auto_split or not has_manual_breaks:
        # Remove any existing page breaks
        final_content = re.sub(r'^\s*---\s*$', '', final_content, flags=re.MULTILINE)
        
        # Core fix: ensure a blank line before headings to avoid block parsing stickiness
        final_content = re.sub(r'([^\n])\n( {0,3}#{1,6}\s)', r'\1\n\n\2', final_content)
        
        # Normalize math syntax
        final_content = re.sub(r'\\\((.*?)\\\)', r'$\1$', final_content)
        final_content = re.sub(r'\\\[(.*?)\\\]', r'$$\1$$', final_content, flags=re.DOTALL)
        
        final_content = re.sub(r'\n{3,}', '\n\n', final_content).strip()
        
        # Run the two-pass engine with dynamic heading split levels
        splitter = EngineSplitter(slide_usable_height=620)
        final_content = await splitter.process(final_content, theme, marp_bin, env, heading_split_levels)

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
    
    # Use async subprocess for final generation
    async def run_marp_async(output_path, format_flag):
        cmd = [marp_bin, md_file, "-o", output_path, "--allow-local-files"]

        themes_dir = os.path.join(base_dir, "themes")
        if os.path.exists(themes_dir):
            cmd.extend(["--theme-set", themes_dir])

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
            return False, "Command timed out"
        except Exception as e:
            return False, str(e)

    if generate_pptx:
        ok, err = await run_marp_async(pptx_file, "PPTX")
        results.append(f"✅ PPTX: {pptx_file}" if ok else f"❌ PPTX failed: {err}")

    ok, err = await run_marp_async(pdf_file, "PDF")
    results.append(f"✅ PDF:  {pdf_file}" if ok else f"❌ PDF failed: {err}")

    return "\n".join(results)

if __name__ == "__main__":
    mcp.run()