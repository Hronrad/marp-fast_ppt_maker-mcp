import os
import sys
import shutil
import asyncio
import re
from mcp.server.fastmcp import FastMCP
from engine import EngineSplitter  

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
    local_marp = shutil.which("marp", path=local_marp_dir)
    if local_marp:
        return local_marp
    return shutil.which("marp")

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
    if sys.platform != "win32":
        env["PATH"] = "/usr/local/bin:/opt/homebrew/bin:" + env.get("PATH", "")

    final_content = content.strip()
    
    if final_content.startswith('---'):
        parts = final_content.split('---', 2)
        if len(parts) >= 3:
            final_content = parts[2].strip()

    has_manual_breaks = "\n---" in final_content
    
    if auto_split or not has_manual_breaks:
        final_content = re.sub(r'^\s*---\s*$', '', final_content, flags=re.MULTILINE)
        final_content = re.sub(r'([^\n])\n( {0,3}#{1,6}\s)', r'\1\n\n\2', final_content)
        final_content = re.sub(r'\\\((.*?)\\\)', r'$\1$', final_content)
        final_content = re.sub(r'\\\[(.*?)\\\]', r'$$\1$$', final_content, flags=re.DOTALL)
        final_content = re.sub(r'\n{3,}', '\n\n', final_content).strip()
        
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


@mcp.resource("theme://available")
def list_available_themes() -> str:
    """Get a list of all available Marp themes in the local themes directory."""
    import os
    
    base_dir = os.path.abspath(os.getcwd())
    themes_dir = os.path.join(base_dir, "themes")
    
    themes = ["default", "gaia", "uncover"]
    
    if os.path.exists(themes_dir):
        for file in os.listdir(themes_dir):
            if file.endswith(".css"):
                theme_name = file[:-4]
                themes.append(f"{theme_name} (Custom local theme)")
                
    result = "Available Marp Themes:\n"
    for t in themes:
        result += f"- {t}\n"
        
    return result


@mcp.prompt()
def academic_report_prompt(topic: str) -> str:
    """Create a structured prompt for generating a professional academic report presentation."""
    return (
        f"Please generate a comprehensive and rigorous academic presentation on the topic: '{topic}'.\n\n"
        "Execution Requirements:\n"
        "1. Content Structure: Strictly follow the standard academic format, including Abstract, Introduction, Theoretical Background, Methodology/Framework, Experimental Results, Discussion, and Conclusion.\n"
        "2. Professional Tone: Use formal, objective, and scholarly language suitable for an academic conference or defense.\n"
        "3. Mathematical Rigor: Embed necessary equations using standard LaTeX syntax ($...$ for inline, $$...$$ for block).\n"
        "4. Tool Calling Strategy: Once the content is ready, immediately call the `create_presentation` tool.\n"
        "5. Tool Parameters: Set `theme=\"academic\"` (or another suitable academic theme like `beam`), `heading_split_levels=2`, and `auto_split=True` to ensure perfect semantic pagination."
    )

if __name__ == "__main__":
    mcp.run()