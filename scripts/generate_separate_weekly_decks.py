from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.util import Inches, Pt


OUT_DIR = Path("slides/weekly")

INK = RGBColor(23, 32, 51)
MUTED = RGBColor(86, 101, 123)
WHITE = RGBColor(255, 255, 255)
LIGHT = RGBColor(244, 247, 251)
LINE = RGBColor(217, 226, 238)
DARK = RGBColor(15, 23, 42)
BLUE = RGBColor(37, 99, 235)
GREEN = RGBColor(15, 159, 110)
ORANGE = RGBColor(249, 115, 22)
PURPLE = RGBColor(124, 58, 237)
RED = RGBColor(220, 38, 38)
YELLOW_BG = RGBColor(255, 247, 237)
BLUE_BG = RGBColor(239, 246, 255)
GREEN_BG = RGBColor(236, 253, 245)
PURPLE_BG = RGBColor(245, 243, 255)


LESSONS = [
    {
        "week": 1,
        "title": "Meet Your AI Teammate",
        "subtitle": "ChatGPT basics + prompt formula",
        "color": BLUE,
        "goal": "Students learn that AI can be a tutor, coach, brainstorming partner, and project helper.",
        "output": "Personal AI assistant with 5 example conversations.",
        "flow": ["0-10 Hook: Can AI do my homework?", "10-25 Demo: weak vs strong prompt", "25-70 Build: custom assistant", "70-85 Partner test", "85-90 Wrap"],
        "interactive": ["Thumb vote: helpful or lazy?", "Pair test: ask your partner's assistant one question", "Debug moment: make a vague prompt clearer"],
        "prompt": "Act as a friendly study coach for a high school student.\nHelp me study [topic]. Ask one question at a time. If I get it wrong, explain simply and let me try again.",
        "homework": "Create 5 example conversations with your assistant. Replace this with your own homework if needed.",
    },
    {
        "week": 2,
        "title": "VS Code + Codex Setup",
        "subtitle": "Folders, files, preview, first web UI",
        "color": GREEN,
        "goal": "Students open a project folder, understand files, use VS Code/Codex, and preview a simple page.",
        "output": "Working coding folder plus a personalized AI Builder profile page.",
        "flow": ["0-10 Hook: AI inside your editor", "10-30 Setup walkthrough", "30-45 Demo: create index.html", "45-80 Build: profile page", "80-90 Share one change"],
        "interactive": ["Setup checkpoint: everyone shows folder", "Choice board: colors, sections, button action", "Explain-back: what is index.html?"],
        "prompt": "Create a beginner-friendly personal web page in one file named index.html. Theme: AI Builder profile. Include name, favorite apps, project ideas, and one button.",
        "homework": "Customize the page with one personal section and one visual style change. Replace this later if needed.",
    },
    {
        "week": 3,
        "title": "AI Game Studio",
        "subtitle": "Build and improve a browser game",
        "color": ORANGE,
        "goal": "Students use Codex to create, test, debug, and improve a simple browser game.",
        "output": "Playable HTML/CSS/JavaScript game.",
        "flow": ["0-10 Hook: Can AI make a game?", "10-25 Demo: tiny clicker game", "25-65 Build: choose game type", "65-82 Playtest", "82-90 Add one feature plan"],
        "interactive": ["Game menu: clicker, dodge, quiz", "Bug hunt: what broke?", "Peer playtest: one compliment, one upgrade"],
        "prompt": "Create a simple browser clicker game in one HTML file. Add score, 30-second timer, game over message, clean styling, and comments explaining the JavaScript.",
        "homework": "Add one feature: score effect, levels, sound, timer change, or game-over screen. Edit this homework as desired.",
    },
    {
        "week": 4,
        "title": "AI Agents + MCP",
        "subtitle": "Tool-using AI + YouTube research workflow",
        "color": PURPLE,
        "goal": "Students learn the difference between chatbot and agent, and see MCP as a connector for tools.",
        "output": "Top 3 YouTube research report with recommendation.",
        "flow": ["0-10 Hook: Can AI use tools?", "10-25 Explain chatbot vs agent", "25-45 Demo research workflow", "45-75 Build report", "75-90 Verify and discuss"],
        "interactive": ["Sort cards: chatbot or agent?", "Topic vote: choose research topic", "Verification checkpoint: what should humans check?"],
        "prompt": "Act as a careful research assistant. Topic: [student topic]. Compare three beginner-friendly YouTube videos. Summarize each and recommend the best first video.",
        "homework": "Create a Top 3 video comparison report. Replace or adjust this homework for your class.",
    },
    {
        "week": 5,
        "title": "Browser Use + Computer Use",
        "subtitle": "AI as an operator with safety boundaries",
        "color": RED,
        "goal": "Students design safe browser/computer tasks and understand human approval checkpoints.",
        "output": "A safe five-step automation task design.",
        "flow": ["0-10 Hook: AI operating a browser", "10-25 Safety rules", "25-45 Demo public-info task", "45-75 Design task", "75-90 Risk review"],
        "interactive": ["Red/yellow/green: is this task safe?", "Group design: 5-step workflow", "Stop button practice: when to interrupt AI"],
        "prompt": "Design a safe browser task for an AI agent. Use public information only. Do not log in, buy anything, or submit real forms. Include a human approval checkpoint.",
        "homework": "Write a safe five-step browser task an AI agent could perform. Change this homework as needed.",
    },
    {
        "week": 6,
        "title": "Final AI Workflow Demo",
        "subtitle": "Present, reflect, and choose next build",
        "color": BLUE,
        "goal": "Students connect course ideas into a final project concept, demo, or workflow presentation.",
        "output": "Final demo or project pitch.",
        "flow": ["0-10 Hook: what could you build next?", "10-25 Demo presentation model", "25-60 Build/polish", "60-85 Presentations", "85-90 Feedback and next step"],
        "interactive": ["Gallery walk: leave one note", "Presentation timer: 2 minutes each", "Next-build vote: what would you add?"],
        "prompt": "Help me prepare a 2-minute project presentation. Include: what I built, who it helps, best prompt, what broke, how I fixed it, and what I would add next.",
        "homework": "Optional: revise the final project after feedback. You can replace this with a final reflection or extension task.",
    },
]


def fill_shape(shape, color, line=None):
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.color.rgb = line or color


def fill_background(slide, color):
    bg = slide.background.fill
    bg.solid()
    bg.fore_color.rgb = color


def rect(slide, x, y, w, h, color, line=None, radius=True):
    kind = MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE if radius else MSO_AUTO_SHAPE_TYPE.RECTANGLE
    shape = slide.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    fill_shape(shape, color, line)
    return shape


def text(slide, x, y, w, h, value, size=14, bold=False, color=INK, align=None):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.TOP
    p = tf.paragraphs[0]
    p.text = value
    if align:
        p.alignment = align
    run = p.runs[0]
    run.font.name = "Aptos"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    return box


def bullet_list(slide, x, y, w, h, items, size=11.5):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = item
        p.font.name = "Aptos"
        p.font.size = Pt(size)
        p.font.color.rgb = INK
        p.space_after = Pt(5)
    return box


def card(slide, x, y, w, h, title, body, accent, bg=WHITE):
    rect(slide, x, y, w, h, bg, LINE)
    rect(slide, x, y, 0.08, h, accent, accent, radius=False)
    text(slide, x + 0.18, y + 0.13, w - 0.34, 0.24, title, 13, True, accent)
    if isinstance(body, list):
        bullet_list(slide, x + 0.22, y + 0.47, w - 0.42, h - 0.55, body, 11.4)
    else:
        text(slide, x + 0.18, y + 0.45, w - 0.34, h - 0.50, body, 12.2)


def add_lesson_slide(slide, item):
    color = item["color"]
    fill_background(slide, LIGHT)
    rect(slide, 0, 0, 13.333, 0.72, color, color, radius=False)
    text(slide, 0.55, 0.18, 1.35, 0.22, f"WEEK {item['week']}", 11, True, WHITE)
    text(slide, 1.7, 0.13, 6.9, 0.35, item["title"], 20, True, WHITE)
    text(slide, 8.65, 0.2, 4.05, 0.2, item["subtitle"], 10.5, False, WHITE, PP_ALIGN.RIGHT)

    card(slide, 0.45, 0.96, 3.0, 1.15, "Goal", item["goal"], color)
    card(slide, 3.65, 0.96, 2.55, 1.15, "Student output", item["output"], GREEN)
    rect(slide, 6.4, 0.96, 6.45, 1.15, YELLOW_BG, RGBColor(253, 186, 116))
    text(slide, 6.58, 1.08, 5.9, 0.22, "Homework - editable", 13, True, ORANGE)
    text(slide, 6.58, 1.40, 5.95, 0.55, item["homework"], 12.2)

    card(slide, 0.45, 2.34, 3.85, 2.03, "90-minute flow", item["flow"], color)
    card(slide, 4.55, 2.34, 3.8, 2.03, "Interactive moments", item["interactive"], PURPLE, PURPLE_BG)
    card(slide, 8.6, 2.34, 4.25, 2.03, "Tutor moves", [
        "Ask students to predict before AI answers.",
        "Pause after every demo to name what changed.",
        "Make one small improvement, then test.",
    ], ORANGE, BLUE_BG)

    text(slide, 0.55, 4.65, 2.2, 0.22, "Live demo prompt", 15, True, DARK)
    rect(slide, 0.45, 4.95, 8.0, 1.5, DARK, DARK)
    text(slide, 0.65, 5.15, 7.58, 1.05, item["prompt"], 11.4, False, WHITE)
    card(slide, 8.72, 4.95, 4.13, 1.5, "Exit ticket", "Before students leave, ask: What did you make? What was one prompt that helped? What is one thing you would change next?", GREEN, GREEN_BG)

    text(slide, 0.45, 7.15, 6.4, 0.17, "AI Builders Academy | Separate weekly lesson slide", 8.5, False, MUTED)
    text(slide, 12.05, 7.15, 0.85, 0.17, f"Week {item['week']}", 8.5, False, MUTED, PP_ALIGN.RIGHT)


def build_deck(item):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    add_lesson_slide(slide, item)
    path = OUT_DIR / f"Week_{item['week']}_AI_Builders_{item['title'].replace(' ', '_').replace('+', 'and')}.pptx"
    prs.save(path)
    return path


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for lesson in LESSONS:
        print(build_deck(lesson))


if __name__ == "__main__":
    main()
