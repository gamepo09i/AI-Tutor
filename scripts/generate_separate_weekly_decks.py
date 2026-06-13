from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
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
YELLOW = RGBColor(255, 247, 237)
BLUE_SOFT = RGBColor(239, 246, 255)
GREEN_SOFT = RGBColor(236, 253, 245)
PURPLE_SOFT = RGBColor(245, 243, 255)


LESSONS = [
    {
        "week": 1,
        "title": "Meet Your AI Teammate",
        "subtitle": "ChatGPT Basics",
        "color": BLUE,
        "curriculum": "ChatGPT is not just a search engine. It can act as a brainstorming partner, tutor, coach, writer, and project assistant.",
        "tools": ["ChatGPT", "Prompt Engineering", "Role + Task + Context + Format"],
        "goal": "Move students from simply asking AI questions to designing a useful personal assistant.",
        "deliverable": "AI Assistant with 5 example conversations.",
        "hook": "Can AI do my homework?",
        "big_idea": "The student is the director. AI is the teammate.",
        "concepts": ["AI can take a role", "Specific context improves answers", "A good format makes answers easier to use"],
        "demo_title": "Bad prompt vs. good prompt",
        "demo": ["Bad: Help me study.", "Better: Act as a friendly biology tutor. Quiz me on cells one question at a time. If I miss one, explain it simply."],
        "build": ["Choose an assistant type: Study Buddy, Minecraft Coach, Music Advisor, Sports Trainer, or Writing Helper.", "Create the assistant's role and rules.", "Test it with 3 questions.", "Improve the prompt after seeing the answers."],
        "interactive": ["Thumb vote: helpful or lazy?", "Pair test: ask your partner's assistant one question.", "Repair challenge: turn a vague prompt into a clear prompt."],
        "prompt": "Act as a friendly study coach for a high school student.\nHelp me study [topic]. Ask one question at a time. If I get it wrong, explain simply and let me try again.",
        "homework": "Create an AI assistant with 5 example conversations. Edit this homework as needed.",
    },
    {
        "week": 2,
        "title": "AI Coding Environment",
        "subtitle": "VS Code + Codex Setup with Web UI Example",
        "color": GREEN,
        "curriculum": "Students install or open VS Code, set up Codex, learn the UI, understand folders and files, practice token management, and use Codex to create a simple web UI.",
        "tools": ["VS Code", "Codex", "Folders", "Files", "Token Management", "HTML", "CSS"],
        "goal": "Help students get a working coding setup and see a visible web page quickly.",
        "deliverable": "Ready-to-use AI coding environment + simple web UI.",
        "hook": "Can AI become your coding partner inside VS Code?",
        "big_idea": "A project is just a folder with files that work together.",
        "concepts": ["Open a project folder", "Create and edit files", "Preview a web page", "Ask Codex for small changes"],
        "demo_title": "Folder to first web page",
        "demo": ["Open VS Code.", "Open a course folder.", "Ask Codex to create index.html.", "Preview the page.", "Ask Codex to explain the file."],
        "build": ["Create an AI Builder profile page.", "Add name, favorite apps, and project ideas.", "Customize one color or section.", "Preview and explain what changed."],
        "interactive": ["Setup checkpoint: everyone shows the project folder.", "Choice board: pick one section to customize.", "Explain-back: what does index.html do?"],
        "prompt": "Create a beginner-friendly personal web page in one file named index.html. Theme: AI Builder profile. Include name, favorite apps, project ideas, and one button.",
        "homework": "Complete setup and customize the web UI with one personal section. Edit this homework as needed.",
    },
    {
        "week": 3,
        "title": "AI Game Studio",
        "subtitle": "Build a Game with Codex",
        "color": ORANGE,
        "curriculum": "Students use the VS Code + Codex environment prepared in Class 2 to build a simple browser game. Options include a clicker game, dodge game, snake-style game, or simple space shooter.",
        "tools": ["Codex", "HTML", "CSS", "JavaScript", "Game Development"],
        "goal": "Use Codex to generate, explain, debug, and improve a small browser game step by step.",
        "deliverable": "Playable HTML/JavaScript game.",
        "hook": "Can AI build a game that your friends can play?",
        "big_idea": "Games teach debugging because problems are easy to see.",
        "concepts": ["Start with a tiny playable version", "Test before adding features", "Describe bugs clearly", "Improve one feature at a time"],
        "demo_title": "Build a tiny game live",
        "demo": ["Ask Codex for a one-file clicker game.", "Run the game.", "Find one thing to improve.", "Ask Codex to add score, timer, or levels."],
        "build": ["Pick a game type: clicker, dodge, snake-style, or space shooter.", "Generate the first version.", "Playtest with a partner.", "Add one improvement."],
        "interactive": ["Game menu vote: which game should we build?", "Bug hunt: what broke or feels weird?", "Peer playtest: one compliment, one upgrade idea."],
        "prompt": "Create a simple browser clicker game in one HTML file. Add score, 30-second timer, game over message, clean styling, and comments explaining the JavaScript.",
        "homework": "Add one new feature and get feedback from a parent or friend. Edit this homework as needed.",
    },
    {
        "week": 4,
        "title": "AI Agents & MCP",
        "subtitle": "YouTube Research Agent",
        "color": PURPLE,
        "curriculum": "Students learn the concept of an AI agent and MCP. MCP is introduced as a connector that lets AI use external tools. Students build or simulate a YouTube Research Agent.",
        "tools": ["AI Agent", "MCP", "YouTube Agent", "Research", "Verification"],
        "goal": "Help students understand tool-using AI through a practical YouTube research workflow.",
        "deliverable": "AI YouTube Research Report.",
        "hook": "Can AI use internet tools by itself?",
        "big_idea": "An agent is AI plus a job plus tools.",
        "concepts": ["Chatbots answer inside a conversation", "Agents plan steps and use tools", "MCP connects AI to external tools", "Research needs verification"],
        "demo_title": "YouTube research workflow",
        "demo": ["Choose a student-selected topic.", "Find or simulate three video results.", "Summarize each result.", "Compare which video is best for beginners."],
        "build": ["Pick a topic.", "Create a Top 3 video comparison.", "Write why each video is useful.", "Recommend the best first video."],
        "interactive": ["Sort cards: chatbot or agent?", "Topic vote: choose the class research topic.", "Verification checkpoint: what should humans check?"],
        "prompt": "Act as a careful research assistant. Topic: [student topic]. Compare three beginner-friendly YouTube videos. Summarize each and recommend the best first video.",
        "homework": "Create a Top 3 video comparison report. Edit this homework as needed.",
    },
    {
        "week": 5,
        "title": "Browser Use & Computer Use",
        "subtitle": "AI as an Operator",
        "color": RED,
        "curriculum": "Students learn how AI can use a browser or computer-like interface to perform tasks. The class focuses on safe, understandable demos: opening pages, extracting information, filling simple forms, comparing results, and explaining what the agent is doing.",
        "tools": ["Browser Use", "Computer Use", "AI Operator", "Safety Rules"],
        "goal": "Teach students to design safe AI operator tasks with clear human approval points.",
        "deliverable": "Browser/Computer task automation demo project or task design.",
        "hook": "Can AI operate a browser like a human assistant?",
        "big_idea": "AI operators need boundaries, supervision, and safe task design.",
        "concepts": ["Use public information", "Avoid private accounts and purchases", "Watch what the agent does", "Add human approval checkpoints"],
        "demo_title": "Safe browser task",
        "demo": ["Open a public page.", "Extract simple information.", "Compare results.", "Explain every step before moving on."],
        "build": ["Design a five-step browser task.", "Mark safe and risky steps.", "Add a human approval checkpoint.", "Share the task with a partner."],
        "interactive": ["Red/yellow/green: is this task safe?", "Group design: make a five-step workflow.", "Stop-button practice: when should a tutor interrupt AI?"],
        "prompt": "Design a safe browser task for an AI agent. Use public information only. Do not log in, buy anything, or submit real forms. Include a human approval checkpoint.",
        "homework": "Design a 5-step browser task that an AI agent could perform. Edit this homework as needed.",
    },
    {
        "week": 6,
        "title": "OpenClaw + Codex OAuth",
        "subtitle": "Building an AI Agent Workflow",
        "color": BLUE,
        "curriculum": "Students are introduced to OpenClaw as a more advanced AI agent framework. The class explains OAuth in simple terms, shows how Codex OAuth connects tools securely, and demonstrates an OpenClaw-style workflow.",
        "tools": ["OpenClaw", "OAuth", "Codex OAuth", "AI Agent Workflow", "Presentation"],
        "goal": "Help students connect the course concepts into a final AI agent workflow or project demo.",
        "deliverable": "OpenClaw-based final AI agent concept or demo presentation.",
        "hook": "Can we connect AI tools like apps on a phone?",
        "big_idea": "OAuth is a permission system: it lets tools connect without sharing passwords directly.",
        "concepts": ["Agent frameworks organize workflows", "OAuth grants limited permission", "Connected tools need trust and safety", "Final projects should be explainable"],
        "demo_title": "Agent workflow concept",
        "demo": ["Show a simple OpenClaw-style workflow.", "Explain where login/permission fits.", "Connect the idea back to Codex and tools.", "Model a 2-minute final presentation."],
        "build": ["Choose final project idea.", "Define user, problem, tools, and workflow.", "Prepare a short demo or concept pitch.", "Give peer feedback."],
        "interactive": ["App connection analogy: what permissions would this need?", "Gallery walk: leave one feedback note.", "Next-build vote: what should this project add next?"],
        "prompt": "Help me prepare a 2-minute project presentation. Include: what I built, who it helps, best prompt, what broke, how I fixed it, and what I would add next.",
        "homework": "Final project concept or demo presentation. Edit this homework as needed.",
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
    shape_type = MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE if radius else MSO_AUTO_SHAPE_TYPE.RECTANGLE
    shape = slide.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    fill_shape(shape, color, line)
    return shape


def text(slide, x, y, w, h, value, size=18, bold=False, color=INK, align=None):
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


def bullets(slide, x, y, w, h, items, size=22, color=INK):
    box = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.clear()
    tf.word_wrap = True
    for idx, item in enumerate(items):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = item
        p.font.name = "Aptos"
        p.font.size = Pt(size)
        p.font.color.rgb = color
        p.space_after = Pt(9)
    return box


def footer(slide, item, slide_no, total):
    text(slide, 0.45, 7.13, 5.8, 0.18, f"AI Builders Academy | Week {item['week']}", 8.5, False, MUTED)
    text(slide, 12.0, 7.13, 0.9, 0.18, f"{slide_no}/{total}", 8.5, False, MUTED, PP_ALIGN.RIGHT)


def title_bar(slide, item):
    rect(slide, 0, 0, 13.333, 0.72, item["color"], item["color"], radius=False)
    text(slide, 0.55, 0.18, 1.25, 0.22, f"WEEK {item['week']}", 11, True, WHITE)
    text(slide, 1.72, 0.12, 7.1, 0.35, item["title"], 20, True, WHITE)
    text(slide, 8.65, 0.19, 4.1, 0.2, item["subtitle"], 10.5, False, WHITE, PP_ALIGN.RIGHT)


def card(slide, x, y, w, h, title, body, accent=BLUE, bg=WHITE):
    rect(slide, x, y, w, h, bg, LINE)
    rect(slide, x, y, 0.08, h, accent, accent, radius=False)
    text(slide, x + 0.18, y + 0.14, w - 0.35, 0.24, title, 13, True, accent)
    if isinstance(body, list):
        bullets(slide, x + 0.22, y + 0.48, w - 0.42, h - 0.55, body, 14)
    else:
        text(slide, x + 0.18, y + 0.48, w - 0.36, h - 0.55, body, 14)


def prompt_box(slide, value):
    rect(slide, 0.9, 2.05, 11.55, 3.05, DARK, DARK)
    text(slide, 1.18, 2.32, 11.0, 2.45, value, 21, False, WHITE)


def add_slide(prs, item, slide_no, total, title, subtitle=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fill_background(slide, LIGHT)
    title_bar(slide, item)
    text(slide, 0.72, 1.02, 11.7, 0.52, title, 31, True, INK)
    if subtitle:
        text(slide, 0.76, 1.62, 11.4, 0.28, subtitle, 15, False, MUTED)
    footer(slide, item, slide_no, total)
    return slide


def build_week_deck(item):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    total = 9

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    fill_background(slide, DARK)
    rect(slide, 8.9, -0.5, 3.2, 3.2, item["color"], item["color"])
    rect(slide, 0.75, 0.78, 1.7, 0.38, item["color"], item["color"])
    text(slide, 0.92, 0.86, 1.35, 0.18, f"WEEK {item['week']}", 11, True, WHITE, PP_ALIGN.CENTER)
    text(slide, 0.75, 1.58, 9.7, 0.82, item["title"], 44, True, WHITE)
    text(slide, 0.78, 2.58, 9.6, 0.5, item["subtitle"], 22, False, RGBColor(226, 232, 240))
    text(slide, 0.8, 4.45, 9.8, 0.7, item["curriculum"], 20, False, RGBColor(226, 232, 240))
    footer(slide, item, 1, total)

    slide = add_slide(prs, item, 2, total, "Today’s target", "Taken straight from the curriculum plan.")
    card(slide, 0.85, 2.05, 3.7, 1.65, "Main goal", item["goal"], item["color"], WHITE)
    card(slide, 4.8, 2.05, 3.4, 1.65, "Student deliverable", item["deliverable"], GREEN, GREEN_SOFT)
    card(slide, 8.45, 2.05, 3.95, 1.65, "Tools today", item["tools"], PURPLE, PURPLE_SOFT)
    card(slide, 1.15, 4.45, 10.8, 1.0, "Tutor framing", "Every tool is introduced as a practical building tool, not abstract theory.", ORANGE, YELLOW)

    slide = add_slide(prs, item, 3, total, "Hook", "Start with a question that makes students curious immediately.")
    text(slide, 0.95, 2.05, 11.4, 1.05, item["hook"], 48, True, item["color"], PP_ALIGN.CENTER)
    bullets(slide, 1.35, 4.0, 10.6, 1.35, ["Ask for quick predictions.", "Let students argue both sides.", "Then show the demo before explaining too much."], 23)

    slide = add_slide(prs, item, 4, total, "Big idea", "One concept students should remember after class.")
    text(slide, 1.0, 2.0, 11.3, 0.75, item["big_idea"], 34, True, INK, PP_ALIGN.CENTER)
    card(slide, 1.1, 3.35, 11.1, 1.8, "Key concepts", item["concepts"], item["color"], WHITE)

    slide = add_slide(prs, item, 5, total, item["demo_title"], "Live demo: students see the result before the explanation.")
    card(slide, 0.95, 2.0, 5.25, 3.4, "Demo steps", item["demo"], item["color"], WHITE)
    card(slide, 6.55, 2.0, 5.75, 3.4, "Tutor move", ["Narrate what you are asking AI to do.", "Pause after the output.", "Ask: What should we improve next?"], ORANGE, YELLOW)

    slide = add_slide(prs, item, 6, total, "Live prompt", "Copy, modify, and use during the demo.")
    prompt_box(slide, item["prompt"])
    text(slide, 1.05, 5.55, 10.8, 0.35, "After the response, ask students: What worked? What is missing? What should we ask next?", 16, True, item["color"], PP_ALIGN.CENTER)

    slide = add_slide(prs, item, 7, total, "Hands-on build", "Students create their own version and improve through iteration.")
    card(slide, 0.95, 2.0, 5.45, 3.55, "Build steps", item["build"], GREEN, GREEN_SOFT)
    card(slide, 6.75, 2.0, 5.45, 3.55, "Keep it manageable", ["Start with the smallest working version.", "Let students customize one thing.", "Have them explain the change in plain English."], item["color"], WHITE)

    slide = add_slide(prs, item, 8, total, "Interactive checkpoints", "Use these to keep kids active instead of just watching.")
    card(slide, 1.0, 2.0, 11.2, 2.3, "Student interaction", item["interactive"], PURPLE, PURPLE_SOFT)
    card(slide, 1.0, 4.75, 11.2, 0.95, "Exit ticket", "What did you make? What prompt helped? What would you change next?", GREEN, GREEN_SOFT)

    slide = add_slide(prs, item, 9, total, "Homework", "Editable: change this based on what you want students to do.")
    rect(slide, 0.95, 2.0, 11.45, 2.0, YELLOW, RGBColor(253, 186, 116))
    text(slide, 1.25, 2.35, 10.8, 0.55, item["homework"], 25, True, INK, PP_ALIGN.CENTER)
    card(slide, 1.0, 4.65, 11.2, 1.0, "Assessment reminder", "Assessment is based on projects, not tests. Students should be able to show what they made and explain one improvement.", ORANGE, WHITE)

    path = OUT_DIR / f"Week_{item['week']}_AI_Builders_{item['title'].replace(' ', '_').replace('&', 'and').replace('+', 'and')}.pptx"
    prs.save(path)
    return path


def main():
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for item in LESSONS:
        print(build_week_deck(item))


if __name__ == "__main__":
    main()
