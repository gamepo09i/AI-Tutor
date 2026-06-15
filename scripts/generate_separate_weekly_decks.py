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
        "title": "AI Teammate Foundations",
        "subtitle": "ChatGPT, Prompting, and First Assistant",
        "color": BLUE,
        "curriculum": "Students learn that AI is a creative teammate, not just a search box. They practice prompt structure, test answers, and build a personal assistant they can improve.",
        "tools": ["ChatGPT", "Prompt Engineering", "Role + Task + Context + Format", "Critique"],
        "goal": "Move students from simply asking questions to directing an AI teammate with clear roles, context, and output formats.",
        "deliverable": "Personal AI assistant with 5 useful example conversations.",
        "hook": "Can AI do my homework?",
        "big_idea": "The student is the director. AI is the teammate.",
        "concepts": ["AI can take a role", "Specific context improves answers", "A good format makes answers easier to use", "Humans still check the result"],
        "demo_title": "Bad prompt vs. good prompt",
        "demo": ["Bad: Help me study.", "Better: Act as a friendly biology tutor. Quiz me on cells one question at a time. If I miss one, explain it simply."],
        "build": ["Choose an assistant type: Study Buddy, Minecraft Coach, Music Advisor, Sports Trainer, Writing Helper, or Game Idea Coach.", "Create the assistant's role, rules, and answer format.", "Test it with 3 real questions.", "Improve the prompt after seeing the answers."],
        "interactive": ["Thumb vote: helpful or lazy?", "Pair test: ask your partner's assistant one question.", "Repair challenge: turn a vague prompt into a clear prompt."],
        "timing": ["Hook + discussion: 10 min", "Prompt formula demo: 20 min", "Assistant build: 40 min", "Pair test + revise: 15 min", "Exit ticket: 5 min"],
        "prompt": "Act as a friendly study coach for a high school student.\nHelp me study [topic]. Ask one question at a time. If I get it wrong, explain simply and let me try again.",
        "homework": "Create an AI assistant with 5 example conversations. Add one rule that makes the assistant more helpful or safer.",
    },
    {
        "week": 2,
        "title": "Codex Build Studio",
        "subtitle": "VS Code, Web UI, and Space Invaders Basics",
        "color": GREEN,
        "curriculum": "Students set up the coding workspace, learn files and previews, build a simple web UI, then use Codex to create the basics of a Space Invaders-style browser game.",
        "tools": ["VS Code", "Codex", "HTML", "CSS", "JavaScript", "Browser Preview"],
        "goal": "Help students get a working coding setup and experience the loop of ask, run, debug, and improve.",
        "deliverable": "Personal web page plus a playable Space Invaders starter.",
        "hook": "Can AI become your coding partner inside VS Code?",
        "big_idea": "A project is a folder of files that you can run, inspect, and improve one change at a time.",
        "concepts": ["Open a project folder", "Create and edit files", "Preview a web page", "Describe bugs clearly", "Keep games tiny before adding features"],
        "demo_title": "From web page to tiny game",
        "demo": ["Open VS Code and a course folder.", "Ask Codex to create index.html.", "Preview the page.", "Ask Codex to explain the file.", "Create a tiny canvas game loop."],
        "build": ["Create an AI Builder profile page.", "Customize one section.", "Create space-invaders.html.", "Add player movement, bullets, enemies, score, and game over.", "Playtest and fix one bug."],
        "interactive": ["Setup checkpoint: everyone shows the project folder.", "Choice board: pick one page section to customize.", "Final build sprint: make a basic Space Invaders game with Codex help."],
        "timing": ["Setup + folder tour: 15 min", "Profile page demo: 20 min", "Student customization: 20 min", "Space Invaders starter: 30 min", "Playtest + exit ticket: 5 min"],
        "prompt": "Create a beginner-friendly Space Invaders-style browser game in one file named space-invaders.html. Use HTML, CSS, and JavaScript. Include a player ship, left/right movement, shooting, falling enemies, score, lives, restart button, and comments explaining the main code.",
        "homework": "Expand the Space Invaders starter: add one new enemy type, power-up, sound effect, level, or visual polish. Write 3 sentences explaining what changed.",
    },
    {
        "week": 3,
        "title": "AI Agents and Research Workflows",
        "subtitle": "MCP, YouTube Research, and Browser Safety",
        "color": PURPLE,
        "curriculum": "Students learn how agents plan steps, use tools, and need human supervision. MCP, YouTube research, browser use, and computer use are combined into one practical workflow lesson.",
        "tools": ["AI Agent", "MCP", "YouTube Research", "Browser Use", "Verification", "Safety Rules"],
        "goal": "Help students understand tool-using AI by designing a safe research workflow with clear human checkpoints.",
        "deliverable": "Top 3 video comparison report plus a safe browser task design.",
        "hook": "Can AI use internet tools by itself?",
        "big_idea": "An agent is AI plus a job plus tools.",
        "concepts": ["Chatbots answer inside a conversation", "Agents plan steps and use tools", "MCP connects AI to external tools", "Research needs verification", "Browser tasks need boundaries"],
        "demo_title": "Research agent workflow",
        "demo": ["Choose a student-selected topic.", "Find or simulate three video results.", "Summarize each result.", "Compare which video is best for beginners.", "Mark what a human must verify."],
        "build": ["Pick a topic.", "Create a Top 3 video comparison.", "Write why each video is useful.", "Design a five-step browser task.", "Add one human approval checkpoint."],
        "interactive": ["Sort cards: chatbot or agent?", "Red/yellow/green: is this browser task safe?", "Verification checkpoint: what should humans check?"],
        "timing": ["Agent vs chatbot hook: 10 min", "MCP + tool demo: 20 min", "YouTube research build: 30 min", "Browser safety design: 20 min", "Share-out: 10 min"],
        "prompt": "Act as a careful research assistant. Topic: [student topic]. Compare three beginner-friendly YouTube videos. Summarize each, recommend the best first video, and list what a human should verify before trusting the report.",
        "homework": "Create a Top 3 video comparison report and add a safe 5-step browser task that could extend the research.",
    },
    {
        "week": 4,
        "title": "Final Agent Project",
        "subtitle": "OpenClaw, OAuth, Demo, and Next Build",
        "color": BLUE,
        "curriculum": "Students connect prompting, coding, games, agents, MCP, browser safety, OpenClaw, and OAuth into a final project concept or demo they can explain.",
        "tools": ["OpenClaw", "OAuth", "Codex OAuth", "AI Agent Workflow", "Presentation", "Peer Feedback"],
        "goal": "Help students turn the course pieces into a final project demo, workflow diagram, or next-build plan.",
        "deliverable": "Final AI project concept, demo, or 2-minute presentation.",
        "hook": "Can we connect AI tools like apps on a phone?",
        "big_idea": "A strong AI project is useful, safe, testable, and explainable.",
        "concepts": ["Agent frameworks organize workflows", "OAuth grants limited permission", "Connected tools need trust and safety", "Final projects should be explainable", "Good demos show one working moment"],
        "demo_title": "Agent workflow concept",
        "demo": ["Show a simple OpenClaw-style workflow.", "Explain where login/permission fits.", "Connect the idea back to Codex and tools.", "Model a 2-minute final presentation.", "Show how to ask AI for a next-build checklist."],
        "build": ["Choose final project idea.", "Define user, problem, tools, and workflow.", "Prepare a short demo or concept pitch.", "Give peer feedback.", "Write the next 3 improvements."],
        "interactive": ["App connection analogy: what permissions would this need?", "Gallery walk: leave one feedback note.", "Next-build vote: what should this project add next?"],
        "timing": ["OpenClaw + OAuth concept: 20 min", "Project planning: 25 min", "Demo/pitch build: 25 min", "Gallery walk: 15 min", "Reflection: 5 min"],
        "prompt": "Help me prepare a 2-minute project presentation. Include: what I built, who it helps, best prompt, what broke, how I fixed it, and what I would add next.",
        "homework": "Refine the final project after feedback. Add one improvement to the game, assistant, research workflow, or agent plan and prepare a short demo.",
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

    slide = add_slide(prs, item, 2, total, "Today’s target", "Each lesson is planned for about 90 minutes.")
    card(slide, 0.75, 1.92, 3.65, 1.42, "Main goal", item["goal"], item["color"], WHITE)
    card(slide, 4.65, 1.92, 3.55, 1.42, "Student deliverable", item["deliverable"], GREEN, GREEN_SOFT)
    card(slide, 8.45, 1.92, 3.95, 1.42, "Tools today", item["tools"], PURPLE, PURPLE_SOFT)
    card(slide, 1.0, 3.85, 11.25, 1.8, "90-minute flow", item["timing"], ORANGE, YELLOW)

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
