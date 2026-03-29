from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import datetime

doc = Document()

# ── Page margins ────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin   = Inches(1.1)
    section.right_margin  = Inches(1.1)

# ── Styles ───────────────────────────────────────────────────
style_normal = doc.styles['Normal']
style_normal.font.name = 'Calibri'
style_normal.font.size = Pt(11)

def set_heading(para, text, level=1, color=None):
    para.text = text
    para.style = f'Heading {level}'
    run = para.runs[0]
    run.font.name = 'Calibri'
    if color:
        run.font.color.rgb = RGBColor(*color)

def add_heading(doc, text, level=1, color=(26, 50, 96)):
    p = doc.add_heading('', level=level)
    run = p.add_run(text)
    run.font.name = 'Calibri'
    run.font.color.rgb = RGBColor(*color)
    return p

def add_para(doc, text='', bold_prefix=None, size=11):
    p = doc.add_paragraph()
    if bold_prefix:
        run = p.add_run(bold_prefix)
        run.bold = True
        run.font.size = Pt(size)
        run.font.name = 'Calibri'
    if text:
        run = p.add_run(text)
        run.font.size = Pt(size)
        run.font.name = 'Calibri'
    return p

def add_bullet(doc, text, bold_prefix=None):
    p = doc.add_paragraph(style='List Bullet')
    if bold_prefix:
        r = p.add_run(bold_prefix)
        r.bold = True
        r.font.name = 'Calibri'
        r.font.size = Pt(11)
    r = p.add_run(text)
    r.font.name = 'Calibri'
    r.font.size = Pt(11)
    return p

def add_table(doc, headers, rows, col_widths):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Table Grid'
    # Header row
    hdr = table.rows[0]
    for i, h in enumerate(headers):
        cell = hdr.cells[i]
        cell.width = Inches(col_widths[i])
        p = cell.paragraphs[0]
        run = p.add_run(h)
        run.bold = True
        run.font.name = 'Calibri'
        run.font.size = Pt(10)
        # Header shading
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), '1A3260')
        tcPr.append(shd)
        run.font.color.rgb = RGBColor(255, 255, 255)
    # Data rows
    for ri, row_data in enumerate(rows):
        row = table.rows[ri + 1]
        fill = 'EEF5FF' if ri % 2 == 0 else 'FFFFFF'
        for ci, cell_text in enumerate(row_data):
            cell = row.cells[ci]
            cell.width = Inches(col_widths[ci])
            p = cell.paragraphs[0]
            run = p.add_run(cell_text)
            run.font.name = 'Calibri'
            run.font.size = Pt(10)
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), fill)
            tcPr.append(shd)
    doc.add_paragraph()
    return table

# ═══════════════════════════════════════════════════════════
# TITLE PAGE
# ═══════════════════════════════════════════════════════════
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('AralMate MVP')
run.font.name = 'Calibri'
run.font.size = Pt(28)
run.font.bold = True
run.font.color.rgb = RGBColor(26, 50, 96)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Build Review Document')
run.font.name = 'Calibri'
run.font.size = Pt(16)
run.font.color.rgb = RGBColor(107, 132, 168)

doc.add_paragraph()

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Version 1.0  |  March 2026')
run.font.name = 'Calibri'
run.font.size = Pt(11)
run.font.color.rgb = RGBColor(107, 132, 168)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run('Prepared for Tech Lead Review')
run.font.name = 'Calibri'
run.font.size = Pt(11)
run.font.color.rgb = RGBColor(107, 132, 168)

doc.add_page_break()

# ═══════════════════════════════════════════════════════════
# 1. OVERVIEW
# ═══════════════════════════════════════════════════════════
add_heading(doc, '1. Project Overview', level=1)
add_para(doc, (
    'AralMate is a mobile web application designed for Filipino Grade 4 elementary school students. '
    'The student takes a photo of their school material (worksheet, textbook page, or handwritten notes), '
    'and Alon — the AI tutor character — analyzes the image and automatically generates a structured lesson, '
    'a 5-question multiple-choice quiz, 8 flashcards, and a follow-up chat interface.'
))
doc.add_paragraph()
add_para(doc, 'Primary user: ', bold_prefix='Primary user: ')
doc.paragraphs[-1].runs[-1].text = ''
p = doc.paragraphs[-1]
r = p.add_run('Miguel Clarin, Grade 4 student')
r.font.name = 'Calibri'
r.font.size = Pt(11)

add_para(doc, '', bold_prefix='Built by: ')
p = doc.paragraphs[-1]
r = p.add_run('Michael Clarin (maykel.clarin@gmail.com)')
r.font.name = 'Calibri'
r.font.size = Pt(11)

add_para(doc, '', bold_prefix='Status: ')
p = doc.paragraphs[-1]
r = p.add_run('MVP deployed and working end-to-end as of March 2026')
r.font.name = 'Calibri'
r.font.size = Pt(11)

doc.add_paragraph()

# ═══════════════════════════════════════════════════════════
# 2. ARCHITECTURE
# ═══════════════════════════════════════════════════════════
add_heading(doc, '2. Architecture', level=1)
add_para(doc, (
    'The system uses a three-tier architecture to serve the app securely without exposing API keys in the browser.'
))
doc.add_paragraph()

add_table(doc,
    ['Layer', 'Component', 'Platform', 'Cost'],
    [
        ['Frontend',    'aralmate.html — single-file responsive web app', 'Netlify (static hosting)', 'Free'],
        ['API Proxy',   'worker.js — validates requests, forwards to Kimi', 'Cloudflare Workers',       'Free'],
        ['AI Backend',  'Kimi K2.5 vision model (moonshot-v1-8k-vision-preview)', 'Moonshot AI API',   'Per token'],
    ],
    [1.1, 2.8, 1.5, 0.9]
)

add_heading(doc, 'Request Flow', level=2)
add_para(doc, "User's phone (Chrome browser)")
add_bullet(doc, 'User uploads photo → app converts to base64')
add_bullet(doc, 'POST to Cloudflare Worker with X-AralMate-Secret header')
add_bullet(doc, 'Worker validates secret, attaches KIMI_API_KEY, forwards to Moonshot AI')
add_bullet(doc, 'Kimi returns structured JSON (lesson + quiz + flashcards)')
add_bullet(doc, 'Worker returns response to browser with CORS headers')
add_bullet(doc, 'App renders lesson, quiz, and flashcard tabs')
doc.add_paragraph()

# ═══════════════════════════════════════════════════════════
# 3. LIVE URLS
# ═══════════════════════════════════════════════════════════
add_heading(doc, '3. Deployment URLs', level=1)

add_table(doc,
    ['Resource', 'URL'],
    [
        ['Live App (Netlify)',       'https://fastidious-trifle-8e650c.netlify.app'],
        ['Cloudflare Worker',        'https://aralmate-proxy.maykel-clarin.workers.dev'],
        ['Kimi API Platform',        'platform.moonshot.ai (international account)'],
    ],
    [1.8, 4.5]
)

# ═══════════════════════════════════════════════════════════
# 4. COMPONENTS BUILT
# ═══════════════════════════════════════════════════════════
add_heading(doc, '4. Components Built', level=1)

add_heading(doc, '4.1 Frontend — aralmate.html', level=2)
add_para(doc, (
    'A single self-contained HTML/CSS/JavaScript file. No framework, no build step, no dependencies. '
    'Optimized for mobile (max-width 430px). Key screens and features:'
))
add_bullet(doc, 'Home screen with camera capture and gallery upload buttons')
add_bullet(doc, 'Image preview with quality check banner')
add_bullet(doc, '5-step processing animation while Kimi reads the material')
add_bullet(doc, 'Lesson tab — overview, key points, "Remember This" box, and key terms')
add_bullet(doc, 'Quiz tab — 5 multiple-choice questions with instant feedback and score')
add_bullet(doc, 'Flashcards tab — 8 flip cards with navigation and progress bar')
add_bullet(doc, 'Simplify toggle — switches lesson to easier language for struggling readers')
add_bullet(doc, 'Chat bar — student can ask Alon follow-up questions about the lesson')
add_bullet(doc, 'Tala points gamification system (earned by completing quizzes and flashcards)')
add_bullet(doc, 'Session persistence via localStorage (lesson survives page refresh)')
doc.add_paragraph()

add_heading(doc, '4.2 API Proxy — worker.js', level=2)
add_para(doc, 'Deployed on Cloudflare Workers. Responsibilities:')
add_bullet(doc, 'Handles CORS preflight (OPTIONS requests)')
add_bullet(doc, 'Rejects non-POST requests with HTTP 405')
add_bullet(doc, 'Validates X-AralMate-Secret header against ARALMATE_SECRET env var (401 if mismatch)')
add_bullet(doc, 'Forwards valid requests to Moonshot AI with KIMI_API_KEY from encrypted env vars')
add_bullet(doc, 'Returns Kimi response with CORS headers to the browser')
doc.add_paragraph()

add_heading(doc, '4.3 AI Model', level=2)
add_para(doc, (
    'Vision calls use moonshot-v1-8k-vision-preview (image analysis + lesson generation). '
    'Chat follow-up calls use moonshot-v1-32k (text only, larger context for conversation). '
    'Both route through the Cloudflare Worker proxy.'
))
doc.add_paragraph()

# ═══════════════════════════════════════════════════════════
# 5. SECURITY
# ═══════════════════════════════════════════════════════════
add_heading(doc, '5. Security Assessment', level=1)

add_table(doc,
    ['Item', 'Risk', 'Status', 'Notes'],
    [
        ['KIMI_API_KEY exposure',    'High',   'Mitigated',       'Stored as encrypted Cloudflare Secret — never sent to browser'],
        ['APP_SECRET in HTML source','Low',    'Accepted for MVP','Visible via View Source; do NOT reuse this passphrase elsewhere'],
        ['Worker URL discovery',     'Low',    'Accepted for MVP','Secret header prevents unauthorized use even if URL is found'],
        ['Runaway API costs',        'Medium', 'Action required', 'Set monthly spending cap on platform.moonshot.ai account'],
        ['Image data privacy',       'Low',    'Noted',           'Images sent to Moonshot AI — no persistent storage by the app'],
        ['CORS policy',              'Low',    'Open for MVP',    'Worker allows all origins (*); restrict to Netlify domain in production'],
    ],
    [1.8, 0.8, 1.2, 2.6]
)

# ═══════════════════════════════════════════════════════════
# 6. KNOWN LIMITATIONS
# ═══════════════════════════════════════════════════════════
add_heading(doc, '6. Known Limitations', level=1)
add_bullet(doc, 'APP_SECRET is hardcoded in the HTML — visible in browser DevTools. Acceptable for MVP; must not be reused for other systems.')
add_bullet(doc, 'Claude mode removed entirely — app now runs exclusively on Kimi K2.5 vision via the proxy.')
add_bullet(doc, 'No user accounts or login — any device with the URL can use the app.')
add_bullet(doc, 'Session data stored in localStorage — data is device-specific and not synced.')
add_bullet(doc, 'Tala points and streak count are hardcoded starting values, not persisted to a server.')
add_bullet(doc, 'CORS policy is open (*) — should be restricted to the Netlify domain in a future release.')
add_bullet(doc, 'No rate limiting on the Worker — a single user could make unlimited API calls.')
doc.add_paragraph()

# ═══════════════════════════════════════════════════════════
# 7. HOW TO REDEPLOY
# ═══════════════════════════════════════════════════════════
add_heading(doc, '7. How to Update and Redeploy', level=1)

add_heading(doc, '7.1 Updating the app', level=2)
add_bullet(doc, 'Edit D:\\tutor\\aralmate.html in VS Code')
add_bullet(doc, 'Copy the file to D:\\tutor\\deploy\\index.html')
add_bullet(doc, 'Go to netlify.com → project dashboard → Deploys')
add_bullet(doc, 'Drag the deploy\\ folder onto the deploy drop zone')
add_bullet(doc, 'New version goes live in under 30 seconds')
doc.add_paragraph()

add_heading(doc, '7.2 Updating the Worker', level=2)
add_bullet(doc, 'Go to Cloudflare dashboard → Workers & Pages → aralmate-proxy')
add_bullet(doc, 'Click Edit Code → paste updated worker.js → Deploy')
doc.add_paragraph()

add_heading(doc, '7.3 Rotating secrets', level=2)
add_bullet(doc, 'Go to Cloudflare → aralmate-proxy → Settings → Variables and Secrets')
add_bullet(doc, 'Update KIMI_API_KEY or ARALMATE_SECRET values')
add_bullet(doc, 'Update APP_SECRET in aralmate.html to match ARALMATE_SECRET, then redeploy')
doc.add_paragraph()

# ═══════════════════════════════════════════════════════════
# 8. TECH STACK SUMMARY
# ═══════════════════════════════════════════════════════════
add_heading(doc, '8. Tech Stack Summary', level=1)

add_table(doc,
    ['Category', 'Technology', 'Version / Notes'],
    [
        ['Frontend language',   'HTML / CSS / JavaScript', 'Vanilla — no framework'],
        ['Fonts',               'Nunito + Baloo 2',         'Google Fonts CDN'],
        ['Hosting',             'Netlify',                  'Free Starter plan, drag-and-drop deploy'],
        ['API Proxy',           'Cloudflare Workers',       'Free plan, edge-deployed'],
        ['AI Vision Model',     'moonshot-v1-8k-vision-preview', 'Moonshot AI (platform.moonshot.ai)'],
        ['AI Chat Model',       'moonshot-v1-32k',          'Moonshot AI'],
        ['Local dev server',    'Python http.server',       'python -m http.server 8000'],
        ['Source files',        'D:\\tutor\\',              'aralmate.html + worker.js'],
    ],
    [1.5, 2.2, 2.6]
)

# ═══════════════════════════════════════════════════════════
# 9. PHASE 2 RECOMMENDATIONS
# ═══════════════════════════════════════════════════════════
add_heading(doc, '9. Phase 2 Recommendations', level=1)
add_bullet(doc, 'Restrict Cloudflare CORS to Netlify domain only (remove wildcard *)')
add_bullet(doc, 'Add rate limiting to the Worker (max N requests per IP per hour)')
add_bullet(doc, 'Implement server-side Tala points and streak persistence (e.g. Cloudflare KV store)')
add_bullet(doc, 'Add student login so progress syncs across devices')
add_bullet(doc, 'Set Moonshot AI monthly spending cap immediately')
add_bullet(doc, 'Rename Netlify URL to a custom domain (e.g. aralmate.app)')
add_bullet(doc, 'Add support for multiple students / multiple profiles')
doc.add_paragraph()

# ═══════════════════════════════════════════════════════════
# FOOTER NOTE
# ═══════════════════════════════════════════════════════════
doc.add_paragraph()
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run(f'Generated {datetime.date.today().strftime("%B %d, %Y")}  |  AralMate MVP v1.0  |  Confidential')
run.font.name = 'Calibri'
run.font.size = Pt(9)
run.font.color.rgb = RGBColor(107, 132, 168)

# ── Save ─────────────────────────────────────────────────────
doc.save('D:/tutor/AralMate_Build_Review_v1.0.docx')
print('Done: D:/tutor/AralMate_Build_Review_v1.0.docx')
