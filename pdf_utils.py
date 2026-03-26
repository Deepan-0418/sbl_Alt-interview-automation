from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.platypus import Image
import io
from datetime import datetime
import textwrap
import os
import sys
import logging
import pytz

# ── Logging ────────────────────────────────────────────────────
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# ── Timezone ───────────────────────────────────────────────────
IST = pytz.timezone('Asia/Kolkata')

def now_ist():
    """Return current datetime string in IST — correct on both local and Render."""
    return datetime.now(IST).strftime('%Y-%m-%d %H:%M:%S IST')

# ── Base path (PyInstaller or normal) ─────────────────────────
if getattr(sys, 'frozen', False):
    base_path = sys._MEIPASS
else:
    base_path = os.path.dirname(__file__)

# ── Logo path ──────────────────────────────────────────────────
# Check user_dir first (where app.py copies logo.png on first boot),
# then fall back to static/ for environments where it may still live.
def _find_logo():
    upload_folder = os.environ.get('DATA_ROOT', base_path)
    candidates = [
        os.path.join(upload_folder, 'logo.png'),        # user_dir — VS Code & Render
        os.path.join(base_path, 'static', 'logo.png'),  # legacy static/ location
        os.path.join(base_path, 'logo.png'),             # project root fallback
    ]
    for path in candidates:
        if os.path.exists(path):
            logger.info(f"Logo found at: {path}")
            return path
    logger.error("Logo not found in any expected location.")
    return None

LOGO_PATH = _find_logo()

# ── Attempt label helper ───────────────────────────────────────
ATTEMPT_LABELS = {
    1: '1st Attempt',
    2: '2nd Attempt',
    3: '3rd Attempt',
}


# ══════════════════════════════════════════════════════════════
# MAIN RESULTS PDF
# ══════════════════════════════════════════════════════════════

def generate_typing_test_pdf(
    name, typing_results, handwritten_results=None, excel_quiz_results=None,
    excel_score=0, excel_total=0, excel_practical_file=None,
    excel_practical_tasks=None, excel_practical_score=None,
    excel_sheet_scores=None, location="", distance=0.0,
    attempt_number="", signup_date="", dob=""
):
    # ── Sanitize filename parts ────────────────────────────────
    sanitized_name = (
        "".join(c for c in name if c.isalnum() or c in (' ', '_'))
        .strip().replace(' ', '_')
    )
    try:
        signup_date_obj       = datetime.strptime(signup_date, '%Y-%m-%d %H:%M:%S')
        sanitized_date        = signup_date_obj.strftime('%Y%m%d_%H%M%S')
        formatted_signup_date = signup_date_obj.strftime('%d %B %Y')
    except ValueError:
        sanitized_date        = 'unknown_date'
        formatted_signup_date = 'Unknown'

    try:
        dob_obj       = datetime.strptime(dob, '%Y-%m-%d')
        formatted_dob = dob_obj.strftime('%d %B %Y')
        sanitized_dob = dob_obj.strftime('%Y%m%d')
    except ValueError:
        formatted_dob = 'Unknown'
        sanitized_dob = 'unknown_dob'

    sanitized_attempt = str(attempt_number).replace(' ', '_').lower()
    filename = (
        f"Test_Result_{sanitized_name}_{sanitized_dob}"
        f"_{sanitized_date}_{sanitized_attempt}.pdf"
    )

    # ── Document setup ─────────────────────────────────────────
    buffer = io.BytesIO()
    doc    = SimpleDocTemplate(
        buffer, pagesize=letter,
        rightMargin=0.5*inch, leftMargin=0.5*inch,
        topMargin=0.75*inch,  bottomMargin=1.4*inch,
    )

    styles = getSampleStyleSheet()
    custom_styles = {
        'Title': ParagraphStyle(
            name='CustomTitle', parent=styles['Title'],
            fontSize=16, spaceAfter=8, alignment=1, textColor=colors.black,
        ),
        'Heading2': ParagraphStyle(
            name='CustomHeading2', parent=styles['Heading2'],
            fontSize=12, spaceAfter=6, spaceBefore=12, textColor=colors.navy,
        ),
        'Normal': ParagraphStyle(
            name='CustomNormal', parent=styles['Normal'],
            fontSize=8, spaceAfter=4,
        ),
    }

    # ── Header ─────────────────────────────────────────────────
    def add_header(c, doc):
        c.saveState()
        try:
            # Generated time in IST — accurate on both local and Render
            c.setFont('Helvetica', 8)
            c.drawString(
                0.5*inch, letter[1] - 0.4*inch,
                f"Generated on: {now_ist()}"
            )
            if LOGO_PATH and os.path.exists(LOGO_PATH):
                logo_width = 1 * inch
                logo       = Image(LOGO_PATH)
                aspect     = logo.drawHeight / logo.drawWidth
                x_pos = letter[0] - logo_width - 0.5*inch
                y_pos = letter[1] - logo_width * aspect - 0.25*inch
                c.drawImage(
                    LOGO_PATH, x_pos, y_pos,
                    width=logo_width, height=logo_width * aspect, mask='auto'
                )
            else:
                c.setFont('Helvetica', 8)
                c.drawString(letter[0] - 2*inch, letter[1] - 0.4*inch, "Logo not found")
        except Exception as e:
            logger.error(f"Header error: {e}")
            c.setFont('Helvetica', 8)
            c.drawString(letter[0] - 2*inch, letter[1] - 0.4*inch, f"Logo error: {e}")
        c.restoreState()

    # ── Footer ─────────────────────────────────────────────────
    def add_footer(c, doc):
        c.saveState()
        c.setFont('Helvetica', 8)
        c.drawString(0.5*inch,                0.85*inch, "Applicant Signature")
        c.drawCentredString(letter[0]/2,      0.85*inch, "Evaluator Signature")
        c.drawRightString(letter[0]-0.5*inch, 0.85*inch, "Hiring Manager Signature")
        c.setLineWidth(0.5)
        c.line(0.5*inch, 0.65*inch, letter[0]-0.5*inch, 0.65*inch)
        c.drawCentredString(letter[0]/2, 0.5*inch, "© SBL | InterviewAutomation2026")
        c.restoreState()

    story = []

    # ── Title ──────────────────────────────────────────────────
    story.append(Paragraph("Interview Automation Test Results", custom_styles['Title']))
    story.append(Spacer(1, 8))

    # ── Candidate Info ─────────────────────────────────────────
    story.append(Paragraph("Candidate Information", custom_styles['Heading2']))
    candidate_data = [
        ['Name:',           name,                  'Location:',     location],
        ['Sign-up Date:',   formatted_signup_date, 'Date of Birth:', formatted_dob],
        ['Attempt Number:', attempt_number,         'Distance:',     f"{distance} km"],
    ]
    candidate_table = Table(
        candidate_data,
        colWidths=[1.5*inch, 2*inch, 1.5*inch, 2*inch],
        hAlign='LEFT'
    )
    candidate_table.setStyle(TableStyle([
        ('BACKGROUND',    (0, 0), (-1, -1), colors.white),
        ('TEXTCOLOR',     (0, 0), (-1, -1), colors.black),
        ('ALIGN',         (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME',      (0, 0), (0, -1),  'Helvetica-Bold'),
        ('FONTNAME',      (2, 0), (2, -1),  'Helvetica-Bold'),
        ('FONTSIZE',      (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING',    (0, 0), (-1, -1), 4),
        ('GRID',          (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING',   (0, 0), (-1, -1), 5),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 5),
        ('LINEBEFORE',    (2, 0), (2, -1),  2, colors.black),
    ]))
    story.append(candidate_table)
    story.append(Spacer(1, 12))

    # ── Typing Test Results ────────────────────────────────────
    story.append(Paragraph("Typing Test Results", custom_styles['Heading2']))

    # Pass: at least 2 of 3 attempts with WPM >= 25 AND Accuracy >= 90
    pass_count = sum(
        1 for r in typing_results
        if r['wpm'] >= 25 and r['accuracy'] >= 90
    )
    overall_typing_result = (
        'Pass' if typing_results and pass_count >= 2 else 'Fail'
    )

    if typing_results:
        typing_data = [['Attempt', 'WPM', 'Accuracy', 'Time', 'Result']]
        for result in typing_results[:3]:
            attempt_passed = result['wpm'] >= 25 and result['accuracy'] >= 90
            label = ATTEMPT_LABELS.get(result['attempt'], f"Attempt {result['attempt']}")
            typing_data.append([
                label,
                f"{result['wpm']:.1f}",
                f"{result['accuracy']:.1f}%",
                f"{result['time_limit'] // 60}:{result['time_limit'] % 60:02d}",
                'Pass' if attempt_passed else 'Fail',
            ])
        typing_table = Table(
            typing_data,
            colWidths=[1.4*inch, 1*inch, 1.2*inch, 1*inch, 1*inch],
            hAlign='LEFT'
        )
        typing_table.setStyle(TableStyle([
            ('BACKGROUND',    (0, 0), (-1, 0),  colors.grey),
            ('TEXTCOLOR',     (0, 0), (-1, 0),  colors.whitesmoke),
            ('BACKGROUND',    (0, 1), (-1, -1), colors.white),
            ('TEXTCOLOR',     (0, 1), (-1, -1), colors.black),
            ('ALIGN',         (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME',      (0, 0), (-1, 0),  'Helvetica-Bold'),
            ('FONTSIZE',      (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING',    (0, 0), (-1, -1), 4),
            ('GRID',          (0, 0), (-1, -1), 0.5, colors.black),
            ('VALIGN',        (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING',   (0, 0), (-1, -1), 5),
            ('RIGHTPADDING',  (0, 0), (-1, -1), 5),
        ]))
        story.append(typing_table)
    else:
        story.append(Table(
            [["No typing test results available."]],
            colWidths=[6.5*inch], hAlign='LEFT'
        ))

    # Pass/Fail badge
    typing_result_table = Table(
        [[f"Result: {overall_typing_result}"]],
        colWidths=[2*inch], hAlign='RIGHT'
    )
    typing_result_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0),(-1,-1), colors.green if overall_typing_result == 'Pass' else colors.red),
        ('TEXTCOLOR',  (0,0),(-1,-1), colors.white),
        ('ALIGN',      (0,0),(-1,-1), 'CENTER'),
        ('FONTNAME',   (0,0),(-1,-1), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0),(-1,-1), 8),
        ('BOX',        (0,0),(-1,-1), 1, colors.black),
        ('PADDING',    (0,0),(-1,-1), 4),
    ]))
    story.append(typing_result_table)
    story.append(Spacer(1, 12))

    # ── Handwritten Results ────────────────────────────────────
    story.append(Paragraph("Handwritten Verification Results", custom_styles['Heading2']))
    handwritten_correct = (
        sum(1 for r in handwritten_results if r['status'] == 'Correct')
        if handwritten_results else 0
    )
    handwritten_total  = len(handwritten_results) if handwritten_results else 0
    handwritten_result = 'Pass' if handwritten_correct >= 8 else 'Fail'

    if handwritten_results:
        hw_table = Table(
            [['Correct Answers:', f"{handwritten_correct}/{handwritten_total}"]],
            colWidths=[1.5*inch, 4*inch], hAlign='LEFT'
        )
        hw_table.setStyle(TableStyle([
            ('BACKGROUND',    (0,0),(-1,-1), colors.white),
            ('TEXTCOLOR',     (0,0),(-1,-1), colors.black),
            ('ALIGN',         (0,0),(-1,-1), 'LEFT'),
            ('FONTNAME',      (0,0),(0,-1),  'Helvetica-Bold'),
            ('FONTSIZE',      (0,0),(-1,-1), 8),
            ('BOTTOMPADDING', (0,0),(-1,-1), 4),
            ('TOPPADDING',    (0,0),(-1,-1), 4),
            ('GRID',          (0,0),(-1,-1), 0.5, colors.black),
            ('VALIGN',        (0,0),(-1,-1), 'MIDDLE'),
            ('LEFTPADDING',   (0,0),(-1,-1), 5),
            ('RIGHTPADDING',  (0,0),(-1,-1), 5),
        ]))
        story.append(hw_table)
    else:
        story.append(Table(
            [["No handwritten test results available."]],
            colWidths=[6.5*inch], hAlign='LEFT'
        ))

    hw_result_table = Table(
        [[f"Result: {handwritten_result}"]],
        colWidths=[2*inch], hAlign='RIGHT'
    )
    hw_result_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0),(-1,-1), colors.green if handwritten_result == 'Pass' else colors.red),
        ('TEXTCOLOR',  (0,0),(-1,-1), colors.white),
        ('ALIGN',      (0,0),(-1,-1), 'CENTER'),
        ('FONTNAME',   (0,0),(-1,-1), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0),(-1,-1), 8),
        ('BOX',        (0,0),(-1,-1), 1, colors.black),
        ('PADDING',    (0,0),(-1,-1), 4),
    ]))
    story.append(hw_result_table)
    story.append(Spacer(1, 12))

    # ── Excel Results ──────────────────────────────────────────
    story.append(Paragraph("Excel Results", custom_styles['Heading2']))
    excel_total_marks = excel_score + (
        sum(excel_sheet_scores.values()) if excel_sheet_scores else 0
    )
    excel_result = 'Pass' if excel_total_marks >= 15 else 'Fail'

    excel_data = [
        ['Quiz:',      f"Correct Answers: {excel_score}/{excel_total}", f"Excel Total: {excel_total_marks}/20"],
        ['Practical:', f"Total Score: {sum(excel_sheet_scores.values()) if excel_sheet_scores else 0}/10", ''],
    ]
    excel_table = Table(excel_data, colWidths=[1.5*inch, 2*inch, 2*inch], hAlign='LEFT')
    excel_table.setStyle(TableStyle([
        ('BACKGROUND',    (0,0),(-1,-1), colors.white),
        ('TEXTCOLOR',     (0,0),(-1,-1), colors.black),
        ('ALIGN',         (0,0),(-1,-1), 'LEFT'),
        ('FONTNAME',      (0,0),(0,-1),  'Helvetica-Bold'),
        ('FONTNAME',      (2,0),(2,0),   'Helvetica-Bold'),
        ('FONTSIZE',      (0,0),(-1,-1), 8),
        ('BOTTOMPADDING', (0,0),(-1,-1), 4),
        ('TOPPADDING',    (0,0),(-1,-1), 4),
        ('GRID',          (0,0),(-1,-1), 0.5, colors.black),
        ('VALIGN',        (0,0),(-1,-1), 'MIDDLE'),
        ('LEFTPADDING',   (0,0),(-1,-1), 5),
        ('RIGHTPADDING',  (0,0),(-1,-1), 5),
        ('SPAN',          (2,0),(2,1)),
    ]))
    story.append(excel_table)

    excel_result_table = Table(
        [[f"Result: {excel_result}"]],
        colWidths=[2*inch], hAlign='RIGHT'
    )
    excel_result_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0),(-1,-1), colors.green if excel_result == 'Pass' else colors.red),
        ('TEXTCOLOR',  (0,0),(-1,-1), colors.white),
        ('ALIGN',      (0,0),(-1,-1), 'CENTER'),
        ('FONTNAME',   (0,0),(-1,-1), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0),(-1,-1), 8),
        ('BOX',        (0,0),(-1,-1), 1, colors.black),
        ('PADDING',    (0,0),(-1,-1), 4),
    ]))
    story.append(excel_result_table)
    story.append(Spacer(1, 12))

    # ── Overall Result ─────────────────────────────────────────
    story.append(Paragraph("Overall Result", custom_styles['Heading2']))
    section_results    = [overall_typing_result, handwritten_result, excel_result]
    overall_pass_count = sum(1 for r in section_results if r == 'Pass')
    overall_result     = 'Pass' if overall_pass_count >= 3 else 'Fail'

    overall_table = Table([[overall_result]], colWidths=[7*inch], hAlign='LEFT')
    overall_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0),(-1,-1), colors.green if overall_result == 'Pass' else colors.red),
        ('TEXTCOLOR',  (0,0),(-1,-1), colors.white),
        ('ALIGN',      (0,0),(-1,-1), 'CENTER'),
        ('FONTNAME',   (0,0),(-1,-1), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0),(-1,-1), 9),
        ('BOX',        (0,0),(-1,-1), 1, colors.black),
        ('PADDING',    (0,0),(-1,-1), 6),
    ]))
    story.append(overall_table)
    story.append(Spacer(1, 12))

    # ── Build PDF ──────────────────────────────────────────────
    doc.build(
        story,
        onFirstPage=lambda c, d: (add_header(c, d), add_footer(c, d)),
        onLaterPages=lambda c, d: (add_header(c, d), add_footer(c, d)),
    )
    buffer.seek(0)
    return buffer, filename


# ══════════════════════════════════════════════════════════════
# ERROR REPORT PDF
# ══════════════════════════════════════════════════════════════

def generate_error_report_pdf(
    name, handwritten_results=None, excel_quiz_results=None,
    signup_date="", dob=""
):
    # ── Sanitize filename parts ────────────────────────────────
    sanitized_name = (
        "".join(c for c in name if c.isalnum() or c in (' ', '_'))
        .strip().replace(' ', '_')
    )
    try:
        signup_date_obj = datetime.strptime(signup_date, '%Y-%m-%d %H:%M:%S')
        sanitized_date  = signup_date_obj.strftime('%Y%m%d_%H%M%S')
    except ValueError:
        sanitized_date  = 'unknown_date'

    try:
        dob_obj       = datetime.strptime(dob, '%Y-%m-%d')
        sanitized_dob = dob_obj.strftime('%Y%m%d')
    except ValueError:
        sanitized_dob = 'unknown_dob'

    filename = f"Error_Report_{sanitized_name}_{sanitized_dob}_{sanitized_date}.pdf"

    # ── Document setup ─────────────────────────────────────────
    buffer = io.BytesIO()
    doc    = SimpleDocTemplate(
        buffer, pagesize=letter,
        rightMargin=0.5*inch, leftMargin=0.5*inch,
        topMargin=0.75*inch,  bottomMargin=0.75*inch,
    )

    styles = getSampleStyleSheet()
    custom_styles = {
        'Title': ParagraphStyle(
            name='CustomTitle', parent=styles['Title'],
            fontSize=18, spaceAfter=20, alignment=1,
        ),
        'Heading2': ParagraphStyle(
            name='CustomHeading2', parent=styles['Heading2'],
            fontSize=14, spaceAfter=12, spaceBefore=12, textColor=colors.navy,
        ),
        'Normal': ParagraphStyle(
            name='CustomNormal', parent=styles['Normal'],
            fontSize=10, spaceAfter=8,
        ),
    }

    # ── Header ─────────────────────────────────────────────────
    def add_header(c, doc):
        c.saveState()
        try:
            # Generated time in IST — accurate on both local and Render
            c.setFont('Helvetica', 9)
            c.drawString(
                0.5*inch, letter[1] - 0.65*inch,
                f"Generated on: {now_ist()}"
            )
            if LOGO_PATH and os.path.exists(LOGO_PATH):
                logo_width = 1 * inch
                logo       = Image(LOGO_PATH)
                aspect     = logo.drawHeight / logo.drawWidth
                x_pos = letter[0] - logo_width - 0.5*inch
                y_pos = letter[1] - logo_width * aspect - 0.25*inch
                c.drawImage(
                    LOGO_PATH, x_pos, y_pos,
                    width=logo_width, height=logo_width * aspect, mask='auto'
                )
            else:
                c.setFont('Helvetica', 9)
                c.drawString(letter[0] - 2*inch, letter[1] - 0.65*inch, "Logo not found")
        except Exception as e:
            logger.error(f"Header error: {e}")
            c.setFont('Helvetica', 9)
            c.drawString(letter[0] - 2*inch, letter[1] - 0.65*inch, f"Logo error: {e}")
        c.restoreState()

    # ── Footer ─────────────────────────────────────────────────
    def add_footer(c, doc):
        c.saveState()
        c.setFont('Helvetica', 9)
        c.drawCentredString(letter[0]/2, 0.5*inch, f"Page {c.getPageNumber()}")
        c.restoreState()

    story = []
    story.append(Paragraph("Interview Automation Error Report", custom_styles['Title']))
    story.append(Spacer(1, 12))

    # ── Handwritten errors ─────────────────────────────────────
    if handwritten_results:
        correct_count = sum(1 for r in handwritten_results if r['status'] == 'Correct')
        total_count   = len(handwritten_results)

        story.append(Paragraph("Handwritten Verification Results", custom_styles['Heading2']))
        story.append(Paragraph(
            f"Correct Answers: {correct_count}/{total_count}",
            custom_styles['Normal']
        ))

        incorrect = [r for r in handwritten_results if r['status'] == 'Incorrect']
        if incorrect:
            story.append(Spacer(1, 12))
            story.append(Paragraph("Incorrect Handwritten Answers:", custom_styles['Heading2']))

            hw_data = [['Image', 'User Input', 'Correct Text']]
            for result in incorrect:
                user_input   = result.get('user_input',   'None')
                correct_text = result.get('correct_text', 'None')
                hw_data.append([
                    Paragraph(result.get('image', 'Unknown'), custom_styles['Normal']),
                    Paragraph(
                        '<br/>'.join(textwrap.wrap(user_input,   50)) if user_input   else 'None',
                        custom_styles['Normal']
                    ),
                    Paragraph(
                        '<br/>'.join(textwrap.wrap(correct_text, 50)) if correct_text else 'None',
                        custom_styles['Normal']
                    ),
                ])

            hw_table = Table(hw_data, colWidths=[2*inch, 2.5*inch, 2.5*inch])
            hw_table.setStyle(TableStyle([
                ('BACKGROUND',    (0,0),(-1,0),  colors.grey),
                ('TEXTCOLOR',     (0,0),(-1,0),  colors.whitesmoke),
                ('ALIGN',         (0,0),(-1,-1), 'LEFT'),
                ('FONTNAME',      (0,0),(-1,0),  'Helvetica-Bold'),
                ('FONTSIZE',      (0,0),(-1,-1), 10),
                ('BACKGROUND',    (0,1),(-1,-1), colors.beige),
                ('GRID',          (0,0),(-1,-1), 0.5, colors.black),
                ('VALIGN',        (0,0),(-1,-1), 'MIDDLE'),
                ('LEFTPADDING',   (0,0),(-1,-1), 5),
                ('RIGHTPADDING',  (0,0),(-1,-1), 5),
                ('TOPPADDING',    (0,0),(-1,-1), 5),
                ('BOTTOMPADDING', (0,0),(-1,-1), 5),
            ]))
            story.append(hw_table)

        story.append(Spacer(1, 24))

    # ── Excel quiz errors ──────────────────────────────────────
    if excel_quiz_results:
        eq_score = sum(1 for r in excel_quiz_results if r['status'] == 'Correct')
        eq_total = len(excel_quiz_results)

        story.append(Paragraph("Excel Quiz Results", custom_styles['Heading2']))
        story.append(Paragraph(
            f"Correct Answers: {eq_score}/{eq_total}",
            custom_styles['Normal']
        ))

        incorrect = [r for r in excel_quiz_results if r['status'] == 'Incorrect']
        if incorrect:
            story.append(Spacer(1, 12))
            story.append(Paragraph("Incorrect Excel Quiz Answers:", custom_styles['Heading2']))

            eq_data = [['Question', 'User Answer', 'Correct Answer']]
            for result in incorrect:
                question       = result.get('question',       'Unknown')
                user_answer    = result.get('user_answer',    'None')
                correct_answer = result.get('correct_answer', 'None')
                eq_data.append([
                    Paragraph(
                        '<br/>'.join(textwrap.wrap(question,       50)) if question       else 'Unknown',
                        custom_styles['Normal']
                    ),
                    Paragraph(
                        '<br/>'.join(textwrap.wrap(user_answer,    50)) if user_answer    else 'None',
                        custom_styles['Normal']
                    ),
                    Paragraph(
                        '<br/>'.join(textwrap.wrap(correct_answer, 50)) if correct_answer else 'None',
                        custom_styles['Normal']
                    ),
                ])

            eq_table = Table(eq_data, colWidths=[2.5*inch, 2*inch, 2*inch])
            eq_table.setStyle(TableStyle([
                ('BACKGROUND',    (0,0),(-1,0),  colors.grey),
                ('TEXTCOLOR',     (0,0),(-1,0),  colors.whitesmoke),
                ('ALIGN',         (0,0),(-1,-1), 'LEFT'),
                ('FONTNAME',      (0,0),(-1,0),  'Helvetica-Bold'),
                ('FONTSIZE',      (0,0),(-1,-1), 10),
                ('BACKGROUND',    (0,1),(-1,-1), colors.beige),
                ('GRID',          (0,0),(-1,-1), 0.5, colors.black),
                ('VALIGN',        (0,0),(-1,-1), 'MIDDLE'),
                ('LEFTPADDING',   (0,0),(-1,-1), 5),
                ('RIGHTPADDING',  (0,0),(-1,-1), 5),
                ('TOPPADDING',    (0,0),(-1,-1), 5),
                ('BOTTOMPADDING', (0,0),(-1,-1), 5),
            ]))
            story.append(eq_table)

        story.append(Spacer(1, 24))

    # ── Build PDF ──────────────────────────────────────────────
    doc.build(
        story,
        onFirstPage=lambda c, d: (add_header(c, d), add_footer(c, d)),
        onLaterPages=lambda c, d: (add_header(c, d), add_footer(c, d)),
    )
    buffer.seek(0)
    return buffer, filename