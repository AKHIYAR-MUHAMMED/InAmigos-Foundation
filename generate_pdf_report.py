import os
from fpdf import FPDF

class NGOReportPDF(FPDF):
    def header(self):
        # We don't want a header on page 1, or maybe a simple one
        pass
        
    def footer(self):
        # Add a page number footer
        self.set_y(-15)
        self.set_font("helvetica", "I", 8)
        self.set_text_color(74, 85, 104) # Slate Color
        # Text: InAmigos Foundation | AI-Powered Data Analysis Report (2026) - Page X of Y
        text = f"InAmigos Foundation | AI-Powered Data Analysis Report (2026)  -  Page {self.page_no()} of {{nb}}"
        self.cell(0, 10, text, border=0, align="R")

def generate_pdf():
    pdf = NGOReportPDF(orientation="P", unit="mm", format="A4")
    pdf.set_margins(25.4, 25.4, 25.4) # 1 inch margins
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.alias_nb_pages()
    
    # Page 1
    pdf.add_page()
    
    # Primary colors
    NAVY = (26, 54, 93)
    TEAL = (49, 151, 149)
    CHARCOAL = (45, 55, 72)
    SLATE = (74, 85, 104)
    LIGHT_GRAY = (247, 250, 252)
    BORDER_GRAY = (226, 232, 240)
    
    # Document Title
    pdf.set_font("helvetica", "B", 18)
    pdf.set_text_color(*NAVY)
    pdf.cell(0, 10, "ARTIFICIAL INTELLIGENCE IN NGO OPERATIONS", new_x="LMARGIN", new_y="NEXT", align="C")
    
    # Subtitle
    pdf.set_font("helvetica", "I", 11)
    pdf.set_text_color(*SLATE)
    pdf.cell(0, 6, "Enhancing Fundraising, Donor Retention, and Operational Efficiency in 2026", new_x="LMARGIN", new_y="NEXT", align="C")
    
    # Spacing & line divider
    pdf.ln(4)
    pdf.set_draw_color(*TEAL)
    pdf.set_line_width(0.5)
    pdf.line(25.4, pdf.get_y(), 210 - 25.4, pdf.get_y())
    pdf.ln(6)
    
    # Helper for Section Heading (H1)
    def add_h1(text):
        pdf.set_font("helvetica", "B", 13)
        pdf.set_text_color(*NAVY)
        pdf.ln(4)
        pdf.cell(0, 8, text, new_x="LMARGIN", new_y="NEXT")
        # Add a thin bottom border line under the text
        pdf.set_draw_color(*TEAL)
        pdf.set_line_width(0.3)
        pdf.line(25.4, pdf.get_y(), 210 - 25.4, pdf.get_y())
        pdf.ln(3)
        
    # Helper for Subsection Heading (H2)
    def add_h2(text):
        pdf.set_font("helvetica", "B", 10.5)
        pdf.set_text_color(*TEAL)
        pdf.ln(2)
        pdf.cell(0, 6, text, new_x="LMARGIN", new_y="NEXT")
        
    # Helper for Body Paragraph
    def add_body(text, bold_prefix=None):
        pdf.set_font("helvetica", "", 10)
        pdf.set_text_color(*CHARCOAL)
        
        if bold_prefix:
            pdf.set_font("helvetica", "B", 10)
            pdf.write(5, bold_prefix)
            pdf.set_font("helvetica", "", 10)
            
        pdf.write(5, text)
        pdf.ln(6) # line spacing / paragraph spacing
        
    # Helper for Bullet point
    def add_bullet(text, bold_prefix=None):
        pdf.set_font("helvetica", "", 10)
        pdf.set_text_color(*CHARCOAL)
        
        # Draw a small teal bullet square
        x = pdf.get_x()
        y = pdf.get_y()
        pdf.set_fill_color(*TEAL)
        pdf.set_draw_color(*TEAL)
        pdf.rect(x + 2, y + 1.5, 1.5, 1.5, style="F")
        
        # Write text with indent
        pdf.set_x(x + 6)
        if bold_prefix:
            pdf.set_font("helvetica", "B", 10)
            pdf.write(5, bold_prefix)
            pdf.set_font("helvetica", "", 10)
            
        pdf.write(5, text)
        pdf.ln(6)
        
    # Executive Summary H1
    add_h1("Executive Summary")
    
    # Executive Summary Callout Box
    summary_text = (
        "As nonprofit organizations navigate a landscape of changing donor habits and stagnating traditional retention rates in 2026, "
        "Artificial Intelligence (AI) has emerged as a critical driver of operational capacity and strategic engagement. Rather than merely "
        "automating administrative workflows, leading NGOs are leveraging predictive AI and data-driven personalization to transition from generic mass appeals "
        "to deeper, individualized donor relations. This report examines current sector-wide adoption trends, analyzes key metrics, shares "
        "successful real-world case studies, and outlines responsible implementation strategies for forward-thinking NGOs like the InAmigos Foundation."
    )
    
    box_x = 25.4
    box_y = pdf.get_y()
    
    # Light gray background container
    pdf.set_fill_color(*LIGHT_GRAY)
    pdf.set_draw_color(*LIGHT_GRAY)
    pdf.set_font("helvetica", "I", 9.5)
    pdf.set_text_color(*CHARCOAL)
    
    # Print multi cell inside the box
    pdf.multi_cell(159.2, 5, summary_text, border=0, align="L", fill=True)
    
    # Draw left teal border
    box_end_y = pdf.get_y()
    pdf.set_draw_color(*TEAL)
    pdf.set_line_width(1.2)
    pdf.line(box_x, box_y, box_x, box_end_y)
    pdf.ln(4)
    
    # 1. AI Adoption Trends & Statistics in 2026
    add_h1("1. AI Adoption Trends & Statistics")
    add_body(
        "The current year has marked a transition from experimental, ad-hoc AI usage to structured, CRM-integrated workflows. "
        "However, this widespread adoption highlights a unique dichotomy in how technology is utilized across the sector:"
    )
    
    add_bullet(
        "92% of nonprofits report using AI in their workflows in 2026. Despite this high figure, only 7% report that AI has fundamentally transformed their mission delivery or strategic outcomes. Most organizations remain on an 'efficiency plateau,' primarily employing AI for basic task acceleration (drafting emails, summarizing meeting logs, generating ideas).",
        bold_prefix="The Adoption Paradox (92% vs. 7%): "
    )
    
    add_bullet(
        "Approximately 30% of organizations report direct, measurable increases in fundraising revenue as a result of integrating AI tools. This indicates a significant return on investment for those who leverage AI beyond simple content drafting.",
        bold_prefix="Fundraising Revenue Impact: "
    )
    
    add_bullet(
        "On average, nonprofit staff members utilizing integrated AI tools report saving between 15 to 20 hours per week on administrative data manipulation, audience-list creation, and report preparation. This time is directly reallocated toward high-value activities, such as major donor cultivation and community outreach.",
        bold_prefix="Operational Efficiency Gains: "
    )
    
    # 2. Donor Retention and Hyper-Personalization
    add_h1("2. Optimizing Donor Retention & Engagement")
    add_body(
        "With recent donor participation numbers showing overall flat or downward trends globally, retaining active donors is a primary financial priority. Nonprofits are shifting focus from donor acquisition to donor retention using two AI methods:"
    )
    
    add_h2("Predictive Churn Modeling")
    add_body(
        "By analyzing historical transaction records, engagement frequency, event attendance, and newsletter open rates, predictive AI algorithms calculate individual donor 'risk scores.' Fundraisers are automatically alerted when a previously active donor's behavior matches historical churn patterns (e.g., missed email openings, skipping a seasonal campaign). This allows staff to intervene with customized re-engagement strategies before the donor lapses completely."
    )
    
    add_h2("Hyper-Personalization and Donor Trust")
    add_body(
        "Generic, one-size-fits-all fundraising appeals are increasingly ineffective. In 2026, AI-driven segmentation matches donor profiles with the exact projects they support (e.g., highlighting clean water outcomes to clean water advocates). However, trust is paramount:"
    )
    
    add_bullet(
        "93% of donors state that transparency regarding how an organization uses AI is highly important to them.",
        bold_prefix="Demanded Transparency: "
    )
    add_bullet(
        "Nearly 40% of donors express concern or discomfort regarding how their personal and financial data is analyzed by AI systems. This underscores the urgent need for clear data governance and privacy policies.",
        bold_prefix="Data Privacy Concerns: "
    )
    
    # Page Break
    pdf.add_page()
    
    # 3. Real-world Case Studies
    add_h1("3. Real-World Case Studies")
    add_body(
        "Real-world deployments demonstrate the potential of AI when paired with human intervention and empathetic storytelling."
    )
    
    add_h2("Case Study 1: Greenpeace Australia Pacific - Predictive Outreach")
    add_body(
        "To combat donor attrition, Greenpeace Australia Pacific integrated a predictive machine learning tool (Dataro Predict) with their CRM. The model assigned real-time retention scores to active monthly donors, identifying those highly likely to churn. Instead of a broad, expensive direct-mail campaign, the organization conducted high-touch, personalized telephone calls exclusively to the at-risk segment. The results were stark:"
    )
    add_bullet("Targeted regular donors were 2.5 times less likely to cancel their subscriptions compared to control groups.")
    add_bullet("The targeted campaign generated $235,000 in saved donations over an 18-month tracking period.")
    add_bullet("The campaign achieved a return on investment (ROI) of over 10:1.")
    
    add_h2("Case Study 2: Charity:Water - Immersive AI Storytelling")
    add_body(
        "Charity:Water deployed an interactive, conversational AI chatbot named 'Yeshi' via Facebook Messenger. Instead of asking for funds immediately, the AI walked users through the daily life of a woman in Ethiopia fetching water. The bot simulated real-time conversation, provided maps of the route, and shared photos and audio files. This storytelling approach transformed donor engagement:"
    )
    add_bullet("It built deep, cognitive empathy before asking for financial support.")
    add_bullet("It successfully engaged younger demographics who typically avoid email-based newsletters.")
    add_bullet("It established a conversational model for donor onboarding that continues to influence fundraising standards.")
    
    # 4. Comparative Metrics Table
    add_h1("4. Key Performance Metrics for AI Adoption")
    add_body(
        "NGOs tracking their digital transformation should evaluate performance across these three main categories:"
    )
    
    # Draw table
    pdf.ln(2)
    # Total width is 159.2 mm. Let's make columns: 45mm, 50mm, 64.2mm.
    col_widths = [45, 50, 64.2]
    row_height = 8
    
    # Header
    pdf.set_fill_color(*NAVY)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("helvetica", "B", 9)
    pdf.cell(col_widths[0], row_height, "Category", border=1, align="L", fill=True)
    pdf.cell(col_widths[1], row_height, "Primary Metric", border=1, align="L", fill=True)
    pdf.cell(col_widths[2], row_height, "Target Benchmark (2026)", border=1, align="L", fill=True)
    pdf.ln()
    
    # Rows
    rows_data = [
        ("Fundraising Effectiveness", "Donor Retention & Churn", "+15% increase in retention for targeted groups"),
        ("Operational Efficiency", "Administrative Time Saved", "15-20 hours saved weekly per staff member"),
        ("Strategic Impact", "Donor Engagement (Email Open/Click)", "Double the industry standard (>45% open)")
    ]
    
    pdf.set_text_color(*CHARCOAL)
    pdf.set_font("helvetica", "", 8.5)
    for idx, (cat, metric, bench) in enumerate(rows_data):
        bg = LIGHT_GRAY if idx % 2 == 0 else (255, 255, 255)
        pdf.set_fill_color(*bg)
        
        x = pdf.get_x()
        y = pdf.get_y()
        pdf.multi_cell(col_widths[0], 6, cat, border=1, align="L", fill=True)
        h1 = pdf.get_y() - y
        
        pdf.set_xy(x + col_widths[0], y)
        pdf.multi_cell(col_widths[1], 6, metric, border=1, align="L", fill=True)
        h2 = pdf.get_y() - y
        
        pdf.set_xy(x + col_widths[0] + col_widths[1], y)
        pdf.multi_cell(col_widths[2], 6, bench, border=1, align="L", fill=True)
        h3 = pdf.get_y() - y
        
        max_h = max(h1, h2, h3)
        pdf.set_xy(25.4, y + max_h)
        
    pdf.ln(4)
    
    # 5. Strategic Recommendations for InAmigos Foundation
    add_h1("5. Strategic Recommendations for InAmigos Foundation")
    
    add_bullet(
        "Establish an ethical AI policy and share it publicly on your website. Since 93% of donors prize AI transparency, stating clearly how donor data is handled builds foundational trust.",
        bold_prefix="1. Build Transparency First: "
    )
    
    add_bullet(
        "Ensure CRM databases have clean, uniform, and deduplicated records. Predictive churn algorithms and personalization systems fail if underlying donor profiles contain errors or incomplete data.",
        bold_prefix="2. Audit CRM Data Quality: "
    )
    
    add_bullet(
        "Utilize AI-generated content (emails, social posts, letters) as initial drafts only. Incorporate human reviews for tone, local cultural relevance, and alignment with the foundation's core values before dispatch.",
        bold_prefix="3. Maintain the Human-in-the-Loop: "
    )
    
    add_bullet(
        "Introduce micro-experiments using free or low-cost AI tools (like ChatGPT for initial proposal brainstorming or small predictive scripts on donor spreadsheets) before investing in enterprise-grade platforms.",
        bold_prefix="4. Start with Focused Micro-Experiments: "
    )
    
    pdf.output("AI_NGO_Fundraising_Report.pdf")
    print("PDF generated successfully!")

if __name__ == "__main__":
    generate_pdf()
