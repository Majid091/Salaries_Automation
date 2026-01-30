from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import inch, cm, mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from PyPDF2 import PdfReader, PdfWriter
from PIL import Image
from io import BytesIO
from datetime import datetime
import os

# TechEmulsion Brand Colors (exact from letter head)
TEAL_PRIMARY = colors.HexColor('#0e8282')  # Table header color
DARK_TEAL = colors.HexColor('#073630')  # Net Salary box color
WHITE = colors.white
BLACK = colors.HexColor('#212121')
SOLID_BLACK = colors.HexColor('#000000')  # Pure black for table text and labels
LIGHT_GRAY = colors.HexColor('#F5F5F5')
GRAY_BORDER = colors.HexColor('#BDBDBD')

# Page size (A4 - same as letter head)
PAGE_WIDTH, PAGE_HEIGHT = A4  # 595.2 x 841.92 points


class PDFGenerator:
    def __init__(self):
        self.styles = getSampleStyleSheet()
        self._setup_custom_styles()
        self.company_name = "TECH EMULSION"
        self.letter_head_path = os.path.join(os.path.dirname(__file__), "letter_head", "letter head 01 ff.pdf")
        self.stamp_path = os.path.join(os.path.dirname(__file__), "stamp", "TE_STAMP.png")
        self.paid_stamp_path = os.path.join(os.path.dirname(__file__), "stamp", "TE_STAMP.png")

    def set_company_info(self, company_name, app_name=""):
        """Set company information"""
        if company_name:
            self.company_name = company_name

    def _setup_custom_styles(self):
        """Setup custom styles for PDF using brand colors"""
        self.title_style = ParagraphStyle(
            'SalarySlipTitle',
            parent=self.styles['Heading1'],
            fontSize=16,
            textColor=DARK_TEAL,
            spaceAfter=20,
            spaceBefore=5,
            alignment=TA_CENTER,
            fontName='Helvetica-Bold'
        )

        self.label_style = ParagraphStyle(
            'LabelStyle',
            parent=self.styles['Normal'],
            fontSize=10,
            fontName='Helvetica-Bold',
            textColor=BLACK
        )

        self.value_style = ParagraphStyle(
            'ValueStyle',
            parent=self.styles['Normal'],
            fontSize=10,
            fontName='Helvetica',
            textColor=BLACK
        )

        self.footer_style = ParagraphStyle(
            'FooterStyle',
            parent=self.styles['Normal'],
            fontSize=8,
            textColor=BLACK,
            alignment=TA_RIGHT
        )

    def create_pdf(self, record, month_name):
        """Create a PDF salary slip using letter head as background"""
        filename = self.get_pdf_filename(record, month_name)
        pdf_path = os.path.join("pdfs", filename)

        # Create content PDF in memory
        content_buffer = BytesIO()
        self._create_content_pdf(content_buffer, record, month_name)
        content_buffer.seek(0)

        # Merge letter head background with content
        self._merge_with_letterhead(content_buffer, pdf_path)

        return pdf_path

    def _create_content_pdf(self, buffer, record, month_name):
        """Create the salary slip content PDF"""
        c = canvas.Canvas(buffer, pagesize=A4)

        # Starting Y position (below header area)
        y_pos = PAGE_HEIGHT - 150

        y_pos -= 40

        # Draw SALARY SLIP title on the left side
        c.setFont('Helvetica-Bold', 18)
        c.setFillColor(SOLID_BLACK)
        c.drawString(50, y_pos, "SALARY SLIP")

        # Employee Information on the right side
        name = self._get_field_value(record, ['Name', 'name']) or '___________________'
        designation = self._get_field_value(record, ['Designation', 'designation']) or '___________________'

        # Define fixed positions for alignment
        label_x = PAGE_WIDTH - 280  # Label start position
        value_x = PAGE_WIDTH - 175  # Value start position (same for all values)

        # Employee Name (right side)
        c.setFont('Helvetica-Bold', 10)
        c.setFillColor(SOLID_BLACK)
        c.drawString(label_x, y_pos, "Employee Name:")
        c.setFont('Helvetica', 10)
        c.drawString(value_x, y_pos, str(name))

        y_pos -= 20

        # Designation (right side)
        c.setFont('Helvetica-Bold', 10)
        c.setFillColor(SOLID_BLACK)
        c.drawString(label_x, y_pos, "Designation:")
        c.setFont('Helvetica', 10)
        c.drawString(value_x, y_pos, str(designation))

        y_pos -= 20

        # Month & Year (right side)
        c.setFont('Helvetica-Bold', 10)
        c.setFillColor(SOLID_BLACK)
        c.drawString(label_x, y_pos, "Month & Year:")
        c.setFont('Helvetica', 10)
        c.drawString(value_x, y_pos, month_name)

        y_pos -= 35

        # Draw salary table and get table dimensions
        y_pos, table_x, table_width = self._draw_salary_table(c, record, y_pos)

        # Draw net salary box
        y_pos = self._draw_net_salary(c, record, y_pos)

        # Draw company stamp with transparency below the table
        self._draw_stamp(c, table_x, table_width, y_pos)

        # Draw Company Stamp section just above Generated on
        self._draw_company_stamp_section(c)

        # Draw generation date - positioned on left side
        c.setFont('Helvetica', 8)
        c.setFillColor(BLACK)
        generated_text = f"Generated on: {datetime.now().strftime('%B %d, %Y')}"
        c.drawString(50, 87, generated_text)

        c.save()

    def _draw_salary_table(self, c, record, y_start):
        """Draw the earnings and deductions table"""
        # Table dimensions
        table_x = 50
        table_width = PAGE_WIDTH - 100
        col_width = table_width / 2
        row_height = 22
        header_height = 25

        # Get earnings data
        earnings_fields = [
            (['Basic Salary', 'basic salary'], 'Basic Salary'),
            (['Food Allowance', 'food allowance'], 'Food Allowance'),
            (['Travel Allowance', 'travel allowance'], 'Travel Allowance'),
            (['Medical Allowance', 'medical allowance'], 'Medical Allowance'),
            (['Other (subscriptions)', 'other (subscriptions)'], 'Subscriptions'),
            (['Other (Overtime)', 'other (overtime)'], 'Overtime'),
            (['Other (Leave Encashment)', 'other (leave encashment)'], 'Leave Encashment'),
            (['Other (Commision)', 'other (commision)', 'Commission'], 'Commission'),
            (['Others', 'others'], 'Others'),
        ]

        deduction_fields = [
            (['Tax Deductable', 'tax deductable', 'Tax Deduction'], 'Tax Deduction'),
            (['Other (Extra Leaves)', 'other (extra leaves)'], 'Extra Leaves'),
        ]

        # Calculate earnings
        earnings_data = []
        total_earnings = 0.0
        for field_names, label in earnings_fields:
            value = self._get_field_value(record, field_names)
            amount = self._parse_amount(value)
            if amount > 0:
                earnings_data.append((label, amount))
                total_earnings += amount

        # Calculate deductions
        deductions_data = []
        total_deductions = 0.0
        for field_names, label in deduction_fields:
            value = self._get_field_value(record, field_names)
            amount = self._parse_amount(value)
            if amount > 0:
                deductions_data.append((label, amount))
                total_deductions += amount

        # Determine number of rows needed
        max_rows = max(len(earnings_data), len(deductions_data), 1)

        # Draw header row
        y = y_start

        # Earnings header (left)
        c.setFillColor(TEAL_PRIMARY)
        c.rect(table_x, y - header_height, col_width, header_height, fill=1, stroke=0)

        # Deductions header (right)
        c.rect(table_x + col_width, y - header_height, col_width, header_height, fill=1, stroke=0)

        # Header text
        c.setFillColor(WHITE)
        c.setFont('Helvetica-Bold', 10)
        c.drawString(table_x + 10, y - 17, "Earnings")
        c.drawString(table_x + col_width + 10, y - 17, "Deductions")

        y -= header_height

        # Draw data rows
        for i in range(max_rows):
            # Alternate row background
            if i % 2 == 0:
                c.setFillColor(LIGHT_GRAY)
            else:
                c.setFillColor(WHITE)

            c.rect(table_x, y - row_height, col_width, row_height, fill=1, stroke=0)
            c.rect(table_x + col_width, y - row_height, col_width, row_height, fill=1, stroke=0)

            # Draw earnings data - using solid black for clear visibility
            c.setFillColor(SOLID_BLACK)
            c.setFont('Helvetica', 9)
            if i < len(earnings_data):
                label, amount = earnings_data[i]
                c.drawString(table_x + 10, y - 15, label)
                c.drawRightString(table_x + col_width - 10, y - 15, self._format_amount(amount))

            # Draw deductions data - using solid black for clear visibility
            c.setFillColor(SOLID_BLACK)
            if i < len(deductions_data):
                label, amount = deductions_data[i]
                c.drawString(table_x + col_width + 10, y - 15, label)
                c.drawRightString(table_x + table_width - 10, y - 15, self._format_amount(amount))

            y -= row_height

        # Draw totals row
        c.setFillColor(colors.HexColor('#E0E0E0'))
        c.rect(table_x, y - row_height, col_width, row_height, fill=1, stroke=0)
        c.rect(table_x + col_width, y - row_height, col_width, row_height, fill=1, stroke=0)

        # Totals text - using solid black for clear visibility
        c.setFillColor(SOLID_BLACK)
        c.setFont('Helvetica-Bold', 9)
        c.drawString(table_x + 10, y - 15, "Total Earnings")
        c.drawRightString(table_x + col_width - 10, y - 15, self._format_amount(total_earnings))
        c.drawString(table_x + col_width + 10, y - 15, "Total Deductions")
        c.drawRightString(table_x + table_width - 10, y - 15, self._format_amount(total_deductions))

        y -= row_height

        # Draw table borders
        c.setStrokeColor(DARK_TEAL)
        c.setLineWidth(1)
        # Outer border - left table
        c.rect(table_x, y, col_width, y_start - y, fill=0, stroke=1)
        # Outer border - right table
        c.rect(table_x + col_width, y, col_width, y_start - y, fill=0, stroke=1)

        return y - 15, table_x, table_width

    def _draw_net_salary(self, c, record, y_pos):
        """Draw net salary box"""
        net_salary = self._parse_amount(self._get_field_value(record, ['Net Salary', 'net salary']))
        amount_paid = self._parse_amount(self._get_field_value(record, ['Amout Paid', 'amout paid', 'Amount Paid']))
        final_amount = amount_paid if amount_paid > 0 else net_salary

        box_x = 50
        box_width = PAGE_WIDTH - 100
        box_height = 30

        # Draw teal background (same color as table header)
        c.setFillColor(TEAL_PRIMARY)
        c.rect(box_x, y_pos - box_height, box_width, box_height, fill=1, stroke=0)

        # Draw text
        c.setFillColor(WHITE)
        c.setFont('Helvetica-Bold', 12)
        c.drawString(box_x + 15, y_pos - 20, "NET SALARY")
        c.drawRightString(box_x + box_width - 15, y_pos - 20, self._format_amount(final_amount))

        return y_pos - box_height - 15

    def _draw_company_stamp_section(self, c):
        """Draw company stamp section with line and rotated stamp just above Generated on"""
        if not os.path.exists(self.paid_stamp_path):
            return

        try:
            # Position just above Generated on
            line_y = 110
            line_x = 50
            line_width = 120

            # Draw line for company stamp
            c.setStrokeColor(SOLID_BLACK)
            c.setLineWidth(1)
            c.line(line_x, line_y, line_x + line_width, line_y)

            # Draw "Company Stamp" label below the line
            c.setFont('Helvetica-Bold', 9)
            c.setFillColor(SOLID_BLACK)
            c.drawString(line_x + 15, line_y - 12, "Company Stamp")

            # Open the stamp image
            stamp_img = Image.open(self.paid_stamp_path)

            # Convert to RGBA if not already
            if stamp_img.mode != 'RGBA':
                stamp_img = stamp_img.convert('RGBA')

            # Change stamp color to dark green (#073630)
            pixels = stamp_img.load()
            dark_green = (7, 54, 48)  # RGB for #073630

            for i in range(stamp_img.width):
                for j in range(stamp_img.height):
                    r, g, b, a = pixels[i, j]
                    # If pixel is not transparent, change to dark green
                    if a > 0:
                        # Calculate grayscale to maintain shading
                        gray = (r + g + b) // 3
                        # Apply dark green with original intensity
                        if gray < 128:  # Dark pixels (the stamp itself)
                            pixels[i, j] = (dark_green[0], dark_green[1], dark_green[2], a)
                        else:  # Light pixels (background)
                            pixels[i, j] = (r, g, b, 0)  # Make transparent

            # Rotate 40 degrees to the right (clockwise)
            stamp_img = stamp_img.rotate(-40, expand=True, resample=Image.BICUBIC)

            # Calculate stamp size
            stamp_height = 90
            aspect_ratio = stamp_img.width / stamp_img.height
            stamp_width = stamp_height * aspect_ratio

            # Position stamp on the line (centered on line)
            stamp_x = line_x + (line_width - stamp_width) / 2
            stamp_y = line_y - 15  # Place stamp a bit lower

            # Save to bytes buffer
            img_buffer = BytesIO()
            stamp_img.save(img_buffer, format='PNG')
            img_buffer.seek(0)

            # Draw the stamp
            img_reader = ImageReader(img_buffer)
            c.drawImage(img_reader, stamp_x, stamp_y, width=stamp_width, height=stamp_height, mask='auto')

        except Exception as e:
            print(f"Error drawing company stamp: {str(e)}")

    def _draw_stamp(self, c, table_x, table_width, y_pos):
        """Draw company stamp with transparency in center of page"""
        if not os.path.exists(self.stamp_path):
            return

        try:
            # Open the stamp image and add transparency
            stamp_img = Image.open(self.stamp_path)

            # Convert to RGBA if not already
            if stamp_img.mode != 'RGBA':
                stamp_img = stamp_img.convert('RGBA')

            # Create a new image with transparency
            # Very low opacity (8% opacity for very subtle watermark)
            alpha = stamp_img.split()[3] if stamp_img.mode == 'RGBA' else Image.new('L', stamp_img.size, 255)
            alpha = alpha.point(lambda p: int(p * 0.08))  # 8% opacity for very light watermark

            # Create new RGBA image with adjusted alpha
            stamp_img.putalpha(alpha)

            # Calculate stamp size (smaller - 60% of table width)
            stamp_width = table_width * 0.6
            aspect_ratio = stamp_img.height / stamp_img.width
            stamp_height = stamp_width * aspect_ratio

            # Position stamp in center, slightly right and up
            stamp_x = (PAGE_WIDTH - stamp_width) / 2 + 20  # Shifted 20 points to the right
            stamp_y = (PAGE_HEIGHT - stamp_height) / 2 + 12  # Shifted 12 points up

            # Save to bytes buffer
            img_buffer = BytesIO()
            stamp_img.save(img_buffer, format='PNG')
            img_buffer.seek(0)

            # Draw the stamp
            img_reader = ImageReader(img_buffer)
            c.drawImage(img_reader, stamp_x, stamp_y, width=stamp_width, height=stamp_height, mask='auto')

        except Exception as e:
            print(f"Error drawing stamp: {str(e)}")

    def _merge_with_letterhead(self, content_buffer, output_path):
        """Merge content PDF with letter head background"""
        # Read letter head PDF
        letterhead_reader = PdfReader(self.letter_head_path)
        letterhead_page = letterhead_reader.pages[0]

        # Read content PDF
        content_reader = PdfReader(content_buffer)
        content_page = content_reader.pages[0]

        # Merge content onto letterhead
        letterhead_page.merge_page(content_page)

        # Write output
        writer = PdfWriter()
        writer.add_page(letterhead_page)

        with open(output_path, 'wb') as output_file:
            writer.write(output_file)

    def _get_field_value(self, record, field_names):
        """Get value from record trying multiple field name variations"""
        if isinstance(field_names, str):
            field_names = [field_names]

        record_lower = {k.lower().strip(): v for k, v in record.items()}

        for field in field_names:
            if field in record and record[field]:
                return record[field]
            if field.lower() in record_lower and record_lower[field.lower()]:
                return record_lower[field.lower()]

        return None

    def _parse_amount(self, value):
        """Parse amount from various formats"""
        if value is None:
            return 0.0
        try:
            if isinstance(value, (int, float)):
                return float(value)
            clean_value = str(value).replace('PKR', '').replace('Rs', '').replace('Rs.', '')
            clean_value = clean_value.replace(',', '').replace('$', '').strip()
            if clean_value and clean_value not in ['0', '0.0', '0.00', '-', 'N/A', 'n/a', '']:
                return float(clean_value)
        except (ValueError, TypeError):
            pass
        return 0.0

    def _format_amount(self, value):
        """Format value as currency amount"""
        if value is None or value == 0:
            return "-"
        try:
            return f"{float(value):,.2f}"
        except (ValueError, TypeError):
            return "-"

    def get_pdf_filename(self, record, month_name):
        """Generate a filename for the PDF"""
        name = self._get_field_value(record, ['Name', 'name', 'Employee Name']) or ''

        if name:
            clean_name = "".join(c for c in str(name) if c.isalnum() or c in (' ', '-', '_')).strip()
            clean_name = clean_name.replace(' ', '_')
        else:
            clean_name = "Employee"

        clean_month = month_name.replace(' ', '_')

        cnic = self._get_field_value(record, ['CNIC', 'cnic']) or ''
        if cnic:
            clean_cnic = "".join(c for c in str(cnic) if c.isalnum() or c == '-')
            filename = f"{clean_month}_{clean_name}_{clean_cnic}.pdf"
        else:
            filename = f"{clean_month}_{clean_name}.pdf"

        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            filename = filename.replace(char, '_')

        return filename
