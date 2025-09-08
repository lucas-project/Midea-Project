import pandas as pd
import numpy as np
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import re
import os

def read_excel_file(file_path):
    """Read the Excel file and return DataFrame"""
    try:
        # Try reading with different engines
        df = pd.read_excel(file_path, engine='xlrd')
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

def extract_product_data(df):
    """Extract product codes, names, and quantities from the DataFrame"""
    products = {}
    current_product_code = None
    current_product_name = None
    
    # Iterate through each row
    for index, row in df.iterrows():
        col_b = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ""  # Column B
        col_c = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""  # Column C
        col_d = str(row.iloc[3]) if pd.notna(row.iloc[3]) else ""  # Column D
        col_e = str(row.iloc[4]) if pd.notna(row.iloc[4]) else ""  # Column E
        
        # Check if this row contains a product code and name
        # Product codes can be numeric (like 12127000001896) or alphanumeric (like CASG-XG70, E88, MDV-V235WN1(AU)-A)
        # They appear in column B with corresponding product name in column C
        if col_b and col_c and col_b != 'nan' and col_c != 'nan':
            # Skip header rows and location rows
            if ('Name' not in col_d and 'Quantity' not in col_e and 
                'Warehouse' not in col_c and 'Rd' not in col_c and 
                'ROAD' not in col_b and 'PTY LTD' not in col_b and
                'August' not in col_b and 'Sales' not in col_b and
                'Weddel Court' not in col_c and
                col_b not in ['QLD', 'SYD', 'NSW', 'VIC', 'WA', 'SA', 'TAS', 'NT', 'ACT'] and  # Skip state codes
                'WAREHOUSE' not in col_c.upper()):  # Skip warehouse entries
                
                # More inclusive product code detection
                # Accept if: length >= 2, starts with alphanumeric, and not purely alphabetic (unless very short like E88)
                is_product_code = False
                
                if len(col_b) >= 2:
                    # Check various product code patterns
                    if (col_b[0].isalnum() and 
                        (not col_b.isalpha() or len(col_b) <= 4)):  # Allow short codes like E88
                        # Additional checks for common product patterns
                        if (col_b.isdigit() or  # Pure numbers
                            '-' in col_b or    # Contains dashes
                            '(' in col_b or    # Contains parentheses  
                            any(c.isdigit() for c in col_b) or  # Contains at least one digit
                            (len(col_b) <= 4 and col_b.isalpha())):  # Short alpha codes like E88
                            is_product_code = True
                
                if is_product_code:
                    current_product_code = col_b
                    current_product_name = col_c
                    
                    # Initialize product in dictionary if not exists
                    if current_product_code not in products:
                        products[current_product_code] = {
                            'name': current_product_name,
                            'total_quantity': 0
                        }
                        print(f"Found product: {current_product_code} - {current_product_name}")
        
        # Check if this row contains a total (contains 'Total:' in column D)
        if col_d and 'Total:' in col_d and current_product_code:
            try:
                quantity = float(col_e) if col_e and col_e != 'nan' and col_e != '' else 0
                products[current_product_code]['total_quantity'] += quantity
                print(f"  Added total for {current_product_code}: {col_d} = {quantity}")
            except (ValueError, TypeError):
                print(f"  Could not parse quantity: {col_e}")
    
    return products

def register_chinese_fonts():
    """Register Chinese fonts for proper Unicode support"""
    try:
        # Try to register common Chinese fonts available on Windows
        font_paths = [
            "C:/Windows/Fonts/simsun.ttc",  # SimSun (宋体)
            "C:/Windows/Fonts/simhei.ttf",  # SimHei (黑体)
            "C:/Windows/Fonts/msyh.ttc",    # Microsoft YaHei (微软雅黑)
            "C:/Windows/Fonts/simkai.ttf",  # SimKai (楷体)
        ]
        
        for font_path in font_paths:
            if os.path.exists(font_path):
                try:
                    # Register regular font
                    pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
                    # Register bold font (using the same font file for both)
                    pdfmetrics.registerFont(TTFont('ChineseFont-Bold', font_path))
                    print(f"Successfully registered Chinese font: {font_path}")
                    return True
                except Exception as e:
                    print(f"Failed to register font {font_path}: {e}")
                    continue
        
        print("No Chinese fonts found, using default font")
        return False
    except Exception as e:
        print(f"Error registering Chinese fonts: {e}")
        return False

def create_pdf_report(products, excel_filename):
    """Create PDF report with title page and product summary"""
    # Register Chinese fonts
    chinese_font_available = register_chinese_fonts()
    
    # Get current date and time
    current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Extract base filename without extension
    base_filename = os.path.splitext(excel_filename)[0]
    
    # Create PDF filename with timestamp to avoid conflicts
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    pdf_filename = f"{base_filename}_Sale_Summary_Report_{timestamp}.pdf"
    
    # Create document
    doc = SimpleDocTemplate(pdf_filename, pagesize=A4)
    styles = getSampleStyleSheet()
    story = []
    
    # Title page styles with Chinese font support
    font_name = 'ChineseFont' if chinese_font_available else 'Helvetica'
    
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        spaceAfter=30,
        alignment=TA_CENTER,
        fontName=font_name
    )
    
    subtitle_style = ParagraphStyle(
        'CustomSubtitle',
        parent=styles['Heading2'],
        fontSize=18,
        spaceAfter=20,
        alignment=TA_CENTER,
        fontName=font_name
    )
    
    date_style = ParagraphStyle(
        'DateStyle',
        parent=styles['Normal'],
        fontSize=12,
        alignment=TA_CENTER,
        spaceAfter=20,
        fontName=font_name
    )
    
    # Add title page content
    story.append(Paragraph(f"{base_filename} Sale Summary Report", title_style))
    story.append(Spacer(1, 50))
    story.append(Paragraph("Sales Analysis Report", subtitle_style))
    story.append(Spacer(1, 30))
    story.append(Paragraph(f"Generated on: {current_datetime}", date_style))
    
    # Page break
    story.append(PageBreak())
    
    # Prepare table data
    if products:
        # Sort products alphabetically by name
        sorted_products = sorted(products.items(), 
                               key=lambda x: x[1]['name'] if x[1]['name'] else x[0])
        
        # Create table data
        table_data = [['Product Code', 'Product Name', 'Total Quantity']]
        
        for product_code, product_info in sorted_products:
            table_data.append([
                product_code,
                product_info['name'] if product_info['name'] else 'N/A',
                f"{product_info['total_quantity']:,.0f}"
            ])
        
        # Create table
        table = Table(table_data, colWidths=[2*inch, 3*inch, 1.5*inch])
        
        # Apply table style with alternating row colors and Chinese font support
        table_style = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), f'{font_name}-Bold' if chinese_font_available else 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('FONTNAME', (0, 1), (-1, -1), font_name),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ])
        
        # Add alternating row colors
        for i in range(1, len(table_data)):
            if i % 2 == 0:
                table_style.add('BACKGROUND', (0, i), (-1, i), colors.lightgrey)
        
        table.setStyle(table_style)
        
        # Add table title
        table_title_style = ParagraphStyle(
            'TableTitle',
            parent=styles['Heading2'],
            fontSize=16,
            spaceAfter=20,
            alignment=TA_CENTER,
            fontName=font_name
        )
        
        story.append(Paragraph("Product Summary", table_title_style))
        story.append(Spacer(1, 20))
        story.append(table)
        
        # Add summary statistics
        total_products = len(products)
        total_quantity = sum(p['total_quantity'] for p in products.values())
        
        story.append(Spacer(1, 30))
        summary_style = ParagraphStyle(
            'Summary',
            parent=styles['Normal'],
            fontSize=12,
            alignment=TA_LEFT,
            fontName=font_name
        )
        
        story.append(Paragraph(f"<b>Summary Statistics:</b>", summary_style))
        story.append(Paragraph(f"Total Products: {total_products}", summary_style))
        story.append(Paragraph(f"Total Quantity: {total_quantity:,.0f}", summary_style))
    
    else:
        story.append(Paragraph("No product data found in the Excel file.", styles['Normal']))
    
    # Build PDF
    doc.build(story)
    print(f"PDF report generated: {pdf_filename}")
    return pdf_filename

def main():
    excel_file = "八月销售.xls"
    
    if not os.path.exists(excel_file):
        print(f"Excel file '{excel_file}' not found!")
        return
    
    print("Reading Excel file...")
    df = read_excel_file(excel_file)
    
    if df is None:
        print("Failed to read Excel file.")
        return
    
    print(f"Excel file loaded. Shape: {df.shape}")
    print("\nFirst few rows:")
    print(df.head())
    
    print("\nExtracting product data...")
    products = extract_product_data(df)
    
    print(f"\nFound {len(products)} products:")
    for code, info in products.items():
        print(f"  {code}: {info['name']} - Quantity: {info['total_quantity']}")
    
    print("\nGenerating PDF report...")
    pdf_filename = create_pdf_report(products, excel_file)
    
    print(f"Process completed! PDF saved as: {pdf_filename}")

if __name__ == "__main__":
    main()
