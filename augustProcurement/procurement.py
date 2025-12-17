import pandas as pd
import re
import os
from collections import defaultdict
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

def get_unicode_font():
    """
    Get a font that supports Chinese characters.
    Returns the font name to use in ReportLab.
    """
    # Common Chinese font paths on Windows
    font_paths = [
        r'C:\Windows\Fonts\simsun.ttc',  # SimSun
        r'C:\Windows\Fonts\simhei.ttf',  # SimHei
        r'C:\Windows\Fonts\msyh.ttc',    # Microsoft YaHei
        r'C:\Windows\Fonts\arialuni.ttf' # Arial Unicode MS
    ]
    
    for font_path in font_paths:
        if os.path.exists(font_path):
            try:
                # Register the font with ReportLab
                font_name = f"UnicodeFont_{os.path.basename(font_path).split('.')[0]}"
                pdfmetrics.registerFont(TTFont(font_name, font_path))
                print(f"Registered Unicode font: {font_name}")
                return font_name
            except Exception as e:
                print(f"Failed to register font {font_path}: {e}")
                continue
    
    # Fallback to Helvetica if no Chinese font is found
    print("No Chinese font found, using Helvetica")
    return "Helvetica"

def extract_all_products_corrected(excel_file):
    """
    Extract ALL product information from the Excel file, properly handling warehouse locations
    and combining products with the same name but different codes.
    """
    # Read the Excel file
    df = pd.read_excel(excel_file)
    df.columns = ['A', 'B', 'C']
    
    # Known warehouse locations to exclude as products
    warehouse_locations = {
        'Dandenong', 'Polyaire', 'Fourways', 'Mulgrave', 'DLZ', 'SPT', 'QLD'
    }
    
    products = {}
    product_name_to_codes = defaultdict(list)  # Track multiple codes for same product
    
    print("=== EXTRACTING ALL PRODUCTS (CORRECTED) ===")
    
    # First pass: Find all "Total:" lines and work backwards to find product info
    for i, row in df.iterrows():
        if pd.isna(row['B']):
            continue
            
        cell_b = str(row['B']).strip()
        cell_c = str(row['C']).strip() if pd.notna(row['C']) else ""
        
        # Check if this is a total line
        if 'Total:' in cell_b:
            # Extract product name from the total line
            product_name = cell_b.replace('Total:', '').strip()
            
            # Skip if this is a warehouse location
            if product_name in warehouse_locations:
                print(f"  Skipping warehouse location: {product_name}")
                continue
            
            try:
                quantity = int(float(cell_c)) if cell_c and cell_c != 'nan' else 0
            except (ValueError, TypeError):
                quantity = 0
                print(f"  Warning: Could not parse quantity '{cell_c}' for {cell_b}")
            
            # Look backwards to find the product code (if any)
            product_code = None
            product_full_name = product_name
            
            # Special case: Check if this is the "Spare Parts" Total line
            if product_name.strip() == '' and 'Total:' in cell_b:
                # Look for "Spare Parts" in the nearby rows (prioritize this check)
                spare_parts_found = False
                for j in range(max(0, i-10), i):
                    if pd.notna(df.iloc[j]['B']) and str(df.iloc[j]['B']).strip() == 'Spare Parts':
                        product_code = "Spare Parts"
                        product_full_name = "Spare Parts"
                        spare_parts_found = True
                        print(f"  Found Spare Parts at row {j}, setting product for row {i}")
                        break
                
                if not spare_parts_found:
                    # Continue with regular product code search
                    product_code = None
            
            # If not Spare Parts, look for regular product codes
            if product_code is None:
                # Look back up to 15 rows for a product code or product identifier
                for j in range(i-1, max(0, i-15), -1):
                    if pd.notna(df.iloc[j]['B']):
                        prev_cell_b = str(df.iloc[j]['B']).strip()
                        prev_cell_c = str(df.iloc[j]['C']).strip() if pd.notna(df.iloc[j]['C']) else ""
                        
                        # Skip supplier/location/warehouse entries
                        if (prev_cell_b in warehouse_locations or
                            'Warehouse' in prev_cell_b or
                            'Midea Electric Trading' in prev_cell_b or
                            'MIDEA ELECTRONICS AUSTRALIA CO PTY LTD' in prev_cell_b):
                            continue
                        
                        # Check if this looks like a product code or identifier
                        if (re.match(r'^\d{11,}$', prev_cell_b) or  # Long numeric codes
                            re.match(r'^[A-Z0-9\-]{4,}$', prev_cell_b) or  # Alphanumeric codes
                            (len(prev_cell_b) >= 2 and prev_cell_b.replace('-', '').replace('_', '').isalnum())):
                            
                            product_code = prev_cell_b
                            # Use the product name from the code row if available and makes sense
                            if prev_cell_c and prev_cell_c != 'nan' and len(prev_cell_c) > 2:
                                product_full_name = prev_cell_c
                            else:
                                # If no product name found, leave it empty
                                product_full_name = ""
                            break
                        
                        # Special case: if we find a descriptive name that matches our product
                        elif (prev_cell_c and prev_cell_c != 'nan' and 
                              product_name.lower() in prev_cell_c.lower()):
                            product_code = prev_cell_b if len(prev_cell_b) > 1 else None
                            product_full_name = prev_cell_c
                            break
            
            # Handle special cases for products without clear codes
            if not product_code:
                # Check if the product name itself could be a code (like FQH-03A)
                if re.match(r'^[A-Z0-9\-]{3,}$', product_name):
                    product_code = product_name
                    product_full_name = product_name
                elif product_name.strip() == '':
                    # Empty product name with just "Total:" - look for context
                    # Check if there's a "Spare Parts" entry nearby (look more specifically)
                    spare_parts_found = False
                    for j in range(max(0, i-5), i):
                        if pd.notna(df.iloc[j]['B']) and str(df.iloc[j]['B']).strip() == 'Spare Parts':
                            product_name = "Spare Parts"
                            product_code = "Spare Parts"
                            product_full_name = "Spare Parts"
                            spare_parts_found = True
                            break
                    
                    if not spare_parts_found:
                        product_name = "Unknown Product"
                        product_code = "N/A"
                else:
                    product_code = "N/A"
            
            # Normalize product name for grouping
            # If we found a product name from the code row, use it; otherwise use the name from total line
            if product_full_name:
                normalized_name = product_full_name.strip()
            else:
                # If the product name from the total line is just "Total:", leave it empty
                if product_name.strip() == "Total:":
                    normalized_name = ""
                else:
                    normalized_name = product_name.strip()
            
            # Track codes for this product name
            product_name_to_codes[normalized_name].append(product_code)
            
            # Use normalized name as key, but track all codes for this product
            product_key = normalized_name
            
            # Add to products dictionary
            if product_key not in products:
                products[product_key] = {
                    'name': normalized_name,
                    'codes': set(),
                    'totals': [],
                    'total_quantity': 0
                }
            
            # Add the code to the set
            if product_code:
                products[product_key]['codes'].add(product_code)
            
            products[product_key]['totals'].append(quantity)
            print(f"Found: {normalized_name} (Code: {product_code or 'N/A'}) -> +{quantity} (Row {i})")
    
    # Second pass: Process each product entry
    for product_key, info in products.items():
        info['total_quantity'] = sum(info['totals'])
        
        # Convert codes set to a clean list, removing duplicates and N/A if other codes exist
        codes_list = list(info['codes'])
        if len(codes_list) > 1 and 'N/A' in codes_list:
            codes_list.remove('N/A')
        
        # Use the most specific code (longest one) as primary
        if codes_list:
            info['primary_code'] = max(codes_list, key=len)
        else:
            info['primary_code'] = 'N/A'
        
        # Store all codes for reference
        info['all_codes'] = ', '.join(sorted(codes_list)) if codes_list else 'N/A'
        
        print(f"Product {info['name']} (Codes: {info['all_codes']}): {len(info['totals'])} totals = {info['total_quantity']}")
    
    return products

def create_pdf_report_corrected(products, output_file, source_filename):
    """
    Create a PDF report with products sorted alphabetically.
    """
    doc = SimpleDocTemplate(output_file, pagesize=A4)
    elements = []
    
    # Get Unicode font for Chinese characters
    unicode_font = get_unicode_font()
    
    # Styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=16,
        spaceAfter=30,
        alignment=1,  # Center alignment
        fontName=unicode_font
    )
    
    # Title with filename
    import os
    filename_display = os.path.basename(source_filename)
    title_text = f"采购总结 Product Procurement Summary - August 2025<br/><font size='12'>Source: {filename_display}</font>"
    title = Paragraph(title_text, title_style)
    elements.append(title)
    elements.append(Spacer(1, 20))
    
    # Sort products alphabetically by name
    sorted_products = sorted(products.items(), key=lambda x: x[1]['name'].lower())
    
    # Create table data
    table_data = [['Product Code(s)', 'Product Name', 'Total Quantity']]
    
    for key, info in sorted_products:
        table_data.append([
            info['all_codes'],
            info['name'],
            str(info['total_quantity'])
        ])
    
    # Create table
    table = Table(table_data, colWidths=[2.2*inch, 3.3*inch, 1*inch])
    
    # Table style
    table.setStyle(TableStyle([
        # Header row
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), unicode_font),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        
        # Data rows
        ('FONTNAME', (0, 1), (-1, -1), unicode_font),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        
        # Alternating row colors
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
    ]))
    
    elements.append(table)
    
    # Add summary
    elements.append(Spacer(1, 30))
    summary_style = ParagraphStyle(
        'SummaryStyle',
        parent=styles['Normal'],
        fontName=unicode_font
    )
    total_products = len(products)
    total_quantity = sum(info['total_quantity'] for info in products.values())
    
    summary_text = f"""
    
    """
    
    summary = Paragraph(summary_text, summary_style)
    elements.append(summary)
    
    # Build PDF
    doc.build(elements)
    print(f"PDF report saved as: {output_file}")

def main():
    excel_file = '八月采购.xlsx'
    output_pdf = 'Product_Procurement_Summary.pdf'
    
    try:
        # Extract product data
        products = extract_all_products_corrected(excel_file)
        
        print(f"\n=== SUMMARY ===")
        print(f"Total products found: {len(products)}")
        
        # Display products for verification
        print("\n=== CORRECTED PRODUCT LIST (Alphabetical) ===")
        sorted_products = sorted(products.items(), key=lambda x: x[1]['name'].lower())
        for i, (key, info) in enumerate(sorted_products, 1):
            print(f"{i:2d}. {info['name']:<40} (Codes: {info['all_codes']:<15}) - Qty: {info['total_quantity']}")
        
        # Check for specific products mentioned
        print(f"\n=== CHECKING SPECIFIC PRODUCTS ===")
        mini_vrf_28_found = [p for p in products.values() if 'Mini VRF 2.8kw IDU' in p['name']]
        if mini_vrf_28_found:
            for product in mini_vrf_28_found:
                print(f"Mini VRF 2.8kw IDU: {product['all_codes']} - Qty: {product['total_quantity']} (from {len(product['totals'])} totals)")
        
        # Create PDF report
        create_pdf_report_corrected(products, output_pdf, excel_file)
        
        print(f"\n=== COMPLETED ===")
        print(f"Extracted {len(products)} products")
        print(f"PDF report saved as: {output_pdf}")
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
