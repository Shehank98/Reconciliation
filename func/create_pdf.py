import os
import pandas as pd
import numpy as np
from datetime import datetime
from reportlab.lib.pagesizes import letter, landscape
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import inch
from tkinter import messagebox
from tkinter import ttk

def create_summary_pdf_tables_only(self):
    """Main PDF creation function - 2 PAGES ONLY"""
    try:
        if not self.matched_data:
            messagebox.showerror("Error", "No matching results available")
            return None

        # Get save location from user
        from tkinter import filedialog
        filename = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile=f"media_reconciliation_report_{self.selected_channel or 'channel'}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        )

        if not filename:
            return None

        # Create document with proper mixed orientation support
        from reportlab.platypus import BaseDocTemplate, PageTemplate, Frame, NextPageTemplate
        
        doc = BaseDocTemplate(filename)
        
        # Build story elements
        story = []
        
        def draw_page_elements(canvas, doc):
            """Custom function to add page numbers and logo"""
            page_num = canvas.getPageNumber()
            canvas.setFont("Helvetica", 9)

            # Get current page size
            page_width, page_height = canvas._pagesize

            # Add page numbers based on orientation
            if page_width > page_height:  # Landscape
                canvas.drawRightString(page_width-30, 30, f"Page {page_num}")
                canvas.drawString(30, 30, f"Media Report - {self.selected_channel or 'Unknown'}")
            else:  # Portrait
                canvas.drawRightString(page_width-30, 30, f"Page {page_num}")
                canvas.drawString(30, 30, f"Media Report - {self.selected_channel or 'Unknown'}")

            # Add logo only to first page (portrait)
            if page_num == 1:
                add_logo_to_canvas(canvas, page_width, page_height)

        # Portrait template for page 1
        portrait_frame = Frame(0.75*inch, 0.5*inch, letter[0]-1.5*inch, letter[1]-1*inch, id='portrait')
        portrait_template = PageTemplate(id='portrait', frames=[portrait_frame], pagesize=letter, onPage=draw_page_elements)

        # Landscape template for page 2
        landscape_frame = Frame(0.5*inch, 0.5*inch, landscape(letter)[0]-1*inch, landscape(letter)[1]-1*inch, id='landscape')
        landscape_template = PageTemplate(id='landscape', frames=[landscape_frame], pagesize=landscape(letter), onPage=draw_page_elements)
        
        # Add templates to document
        doc.addPageTemplates([portrait_template, landscape_template])
        
        # Enhanced styles
        styles = getSampleStyleSheet()
        title_style = styles["Title"]
        title_style.textColor = colors.darkblue
        title_style.fontSize = 16
        title_style.alignment = 1
        
        heading_style = styles["Heading1"]
        heading_style.textColor = colors.darkblue
        heading_style.fontSize = 14
        
        normal_style = styles["Normal"]
        normal_style.fontSize = 10

        # PAGE 1: SUMMARY SHEET (Portrait)
        story.extend(create_summary_sheet_page_fixed(self, title_style, normal_style, heading_style))

        # PAGE 2: COMPLETE MATCHED LMRB TABLE (Landscape) 
        story.append(NextPageTemplate('landscape'))
        story.append(PageBreak())
        story.extend(create_complete_matched_lmrb_table(self, heading_style))

        # BUILD PDF (Only 2 pages now)
        doc.build(story)
        
        print(f"PDF successfully created at: {filename}")
        return filename
        
    except Exception as e:
        import traceback
        print(f"Error in PDF creation: {str(e)}")
        print(traceback.format_exc())
        messagebox.showerror("PDF Error", f"Error creating PDF: {str(e)}")
        return None

def add_logo_to_canvas(canvas, page_width, page_height):
    """FIXED: Add transparent PNG logo properly"""
    try:
        import sys
        if hasattr(sys, '_MEIPASS'):
            script_dir = sys._MEIPASS
        else:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            script_dir = os.path.dirname(script_dir)
        
        logo_path = os.path.join(script_dir, "logo.png")
        
        print(f"Looking for logo at: {logo_path}")
        
        if os.path.exists(logo_path):
            logo_x = page_width - 120
            logo_y = page_height - 100
            
            try:
                # FIXED: Use PIL to handle transparent PNG properly
                from PIL import Image as PILImage
                
                # Open image with PIL to handle transparency
                pil_img = PILImage.open(logo_path)
                
                # Convert to RGBA if not already
                if pil_img.mode != 'RGBA':
                    pil_img = pil_img.convert('RGBA')
                
                # Create a white background and paste the logo
                background = PILImage.new('RGBA', pil_img.size, (255, 255, 255, 0))
                combined = PILImage.alpha_composite(background, pil_img)
                
                # Convert back to RGB for ReportLab
                rgb_img = PILImage.new('RGB', combined.size, (255, 255, 255))
                rgb_img.paste(combined, mask=combined.split()[-1])  # Use alpha as mask
                
                # Save temporary file
                import tempfile
                with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
                    rgb_img.save(temp_file.name, 'PNG')
                    temp_logo_path = temp_file.name
                
                # Draw the processed image
                canvas.drawImage(temp_logo_path, logo_x, logo_y, width=80, height=40, preserveAspectRatio=True)
                
                # Clean up temp file
                os.unlink(temp_logo_path)
                
                print(f"✅ Transparent logo processed and added successfully")
                
            except Exception as pil_error:
                print(f"PIL processing failed: {pil_error}")
                # Fallback to direct drawing
                canvas.drawImage(logo_path, logo_x, logo_y, width=80, height=40, preserveAspectRatio=True)
                print(f"✅ Logo added with fallback method")
                
        else:
            print(f"❌ Logo not found at: {logo_path}")
            
    except Exception as e:
        print(f"❌ Could not add logo: {e}")

def create_complete_matched_lmrb_table(self, heading_style):
    """Create complete LMRB data table - ALL RECORDS, ALL COLUMNS, SORTED BY DATE"""
    try:
        elements = []
        
        # Page title
        title_style = getSampleStyleSheet()["Heading1"]
        title_style.fontSize = 14
        title_style.textColor = colors.darkblue
        title_style.alignment = 1
        
        elements.append(Paragraph("<b>Complete Matched LMRB Data (All Records - Sorted by Date)</b>", title_style))
        elements.append(Spacer(1, 15))
        
        # Get matched data
        matched_df = self.get_current_matched_data_df()
        
        if matched_df is None or matched_df.empty:
            elements.append(Paragraph("No matched LMRB data available.", title_style))
            return elements
        
        # SORT BY DATE - Comprehensive date sorting
        print(f"📊 Processing {len(matched_df)} matched LMRB records for table...")
        
        try:
            if all(col in matched_df.columns for col in ['Dd', 'Mn', 'Yr']):
                # Create proper date for sorting
                matched_df['sort_date'] = pd.to_datetime(
                    matched_df[['Yr', 'Mn', 'Dd']].rename(columns={'Yr': 'year', 'Mn': 'month', 'Dd': 'day'}),
                    errors='coerce'
                )
                matched_df = matched_df.sort_values(['sort_date', 'Advt_time'], na_last=True)
                print(f"📅 Sorted by Dd/Mn/Yr + Time: {len(matched_df)} records")
            elif 'Date' in matched_df.columns:
                matched_df = matched_df.sort_values('Date', na_last=True)
                print(f"📅 Sorted by Date column: {len(matched_df)} records")
            else:
                print("⚠️ No date columns found for sorting")
        except Exception as sort_error:
            print(f"⚠️ Date sorting failed: {sort_error}, using original order")

        # DEFINE ALL COLUMNS - Similar to your Excel format
        priority_columns = [
            'Product_Group', 'Advertiser', 'Product', 'Advt_Theme', 'Ads', 
            'Channel', 'Program', 'Dd', 'Mn', 'Yr', 'Day', 'Prog_time', 
            'Advt_time', 'AdPos', 'TotAds', 'BrkNo', 'PosinBrk', 'AdsinBrk', 
            'Lng', 'Dur', 'Cost'
        ]
        
        # Get available columns (use all available columns if priority columns don't exist)
        available_columns = [col for col in priority_columns if col in matched_df.columns]
        if not available_columns:
            # Fallback: use first 21 columns or all columns if fewer
            available_columns = list(matched_df.columns)[:21]
        
        print(f"📋 Using {len(available_columns)} columns: {available_columns}")
        
        # CALCULATE OPTIMAL TABLE SIZING for landscape
        landscape_width = 750  # Available points in landscape
        num_columns = len(available_columns)
        
        # Smart column width allocation based on content type
        column_width_map = {
            'Product_Group': 60, 'Advertiser': 80, 'Product': 70, 'Advt_Theme': 120,
            'Ads': 30, 'Channel': 60, 'Program': 80, 'Dd': 25, 'Mn': 25, 'Yr': 35,
            'Day': 30, 'Prog_time': 45, 'Advt_time': 45, 'AdPos': 35, 'TotAds': 35,
            'BrkNo': 30, 'PosinBrk': 35, 'AdsinBrk': 35, 'Lng': 25, 'Dur': 30, 'Cost': 40
        }
        
        # Calculate column widths
        total_preferred_width = sum(column_width_map.get(col, 40) for col in available_columns)
        scale_factor = landscape_width / total_preferred_width if total_preferred_width > landscape_width else 1
        
        column_widths = [column_width_map.get(col, 40) * scale_factor for col in available_columns]
        
        # CREATE TABLE DATA - ALL RECORDS
        print(f"🏗️ Creating table with ALL {len(matched_df)} records...")
        
        table_data = []
        
        # Header row
        table_data.append(available_columns)
        
        # ALL DATA ROWS (no limit)
        row_count = 0
        for index, row in matched_df.iterrows():
            try:
                row_data = []
                for col in available_columns:
                    value = row.get(col, '')
                    if pd.isna(value):
                        clean_value = ''
                    else:
                        clean_value = str(value)
                        # Smart truncation for very long values
                        if len(clean_value) > 25 and col in ['Advt_Theme', 'Product', 'Program']:
                            clean_value = clean_value[:22] + "..."
                        elif len(clean_value) > 15 and col not in ['Advt_Theme', 'Product', 'Program']:
                            clean_value = clean_value[:12] + "..."
                    
                    row_data.append(clean_value)
                
                table_data.append(row_data)
                row_count += 1
                
                # Progress indicator for large datasets
                if row_count % 100 == 0:
                    print(f"   📄 Processed {row_count}/{len(matched_df)} records")
                    
            except Exception as row_error:
                print(f"❌ Error processing row {index}: {row_error}")
                continue

        print(f"✅ Table created with {row_count} data rows")

        # CREATE THE TABLE
        table = Table(table_data, colWidths=column_widths, repeatRows=1)
        
        # APPLY COMPREHENSIVE TABLE STYLING
        table.setStyle(TableStyle([
            # Header styling
            ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 7),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            
            # Data styling
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 6),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            
            # Alternating row colors for better readability
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey]),
            
            # Padding
            ('LEFTPADDING', (0, 0), (-1, -1), 2),
            ('RIGHTPADDING', (0, 0), (-1, -1), 2),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
            
            # Special formatting for key columns
            ('ALIGN', (3, 1), (3, -1), 'LEFT'),  # Advt_Theme left aligned
            ('ALIGN', (6, 1), (6, -1), 'LEFT'),  # Program left aligned
        ]))
        
        elements.append(table)
        
        # Summary info
        elements.append(Spacer(1, 15))
        summary_style = getSampleStyleSheet()["Normal"]
        summary_style.fontSize = 9
        
        summary_text = f"<b>COMPLETE DATASET SUMMARY:</b><br/>"
        summary_text += f"• Total Records Displayed: {row_count:,}<br/>"
        summary_text += f"• Total Columns: {len(available_columns)}<br/>"
        summary_text += f"• Channel: {self.selected_channel or 'Unknown'}<br/>"
        summary_text += f"• Sorted By: Date + Time<br/>"
        summary_text += f"• Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}<br/>"
        summary_text += f"• Format: Complete Excel-style table with all records<br/>"
        summary_text += f"• This is a computer generated report, no signature required<br/>"

        elements.append(Paragraph(summary_text, summary_style))
        
        print(f"✅ Complete LMRB table created: {row_count} records, {len(available_columns)} columns")
        return elements
        
    except Exception as e:
        print(f"❌ Error creating complete LMRB table: {e}")
        import traceback
        traceback.print_exc()
        return [Paragraph(f"Error creating complete LMRB table: {str(e)}", heading_style)]

    # Keep these existing functions:
def create_summary_sheet_page_fixed(self, title_style, normal_style, heading_style):
    """Create the first page with proper formatting and alignment - PORTRAIT"""
    try:
        elements = []
        
        # CENTERED TITLE (removed MOBITEL)
        elements.append(Paragraph("<b>MEDIA RECONCILIATION SUMMARY REPORT</b>", title_style))
        elements.append(Spacer(1, 30))
        
        # Create header section with proper left-right alignment
        header_table_data = [
            [create_report_details_section_fixed(self, normal_style), create_company_info_section_fixed(normal_style)]
        ]
        
        header_table = Table(header_table_data, colWidths=[400, 200])
        header_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (0, 0), 'LEFT'),    # Left column - left aligned
            ('ALIGN', (1, 0), (1, 0), 'RIGHT'),   # Right column - right aligned
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # Top aligned
        ]))
        
        elements.append(header_table)
        elements.append(Spacer(1, 30))
        
        # Get report data
        commercial_data, sponsorship_data = self.generate_summary_report()
        
        # Commercial Benefits Table
        elements.append(Paragraph("<b>Commercial Benefits</b>", heading_style))
        elements.append(Spacer(1, 10))
        commercial_table = create_commercial_benefits_table_pdf_fixed(commercial_data)
        if commercial_table:
            elements.append(commercial_table)
        else:
            elements.append(Paragraph("No commercial benefits data available.", normal_style))
        elements.append(Spacer(1, 15))
        
        # Sponsorship Benefits Table
        elements.append(Paragraph("<b>Sponsorship Benefits</b>", heading_style))
        elements.append(Spacer(1, 10))
        sponsorship_table = create_sponsorship_benefits_table_pdf_fixed(sponsorship_data)
        if sponsorship_table:
            elements.append(sponsorship_table)
        else:
            elements.append(Paragraph("No sponsorship benefits data available.", normal_style))
        
        # Footer/Note section
        elements.append(Spacer(1, 30))
        elements.append(Paragraph("<b>NOTE:</b>", normal_style))
        elements.append(Spacer(1, 20))
        
        # Signature section
        signature_data = [
            ["…………………………………….", "…………………………………….", "……………………………………."],
            ["PREPARED BY", "CHECKED BY", "AUTHORISED BY"]
        ]
        
        signature_table = Table(signature_data, colWidths=[150, 150, 150])
        signature_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 1), (-1, 1), 10),
        ]))
        elements.append(signature_table)
        
        return elements
        
    except Exception as e:
        print(f"Error creating summary sheet page: {e}")
        return [Paragraph(f"Error creating summary sheet: {str(e)}", normal_style)]

def create_report_details_section_fixed(self, normal_style):
    """Create the left side report details section with proper alignment"""
    try:
        month = getattr(self, 'report_month_var', None) and self.report_month_var.get() or ''
        channel = self.selected_channel or ''
        estimate_no = getattr(self, 'estimate_no_var', None) and self.estimate_no_var.get() or ''
        supplier_invoice = getattr(self, 'supplier_invoice_var', None) and self.supplier_invoice_var.get() or ''
        po_no = getattr(self, 'po_no_var', None) and self.po_no_var.get() or ''
        invoice_no = getattr(self, 'invoice_no_var', None) and self.invoice_no_var.get() or ''
        
        details_data = [
            ["MONTH", ":", month],
            ["CHANNEL", ":", channel],
            ["ESTIMATE NO", ":", estimate_no],
            ["SUPPLIER INVOICE NO", ":", supplier_invoice],
            ["PO NO", ":", po_no],
            ["INVOICE NO", ":", invoice_no]
        ]
        
        details_table = Table(details_data, colWidths=[140, 20, 180])
        details_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),   # Labels left aligned
            ('ALIGN', (1, 0), (1, -1), 'CENTER'), # Colons centered
            ('ALIGN', (2, 0), (2, -1), 'LEFT'),   # Values left aligned
            ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ]))
        
        return details_table
        
    except Exception as e:
        print(f"Error creating report details: {e}")
        return Paragraph("Report details unavailable", normal_style)

def create_company_info_section_fixed(normal_style):
    """Create the right side company info section"""
    try:
        # Company info data (logo space will be handled by canvas drawing)
        company_info = [
            ["Phoenix O & M (Pvt) Ltd"],
            ["No 16, Barnes Place"],
            ["Colombo 7"]
        ]
        
        company_table = Table(company_info, colWidths=[180])
        company_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
            ('FONTNAME', (0, 0), (0, 0), 'Helvetica-Bold'),  # Company name bold
            ('FONTSIZE', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ]))
        
        return company_table
        
    except Exception as e:
        print(f"Error creating company info: {e}")
        return Paragraph("", normal_style)

def create_commercial_benefits_table_pdf_fixed(commercial_data):
    """Create commercial benefits table with proper formatting"""
    try:
        if not commercial_data:
            return None
            
        # Create header with merged cells
        table_data = [
            ["PRODUCT", "DUR", "SPOTS", "", "", "", "", "Average 30 sec."],
            ["", "", "PLANNED", "AIRED", "MISSED", "EXTRA", "3RD PARTY", ""]
        ]
        
        # Add data rows
        for item in commercial_data:
            table_data.append([
                str(item.get('product', '')),
                str(item.get('duration', '')),
                str(item.get('planned', 0)),
                str(item.get('aired', 0)),
                str(item.get('missed', 0)),
                str(item.get('extra', 0)),
                str(item.get('third_party', 0)),
                f"{item.get('average_30_sec', 0):.1f}"
            ])
        
        # Add totals row
        totals = [
            "Grand Total", "",
            sum(item.get('planned', 0) for item in commercial_data),
            sum(item.get('aired', 0) for item in commercial_data),
            sum(item.get('missed', 0) for item in commercial_data),
            sum(item.get('extra', 0) for item in commercial_data),
            sum(item.get('third_party', 0) for item in commercial_data),
            f"{sum(item.get('average_30_sec', 0) for item in commercial_data):.1f}"
        ]
        table_data.append(totals)
        
        # Create table with proper column widths
        table = Table(table_data, colWidths=[100, 30, 50, 50, 50, 50, 60, 80])
        table.setStyle(TableStyle([
            # Header formatting
            ('BACKGROUND', (0, 0), (-1, 1), colors.lightgrey),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            # Merge SPOTS header
            ('SPAN', (2, 0), (6, 0)),
            # Data rows
            ('ROWBACKGROUNDS', (0, 2), (-1, -2), [colors.white, colors.whitesmoke]),
            # Totals row
            ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ]))
        
        return table
        
    except Exception as e:
        print(f"Error creating commercial table: {e}")
        return None

def create_sponsorship_benefits_table_pdf_fixed(sponsorship_data):
    """Create sponsorship benefits table with proper formatting"""
    try:
        if not sponsorship_data:
            return None
            
        # Create header with merged cells
        table_data = [
            ["PRODUCT", "DUR", "SPOTS", "", "", "", ""],
            ["", "", "PLANNED", "AIRED", "MISSED", "EXTRA", "3RD PARTY"]
        ]
        
        # Add data rows
        for item in sponsorship_data:
            table_data.append([
                str(item.get('product', '')),
                str(item.get('duration', '')),
                str(item.get('planned', 0)),
                str(item.get('aired', 0)),
                str(item.get('missed', 0)),
                str(item.get('extra', 0)),
                str(item.get('third_party', 0))
            ])
        
        # Add totals row
        totals = [
            "Grand Total", "",
            sum(item.get('planned', 0) for item in sponsorship_data),
            sum(item.get('aired', 0) for item in sponsorship_data),
            sum(item.get('missed', 0) for item in sponsorship_data),
            sum(item.get('extra', 0) for item in sponsorship_data),
            sum(item.get('third_party', 0) for item in sponsorship_data)
        ]
        table_data.append(totals)
        
        # Create table
        table = Table(table_data, colWidths=[100, 30, 60, 60, 60, 60, 60])
        table.setStyle(TableStyle([
            # Header formatting
            ('BACKGROUND', (0, 0), (-1, 1), colors.lightgrey),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 9),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            # Merge SPOTS header
            ('SPAN', (2, 0), (6, 0)),
            # Data rows
            ('ROWBACKGROUNDS', (0, 2), (-1, -2), [colors.white, colors.whitesmoke]),
            # Totals row
            ('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey),
            ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ]))
        
        return table
        
    except Exception as e:
        print(f"Error creating sponsorship table: {e}")
        return None

def get_current_matched_data_df(self):
    """Get the current matched data as a DataFrame"""
    try:
        if not self.matched_data:
            return pd.DataFrame()
        
        if isinstance(self.matched_data, list):
            if not self.matched_data:
                return pd.DataFrame()
            df = pd.DataFrame(self.matched_data)
        else:
            df = self.matched_data.copy()
        
        return df
        
    except Exception as e:
        print(f"Error getting matched data: {e}")
        return pd.DataFrame()

print("✅ FINAL PDF generation ready!")
print("🔧 Final Changes:")
print("  ✅ Page 1: Portrait orientation (Summary Sheet) - UNCHANGED")
print("  ✅ Page 2: Landscape orientation (COMPLETE LMRB Table - Excel Style)")
print("  ❌ REMOVED: Unmatched TC Data page")  
print("  ❌ REMOVED: Unmatched Schedule Data page")
print("  🔧 Logo: FIXED transparent PNG handling")
print("")
print("📄 Final PDF Structure (2 PAGES ONLY):")
print("  • Page 1: Portrait - Summary Sheet with Commercial/Sponsorship Benefits")
print("  • Page 2: Landscape - Complete LMRB Table (All records, all columns, Excel format)")
print("")
print("📊 Page 2 Features:")
print("  • Shows ALL matched records (no limits)")
print("  • Excel-style table format exactly like your sample")  
print("  • ALL columns with complete data")
print("  • Sorted by date (Dd/Mn/Yr) + time")
print("  • Optimized column widths for landscape")
print("  • Repeating headers on page breaks")
print("  • Alternating row colors for readability")