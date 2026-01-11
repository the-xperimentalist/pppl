from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from django.http import HttpResponse
import openpyxl
from decimal import Decimal


class ExcelTemplateGenerator:
    """Generate Excel templates for bulk upload"""
    
    @staticmethod
    def create_raw_materials_template():
        """Create template for raw materials with vertical headers"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Raw Materials"

        # Headers in first column (vertical)
        headers = [
            'Material Name*',
            'Grade',
            'RM Code*',
            'Unit (kg/gm/ton)*',
            'RM Rate*',
            'Frozen Rate',
            'Part Weight*',
            'Runner Weight*',
            'Process Losses',
            'Purging Loss Cost',
            'ICC %',
            'Rejection %',
            'Overhead %',
            'Maintenance %',
            'Profit %'
        ]

        # Style for header column
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")

        # Add headers vertically in column A
        for row_num, header in enumerate(headers, 1):
            cell = ws.cell(row=row_num, column=1)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='left', vertical='center')

        # Set column A width
        ws.column_dimensions['A'].width = 25

        # Add sample data in columns B and C
        sample_data_1 = [
            'PP Compound', 'Grade A', 'RM001', 'kg', 150.50, '', 0.025, 0.005,
            0, 0, 0, 2, 5, 3, 10
        ]
        sample_data_2 = [
            'ABS Plastic', 'Grade B', 'RM002', 'kg', 200.00, 195.00, 0.030, 0.006,
            5, 2, 1, 2.5, 4.5, 2.5, 12
        ]

        # Add sample columns
        for row_num, value in enumerate(sample_data_1, 1):
            ws.cell(row=row_num, column=2, value=value)

        for row_num, value in enumerate(sample_data_2, 1):
            ws.cell(row=row_num, column=3, value=value)

        # Style sample columns
        for col in [2, 3]:
            ws.column_dimensions[get_column_letter(col)].width = 15

        return wb

    @staticmethod
    def create_moulding_machines_template():
        """Create template for moulding machines with vertical headers"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Moulding Machines"

        headers = [
            'Cavity*',
            'Machine Tonnage*',
            'Cycle Time (s)*',
            'Efficiency %*',
            'Shift Rate*',
            'Shift Rate for MTC*',
            'MTC Count*',
            'Rejection %',
            'Overhead %',
            'Maintenance %',
            'Profit %'
        ]

        header_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")

        for row_num, header in enumerate(headers, 1):
            cell = ws.cell(row=row_num, column=1)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='left', vertical='center')

        ws.column_dimensions['A'].width = 25

        sample_data_1 = [2, 150, 45, 85, 5000, 4500, 2, 3, 5, 4, 15]
        sample_data_2 = [4, 250, 60, 90, 7000, 6500, 3, 2.5, 4.5, 3.5, 12]

        for row_num, value in enumerate(sample_data_1, 1):
            ws.cell(row=row_num, column=2, value=value)

        for row_num, value in enumerate(sample_data_2, 1):
            ws.cell(row=row_num, column=3, value=value)

        for col in [2, 3]:
            ws.column_dimensions[get_column_letter(col)].width = 15

        return wb

    @staticmethod
    def create_assemblies_template():
        """Create template for assemblies with vertical headers"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Assemblies"

        headers = [
            'Assembly Name*',
            'Assembly Type',
            'Manual Cost',
            'Other Cost',
            'Profit %',
            'Rejection %',
            'Inspection & Handling %'
        ]

        header_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")

        for row_num, header in enumerate(headers, 1):
            cell = ws.cell(row=row_num, column=1)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='left', vertical='center')

        ws.column_dimensions['A'].width = 25

        sample_data_1 = ['Manual Assembly', 'Standard', 10.5, 5.0, 10, 2.5, 3.0]
        sample_data_2 = ['Automated Assembly', 'Premium', 15.0, 8.0, 12, 3.0, 4.5]

        for row_num, value in enumerate(sample_data_1, 1):
            ws.cell(row=row_num, column=2, value=value)

        for row_num, value in enumerate(sample_data_2, 1):
            ws.cell(row=row_num, column=3, value=value)

        for col in [2, 3]:
            ws.column_dimensions[get_column_letter(col)].width = 15

        return wb

    @staticmethod
    def create_packaging_template():
        """Create template for packaging with vertical headers"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Packaging"

        headers = [
            'Packaging Type*',
            'Packaging Length (mm)*',
            'Packaging Breadth (mm)*',
            'Packaging Height (mm)*',
            'Polybag Length*',
            'Polybag Width*',
            'Lifecycle*',
            'Cost*',
            'Maintenance %',
            'Parts per Polybag*'
        ]

        header_fill = PatternFill(start_color="E26B0A", end_color="E26B0A", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")

        for row_num, header in enumerate(headers, 1):
            cell = ws.cell(row=row_num, column=1)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='left', vertical='center')

        ws.column_dimensions['A'].width = 30

        sample_data_1 = ['pp_box', 600, 400, 250, 16, 20, 100, 5.50, 5, 50]
        sample_data_2 = ['cg_box', 600, 400, 250, 16, 20, 200, 12.00, 4, 100]

        for row_num, value in enumerate(sample_data_1, 1):
            ws.cell(row=row_num, column=2, value=value)

        for row_num, value in enumerate(sample_data_2, 1):
            ws.cell(row=row_num, column=3, value=value)

        for col in [2, 3]:
            ws.column_dimensions[get_column_letter(col)].width = 15

        return wb

    @staticmethod
    def create_transport_template():
        """Create template for transport with vertical headers"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Transport"

        headers = [
            'Transport Length*',
            'Transport Breadth*',
            'Transport Height*',
            'Trip Cost*',
            'Parts per Box*'
        ]

        header_fill = PatternFill(start_color="9933FF", end_color="9933FF", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")

        for row_num, header in enumerate(headers, 1):
            cell = ws.cell(row=row_num, column=1)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='left', vertical='center')

        ws.column_dimensions['A'].width = 25

        sample_data_1 = [3000, 2000, 1500, 5000, 100]
        sample_data_2 = [4000, 2500, 2000, 8000, 200]

        for row_num, value in enumerate(sample_data_1, 1):
            ws.cell(row=row_num, column=2, value=value)

        for row_num, value in enumerate(sample_data_2, 1):
            ws.cell(row=row_num, column=3, value=value)

        for col in [2, 3]:
            ws.column_dimensions[get_column_letter(col)].width = 15

        return wb

    @staticmethod
    def create_complete_quote_template():
        """Create complete template with all sheets in vertical format"""
        wb = Workbook()
        wb.remove(wb.active)  # Remove default sheet

        # Add Quote Definition sheet
        ws_def = wb.create_sheet("Quote Definition")
        def_headers = [
            'Quote Name*',
            'Client Name*',
            'SAP Number',
            'Part Number*',
            'Part Name*',
            'Amendment Number',
            'Description',
            'Quantity*',
            'Handling Charge',
            'Profit %',
            'Notes'
        ]
        ExcelTemplateGenerator._add_vertical_headers(ws_def, def_headers, "2E75B6")
        # Add sample quote columns
        ws_def.cell(1, 2, value="Quote 001")
        ws_def.cell(2, 2, value="ABC Corp")
        ws_def.cell(3, 2, value="SAP001")
        ws_def.cell(4, 2, value="PART-001")
        ws_def.cell(5, 2, value="Widget A")
        ws_def.cell(6, 2, value="A1")
        ws_def.cell(7, 2, value="Standard widget")
        ws_def.cell(8, 2, value=1000)
        ws_def.cell(9, 2, value=100)
        ws_def.cell(10, 2, value=15)
        ws_def.cell(11, 2, value="Initial quote")

        ws_def.cell(1, 3, value="Quote 002")
        ws_def.cell(2, 3, value="XYZ Inc")
        ws_def.cell(3, 3, value="SAP002")
        ws_def.cell(4, 3, value="PART-002")
        ws_def.cell(5, 3, value="Widget B")
        ws_def.cell(6, 3, value="B1")
        ws_def.cell(7, 3, value="Premium widget")
        ws_def.cell(8, 3, value=2000)
        ws_def.cell(9, 3, value=150)
        ws_def.cell(10, 3, value=20)
        ws_def.cell(11, 3, value="Premium quote")

        # Add Raw Materials sheet
        ws_rm = wb.create_sheet("Raw Materials")
        rm_headers = [
            'Material Name*', 'Grade', 'RM Code*', 'Unit (kg/gm/ton)*',
            'RM Rate*', 'Frozen Rate', 'Part Weight*', 'Runner Weight*',
            'Process Losses', 'Purging Loss Cost', 'ICC %',
            'Rejection %', 'Overhead %', 'Maintenance %', 'Profit %'
        ]
        ExcelTemplateGenerator._add_vertical_headers(ws_rm, rm_headers, "4472C4")

        # Add Moulding Machines sheet
        ws_mm = wb.create_sheet("Moulding Machines")
        mm_headers = [
            'Cavity*', 'Machine Tonnage*', 'Cycle Time (s)*', 'Efficiency %*',
            'Shift Rate*', 'Shift Rate for MTC*', 'MTC Count*',
            'Rejection %', 'Overhead %', 'Maintenance %', 'Profit %'
        ]
        ExcelTemplateGenerator._add_vertical_headers(ws_mm, mm_headers, "70AD47")

        # Add Assemblies sheet
        ws_asm = wb.create_sheet("Assemblies")
        asm_headers = [
            'Quote Name*', 'Assembly Name*', 'Assembly Type', 'Remarks',
        'Manual Cost', 'Other Cost', 'Profit %', 'Rejection %', 'Inspection & Handling %'
        ]
        ExcelTemplateGenerator._add_vertical_headers(ws_asm, asm_headers, "FFC000")

        # Add Packaging sheet
        ws_pkg = wb.create_sheet("Packaging")
        pkg_headers = [
            'Packaging Type*', 'Packaging Length (mm)*', 'Packaging Breadth (mm)*',
            'Packaging Height (mm)*', 'Polybag Length*', 'Polybag Width*',
            'Lifecycle*', 'Cost*', 'Maintenance %', 'Parts per Polybag*'
        ]
        ExcelTemplateGenerator._add_vertical_headers(ws_pkg, pkg_headers, "E26B0A")

        # Add Transport sheet
        ws_trans = wb.create_sheet("Transport")
        trans_headers = [
            'Transport Length*', 'Transport Breadth*', 'Transport Height*',
            'Trip Cost*', 'Parts per Box*'
        ]
        ExcelTemplateGenerator._add_vertical_headers(ws_trans, trans_headers, "9933FF")

        return wb

    @staticmethod
    def create_multiple_quotes_template():
        """Create comprehensive template for creating multiple complete quotes in a project"""
        wb = Workbook()
        wb.remove(wb.active)

        # Sheet 1: Quote Definition
        ws_def = wb.create_sheet("Quote Definition")
        def_headers = [
            'Quote Name*',
            'Client Name*',
            'SAP Number',
            'Part Number*',
            'Part Name*',
            'Amendment Number',
            'Description',
            'Quantity*',
            'Handling Charge',
            'Profit %',
            'Notes'
        ]
        ExcelTemplateGenerator._add_vertical_headers(ws_def, def_headers, "2E75B6")

        # Add sample quote columns
        ws_def.cell(1, 2, value="Quote 001")
        ws_def.cell(2, 2, value="ABC Corp")
        ws_def.cell(3, 2, value="SAP001")
        ws_def.cell(4, 2, value="PART-001")
        ws_def.cell(5, 2, value="Widget A")
        ws_def.cell(6, 2, value="A1")
        ws_def.cell(7, 2, value="Standard widget")
        ws_def.cell(8, 2, value=1000)
        ws_def.cell(9, 2, value=100)
        ws_def.cell(10, 2, value=15)
        ws_def.cell(11, 2, value="Initial quote")

        ws_def.cell(1, 3, value="Quote 002")
        ws_def.cell(2, 3, value="XYZ Inc")
        ws_def.cell(3, 3, value="SAP002")
        ws_def.cell(4, 3, value="PART-002")
        ws_def.cell(5, 3, value="Widget B")
        ws_def.cell(6, 3, value="B1")
        ws_def.cell(7, 3, value="Premium widget")
        ws_def.cell(8, 3, value=2000)
        ws_def.cell(9, 3, value=150)
        ws_def.cell(10, 3, value=20)
        ws_def.cell(11, 3, value="Premium quote")

        for col in [2, 3]:
            ws_def.column_dimensions[get_column_letter(col)].width = 15

        # Sheet 2: Raw Materials (with Quote Name in first row)
        ws_rm = wb.create_sheet("Raw Materials")
        rm_headers = [
            'Quote Name*',
            'Material Name*',
            'Grade',
            'RM Code*',
            'Unit (kg/gm/ton)*',
            'RM Rate*',
            'Frozen Rate',
            'Part Weight*',
            'Runner Weight*',
            'Process Losses',
            'Purging Loss Cost',
            'ICC %',
            'Rejection %',
            'Overhead %',
            'Maintenance %',
            'Profit %'
        ]
        ExcelTemplateGenerator._add_vertical_headers(ws_rm, rm_headers, "4472C4")

        # Add sample data for Quote 001
        ws_rm.cell(1, 2, value="Quote 001")
        ws_rm.cell(2, 2, value="PP Compound")
        ws_rm.cell(3, 2, value="Grade A")
        ws_rm.cell(4, 2, value="RM001")
        ws_rm.cell(5, 2, value="kg")
        ws_rm.cell(6, 2, value=150.50)
        ws_rm.cell(7, 2, value="")
        ws_rm.cell(8, 2, value=0.025)
        ws_rm.cell(9, 2, value=0.005)
        ws_rm.cell(10, 2, value=0)
        ws_rm.cell(11, 2, value=0)
        ws_rm.cell(12, 2, value=0)
        ws_rm.cell(13, 2, value=2)
        ws_rm.cell(14, 2, value=5)
        ws_rm.cell(15, 2, value=3)
        ws_rm.cell(16, 2, value=10)

        # Add sample data for Quote 002
        ws_rm.cell(1, 3, value="Quote 002")
        ws_rm.cell(2, 3, value="ABS Plastic")
        ws_rm.cell(3, 3, value="Grade B")
        ws_rm.cell(4, 3, value="RM002")
        ws_rm.cell(5, 3, value="kg")
        ws_rm.cell(6, 3, value=200.00)
        ws_rm.cell(7, 3, value=195.00)
        ws_rm.cell(8, 3, value=0.030)
        ws_rm.cell(9, 3, value=0.006)
        ws_rm.cell(10, 3, value=5)
        ws_rm.cell(11, 3, value=2)
        ws_rm.cell(12, 3, value=1)
        ws_rm.cell(13, 3, value=2.5)
        ws_rm.cell(14, 3, value=4.5)
        ws_rm.cell(15, 3, value=2.5)
        ws_rm.cell(16, 3, value=12)

        for col in [2, 3]:
            ws_rm.column_dimensions[get_column_letter(col)].width = 15

        # Sheet 3: Moulding Machines
        ws_mm = wb.create_sheet("Moulding Machines")
        mm_headers = [
            'Quote Name*',
            'Cavity*',
            'Machine Tonnage*',
            'Cycle Time (s)*',
            'Efficiency %*',
            'Shift Rate*',
            'Shift Rate for MTC*',
            'MTC Count*',
            'Rejection %',
            'Overhead %',
            'Maintenance %',
            'Profit %'
        ]
        ExcelTemplateGenerator._add_vertical_headers(ws_mm, mm_headers, "70AD47")

        # Sample for Quote 001
        ws_mm.cell(1, 2, value="Quote 001")
        ws_mm.cell(2, 2, value=2)
        ws_mm.cell(3, 2, value=150)
        ws_mm.cell(4, 2, value=45)
        ws_mm.cell(5, 2, value=85)
        ws_mm.cell(6, 2, value=5000)
        ws_mm.cell(7, 2, value=4500)
        ws_mm.cell(8, 2, value=2)
        ws_mm.cell(9, 2, value=3)
        ws_mm.cell(10, 2, value=5)
        ws_mm.cell(11, 2, value=4)
        ws_mm.cell(12, 2, value=15)

        # Sample for Quote 002
        ws_mm.cell(1, 3, value="Quote 002")
        ws_mm.cell(2, 3, value=4)
        ws_mm.cell(3, 3, value=250)
        ws_mm.cell(4, 3, value=60)
        ws_mm.cell(5, 3, value=90)
        ws_mm.cell(6, 3, value=7000)
        ws_mm.cell(7, 3, value=6500)
        ws_mm.cell(8, 3, value=3)
        ws_mm.cell(9, 3, value=2.5)
        ws_mm.cell(10, 3, value=4.5)
        ws_mm.cell(11, 3, value=3.5)
        ws_mm.cell(12, 3, value=12)

        for col in [2, 3]:
            ws_mm.column_dimensions[get_column_letter(col)].width = 15

        # Sheet 4: Assemblies
        ws_asm = wb.create_sheet("Assemblies")
        asm_headers = [
            'Quote Name*',
            'Assembly Name*',
            'Assembly Type',
            'Manual Cost',
            'Other Cost',
            'Profit %',
            'Rejection %',
            'Inspection & Handling %'
        ]
        ExcelTemplateGenerator._add_vertical_headers(ws_asm, asm_headers, "FFC000")

        # Sample for Quote 001
        ws_asm.cell(1, 2, value="Quote 001")
        ws_asm.cell(2, 2, value="Manual Assembly")
        ws_asm.cell(3, 2, value="Standard")
        ws_asm.cell(4, 2, value=10.5)
        ws_asm.cell(5, 2, value=5.0)
        ws_asm.cell(6, 2, value=10)
        ws_asm.cell(7, 2, value=2.5)
        ws_asm.cell(8, 2, value=3.0)

        # Sample for Quote 002
        ws_asm.cell(1, 3, value="Quote 002")
        ws_asm.cell(2, 3, value="Automated Assembly")
        ws_asm.cell(3, 3, value="Premium")
        ws_asm.cell(4, 3, value=15.0)
        ws_asm.cell(5, 3, value=8.0)
        ws_asm.cell(6, 3, value=12)
        ws_asm.cell(7, 3, value=3.0)
        ws_asm.cell(8, 3, value=4.5)

        for col in [2, 3]:
            ws_asm.column_dimensions[get_column_letter(col)].width = 15

        # Sheet 5: Packaging
        ws_pkg = wb.create_sheet("Packaging")
        pkg_headers = [
            'Quote Name*',
            'Packaging Type*',
            'Packaging Length (mm)*',
            'Packaging Breadth (mm)*',
            'Packaging Height (mm)*',
            'Polybag Length*',
            'Polybag Width*',
            'Lifecycle*',
            'Cost*',
            'Maintenance %',
            'Parts per Polybag*'
        ]
        ExcelTemplateGenerator._add_vertical_headers(ws_pkg, pkg_headers, "E26B0A")

        # Sample for Quote 001
        ws_pkg.cell(1, 2, value="Quote 001")
        ws_pkg.cell(2, 2, value="pp_box")
        ws_pkg.cell(3, 2, value=600)
        ws_pkg.cell(4, 2, value=400)
        ws_pkg.cell(5, 2, value=250)
        ws_pkg.cell(6, 2, value=16)
        ws_pkg.cell(7, 2, value=20)
        ws_pkg.cell(8, 2, value=100)
        ws_pkg.cell(9, 2, value=5.50)
        ws_pkg.cell(10, 2, value=5)
        ws_pkg.cell(11, 2, value=50)

        # Sample for Quote 002
        ws_pkg.cell(1, 3, value="Quote 002")
        ws_pkg.cell(2, 3, value="cg_box")
        ws_pkg.cell(3, 3, value=600)
        ws_pkg.cell(4, 3, value=400)
        ws_pkg.cell(5, 3, value=250)
        ws_pkg.cell(6, 3, value=16)
        ws_pkg.cell(7, 3, value=20)
        ws_pkg.cell(8, 3, value=200)
        ws_pkg.cell(9, 3, value=12.00)
        ws_pkg.cell(10, 3, value=4)
        ws_pkg.cell(11, 3, value=100)

        for col in [2, 3]:
            ws_pkg.column_dimensions[get_column_letter(col)].width = 15

        # Sheet 6: Transport
        ws_trans = wb.create_sheet("Transport")
        trans_headers = [
            'Quote Name*',
            'Transport Length*',
            'Transport Breadth*',
            'Transport Height*',
            'Trip Cost*',
            'Parts per Box*'
        ]
        ExcelTemplateGenerator._add_vertical_headers(ws_trans, trans_headers, "9933FF")

        # Sample for Quote 001
        ws_trans.cell(1, 2, value="Quote 001")
        ws_trans.cell(2, 2, value=3000)
        ws_trans.cell(3, 2, value=2000)
        ws_trans.cell(4, 2, value=1500)
        ws_trans.cell(5, 2, value=5000)
        ws_trans.cell(6, 2, value=100)

        # Sample for Quote 002
        ws_trans.cell(1, 3, value="Quote 002")
        ws_trans.cell(2, 3, value=4000)
        ws_trans.cell(3, 3, value=2500)
        ws_trans.cell(4, 3, value=2000)
        ws_trans.cell(5, 3, value=8000)
        ws_trans.cell(6, 3, value=200)

        for col in [2, 3]:
            ws_trans.column_dimensions[get_column_letter(col)].width = 15

        # Add instruction sheet
        ws_inst = wb.create_sheet("Instructions", 0)
        ws_inst.column_dimensions['A'].width = 80

        instructions = [
            "INSTRUCTIONS FOR BULK QUOTE CREATION",
            "",
            "This template allows you to create multiple complete quotes in a project.",
            "",
            "STRUCTURE:",
            "- Each sheet represents a different component of the quote",
            "- Column A contains the field names (headers)",
            "- Each subsequent column (B, C, D, etc.) represents one entry",
            "- The first row of each sheet (except Quote Definition) must contain the Quote Name to link components",
            "",
            "SHEETS:",
            "1. Quote Definition: Basic quote information (one column = one quote)",
            "2. Raw Materials: Raw materials for quotes (can have multiple entries per quote)",
            "3. Moulding Machines: Moulding machines for quotes (can have multiple entries per quote)",
            "4. Assemblies: Assemblies for quotes (can have multiple entries per quote)",
            "5. Packaging: Packaging for quotes (can have multiple entries per quote)",
            "6. Transport: Transport for quotes (can have multiple entries per quote)",
            "",
            "HOW TO USE:",
            "1. Fill in Quote Definition sheet with your quotes (each column is one quote)",
            "2. For each quote component sheet, use the EXACT quote name from Quote Definition in row 1",
            "3. You can add multiple columns for the same quote to create multiple components",
            "4. Example: If you have 'Quote 001' in Quote Definition, you can have:",
            "   - 3 columns in Raw Materials with 'Quote 001' to add 3 raw materials",
            "   - 2 columns in Moulding Machines with 'Quote 001' to add 2 machines",
            "   - And so on...",
            "",
            "IMPORTANT:",
            "- Fields marked with * are required",
            "- Quote Name must match exactly between sheets (case-sensitive)",
            "- All quotes will be created in the same project",
            "- All quotes will use the customer group you select during upload",
            "",
            "EXAMPLE:",
            "The template includes sample data for 'Quote 001' and 'Quote 002'",
            "You can see how each quote has components across different sheets",
        ]

        for row_num, instruction in enumerate(instructions, 1):
            cell = ws_inst.cell(row=row_num, column=1)
            cell.value = instruction
            if row_num == 1:
                cell.font = Font(bold=True, size=14)
            elif instruction.endswith(":") or instruction.startswith("STRUCTURE") or instruction.startswith("SHEETS") or instruction.startswith("HOW TO USE") or instruction.startswith("IMPORTANT") or instruction.startswith("EXAMPLE"):
                cell.font = Font(bold=True)

        return wb

    @staticmethod
    def _add_vertical_headers(worksheet, headers, color):
        """Helper to add vertical headers to worksheet"""
        header_fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")

        for row_num, header in enumerate(headers, 1):
            cell = worksheet.cell(row=row_num, column=1)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='left', vertical='center')

        worksheet.column_dimensions['A'].width = 30


class ExcelParser:
    """Parse uploaded Excel files with vertical format"""

    @staticmethod
    def parse_raw_materials(file, quote):
        """Parse raw materials from Excel (vertical format)"""
        from core.models import RawMaterial

        wb = openpyxl.load_workbook(file)
        ws = wb.active

        raw_materials = []
        errors = []

        # Process each column starting from column 2 (B)
        max_col = ws.max_column

        for col_num in range(2, max_col + 1):
            # Check if column has data
            if not ws.cell(1, col_num).value:
                continue

            try:
                rm = RawMaterial(
                    quote=quote,
                    material_name=ws.cell(1, col_num).value or '',
                    grade=ws.cell(2, col_num).value or '',
                    rm_code=ws.cell(3, col_num).value or '',
                    unit_of_measurement=ws.cell(4, col_num).value or 'kg',
                    rm_rate=float(ws.cell(5, col_num).value) if ws.cell(5, col_num).value else 0,
                    frozen_rate=float(ws.cell(6, col_num).value) if ws.cell(6, col_num).value else None,
                    part_weight=float(ws.cell(7, col_num).value) if ws.cell(7, col_num).value else 0,
                    runner_weight=float(ws.cell(8, col_num).value) if ws.cell(8, col_num).value else 0,
                    process_losses=float(ws.cell(9, col_num).value) if ws.cell(9, col_num).value else 0,
                    purging_loss_cost=float(ws.cell(10, col_num).value) if ws.cell(10, col_num).value else 0,
                    icc_percentage=float(ws.cell(11, col_num).value) if ws.cell(11, col_num).value else 0,
                    rejection_percentage=float(ws.cell(12, col_num).value) if ws.cell(12, col_num).value else 0,
                    overhead_percentage=float(ws.cell(13, col_num).value) if ws.cell(13, col_num).value else 0,
                    maintenance_percentage=float(ws.cell(14, col_num).value) if ws.cell(14, col_num).value else 0,
                    profit_percentage=float(ws.cell(15, col_num).value) if ws.cell(15, col_num).value else 0,
                )
                raw_materials.append(rm)
            except Exception as e:
                errors.append(f"Column {get_column_letter(col_num)}: {str(e)}")

        if not errors:
            RawMaterial.objects.bulk_create(raw_materials)

        return len(raw_materials), errors

    @staticmethod
    def parse_moulding_machines(file, quote):
        """Parse moulding machines from Excel (vertical format)"""
        from core.models import MouldingMachineDetail

        wb = openpyxl.load_workbook(file)
        ws = wb.active

        machines = []
        errors = []

        max_col = ws.max_column

        for col_num in range(2, max_col + 1):
            if not ws.cell(1, col_num).value:
                continue

            try:
                mm = MouldingMachineDetail(
                    quote=quote,
                    cavity=int(ws.cell(1, col_num).value) if ws.cell(1, col_num).value else 1,
                    machine_tonnage=float(ws.cell(2, col_num).value) if ws.cell(2, col_num).value else 0,
                    cycle_time=float(ws.cell(3, col_num).value) if ws.cell(3, col_num).value else 0,
                    efficiency=float(ws.cell(4, col_num).value) if ws.cell(4, col_num).value else 0,
                    shift_rate=float(ws.cell(5, col_num).value) if ws.cell(5, col_num).value else 0,
                    shift_rate_for_mtc=float(ws.cell(6, col_num).value) if ws.cell(6, col_num).value else 0,
                    mtc_count=int(ws.cell(7, col_num).value) if ws.cell(7, col_num).value else 0,
                    rejection_percentage=float(ws.cell(8, col_num).value) if ws.cell(8, col_num).value else 0,
                    overhead_percentage=float(ws.cell(9, col_num).value) if ws.cell(9, col_num).value else 0,
                    maintenance_percentage=float(ws.cell(10, col_num).value) if ws.cell(10, col_num).value else 0,
                    profit_percentage=float(ws.cell(11, col_num).value) if ws.cell(11, col_num).value else 0,
                )
                machines.append(mm)
            except Exception as e:
                errors.append(f"Column {get_column_letter(col_num)}: {str(e)}")

        if not errors:
            MouldingMachineDetail.objects.bulk_create(machines)

        return len(machines), errors

    @staticmethod
    def parse_multiple_quotes(file, project, customer_group, user):
        """Parse multiple quotes from Excel file"""
        from core.models import Quote

        wb = openpyxl.load_workbook(file)

        if "Quotes" not in wb.sheetnames:
            return 0, ["'Quotes' sheet not found in the Excel file"]

        ws = wb["Quotes"]
        quotes_created = []
        errors = []

        max_col = ws.max_column

        for col_num in range(2, max_col + 1):
            # Check if column has quote name (first row)
            if not ws.cell(1, col_num).value:
                continue

            try:
                quote = Quote.objects.create(
                    project=project,
                    name=ws.cell(1, col_num).value or '',
                    client_group=customer_group,
                    client_name=ws.cell(2, col_num).value or '',
                    sap_number=ws.cell(3, col_num).value or '',
                    part_number=ws.cell(4, col_num).value or '',
                    part_name=ws.cell(5, col_num).value or '',
                    amendment_number=ws.cell(6, col_num).value or '',
                    description=ws.cell(7, col_num).value or '',
                    quantity=int(ws.cell(8, col_num).value) if ws.cell(8, col_num).value else 1,
                    handling_charge=float(ws.cell(9, col_num).value) if ws.cell(9, col_num).value else 0,
                    profit_percentage=float(ws.cell(10, col_num).value) if ws.cell(10, col_num).value else 0,
                    notes=ws.cell(11, col_num).value or '',
                    created_by=user,
                    quote_definition_complete=True
                )
                quotes_created.append(quote)

                # Add timeline entry
                from core.models import QuoteTimeline
                QuoteTimeline.add_entry(
                    quote=quote,
                    activity_type='quote_created',
                    description=f'Quote "{quote.name}" created via bulk upload',
                    user=user
                )

            except Exception as e:
                errors.append(f"Column {get_column_letter(col_num)}: {str(e)}")

        return len(quotes_created), errors

    @staticmethod
    def parse_complete_quote(file, quote):
        """Parse complete quote from multi-sheet Excel (vertical format)"""
        wb = openpyxl.load_workbook(file)
        results = {
            'raw_materials': {'count': 0, 'errors': []},
            'moulding_machines': {'count': 0, 'errors': []},
            'assemblies': {'count': 0, 'errors': []},
            'packaging': {'count': 0, 'errors': []},
            'transport': {'count': 0, 'errors': []},
        }

        # Parse Raw Materials
        if "Raw Materials" in wb.sheetnames:
            ws = wb["Raw Materials"]
            count, errors = ExcelParser._parse_raw_materials_sheet(ws, quote)
            results['raw_materials'] = {'count': count, 'errors': errors}

        # Parse Moulding Machines
        if "Moulding Machines" in wb.sheetnames:
            ws = wb["Moulding Machines"]
            count, errors = ExcelParser._parse_moulding_machines_sheet(ws, quote)
            results['moulding_machines'] = {'count': count, 'errors': errors}

        # Parse Assemblies
        if "Assemblies" in wb.sheetnames:
            ws = wb["Assemblies"]
            count, errors = ExcelParser._parse_assemblies_sheet(ws, quote)
            results['assemblies'] = {'count': count, 'errors': errors}

        # Parse Packaging
        if "Packaging" in wb.sheetnames:
            ws = wb["Packaging"]
            count, errors = ExcelParser._parse_packaging_sheet(ws, quote)
            results['packaging'] = {'count': count, 'errors': errors}

        # Parse Transport
        if "Transport" in wb.sheetnames:
            ws = wb["Transport"]
            count, errors = ExcelParser._parse_transport_sheet(ws, quote)
            results['transport'] = {'count': count, 'errors': errors}

        return results

    @staticmethod
    def _parse_raw_materials_sheet(worksheet, quote):
        """Helper to parse raw materials sheet (vertical format)"""
        from core.models import RawMaterial

        raw_materials = []
        errors = []

        max_col = worksheet.max_column

        for col_num in range(2, max_col + 1):
            if not worksheet.cell(1, col_num).value:
                continue

            try:
                rm = RawMaterial(
                    quote=quote,
                    material_name=worksheet.cell(1, col_num).value or '',
                    grade=worksheet.cell(2, col_num).value or '',
                    rm_code=worksheet.cell(3, col_num).value or '',
                    unit_of_measurement=worksheet.cell(4, col_num).value or 'kg',
                    rm_rate=float(worksheet.cell(5, col_num).value) if worksheet.cell(5, col_num).value else 0,
                    frozen_rate=float(worksheet.cell(6, col_num).value) if worksheet.cell(6, col_num).value else None,
                    part_weight=float(worksheet.cell(7, col_num).value) if worksheet.cell(7, col_num).value else 0,
                    runner_weight=float(worksheet.cell(8, col_num).value) if worksheet.cell(8, col_num).value else 0,
                    process_losses=float(worksheet.cell(9, col_num).value) if worksheet.cell(9, col_num).value else 0,
                    purging_loss_cost=float(worksheet.cell(10, col_num).value) if worksheet.cell(10, col_num).value else 0,
                    icc_percentage=float(worksheet.cell(11, col_num).value) if worksheet.cell(11, col_num).value else 0,
                    rejection_percentage=float(worksheet.cell(12, col_num).value) if worksheet.cell(12, col_num).value else 0,
                    overhead_percentage=float(worksheet.cell(13, col_num).value) if worksheet.cell(13, col_num).value else 0,
                    maintenance_percentage=float(worksheet.cell(14, col_num).value) if worksheet.cell(14, col_num).value else 0,
                    profit_percentage=float(worksheet.cell(15, col_num).value) if worksheet.cell(15, col_num).value else 0,
                )
                raw_materials.append(rm)
            except Exception as e:
                errors.append(f"Column {get_column_letter(col_num)}: {str(e)}")

        if not errors:
            RawMaterial.objects.bulk_create(raw_materials)

        return len(raw_materials), errors

    @staticmethod
    def _parse_moulding_machines_sheet(worksheet, quote):
        """Helper to parse moulding machines sheet (vertical format)"""
        from core.models import MouldingMachineDetail

        machines = []
        errors = []

        max_col = worksheet.max_column

        for col_num in range(2, max_col + 1):
            if not worksheet.cell(1, col_num).value:
                continue

            try:
                mm = MouldingMachineDetail(
                    quote=quote,
                    cavity=int(worksheet.cell(1, col_num).value) if worksheet.cell(1, col_num).value else 1,
                    machine_tonnage=float(worksheet.cell(2, col_num).value) if worksheet.cell(2, col_num).value else 0,
                    cycle_time=float(worksheet.cell(3, col_num).value) if worksheet.cell(3, col_num).value else 0,
                    efficiency=float(worksheet.cell(4, col_num).value) if worksheet.cell(4, col_num).value else 0,
                    shift_rate=float(worksheet.cell(5, col_num).value) if worksheet.cell(5, col_num).value else 0,
                    shift_rate_for_mtc=float(worksheet.cell(6, col_num).value) if worksheet.cell(6, col_num).value else 0,
                    mtc_count=int(worksheet.cell(7, col_num).value) if worksheet.cell(7, col_num).value else 0,
                    rejection_percentage=float(worksheet.cell(8, col_num).value) if worksheet.cell(8, col_num).value else 0,
                    overhead_percentage=float(worksheet.cell(9, col_num).value) if worksheet.cell(9, col_num).value else 0,
                    maintenance_percentage=float(worksheet.cell(10, col_num).value) if worksheet.cell(10, col_num).value else 0,
                    profit_percentage=float(worksheet.cell(11, col_num).value) if worksheet.cell(11, col_num).value else 0,
                )
                machines.append(mm)
            except Exception as e:
                errors.append(f"Column {get_column_letter(col_num)}: {str(e)}")

        if not errors:
            MouldingMachineDetail.objects.bulk_create(machines)

        return len(machines), errors

    @staticmethod
    def _parse_assemblies_sheet(worksheet, quote):
        """Helper to parse assemblies sheet (vertical format)"""
        from core.models import Assembly

        assemblies = []
        errors = []

        max_col = worksheet.max_column

        for col_num in range(2, max_col + 1):
            if not worksheet.cell(1, col_num).value:
                continue

            try:
                assembly = Assembly(
                    quote=quote,
                    name=worksheet.cell(2, col_num).value or '',
                    remarks=worksheet.cell(4, col_num).value or '',  # Added remarks
                    manual_cost=float(worksheet.cell(5, col_num).value) if worksheet.cell(5, col_num).value else 0,
                    other_cost=float(worksheet.cell(6, col_num).value) if worksheet.cell(6, col_num).value else 0,
                    profit_percentage=float(worksheet.cell(7, col_num).value) if worksheet.cell(7, col_num).value else 0,
                    rejection_percentage=float(worksheet.cell(8, col_num).value) if worksheet.cell(8, col_num).value else 0,
                    inspection_handling_percentage=float(worksheet.cell(9, col_num).value) if worksheet.cell(9, col_num).value else 0,
                )
                assemblies.append(assembly)
            except Exception as e:
                errors.append(f"Column {get_column_letter(col_num)}: {str(e)}")

        if not errors:
            Assembly.objects.bulk_create(assemblies)

        return len(assemblies), errors

    @staticmethod
    def _parse_packaging_sheet(worksheet, quote):
        """Helper to parse packaging sheet (vertical format)"""
        from core.models import Packaging

        packagings = []
        errors = []

        max_col = worksheet.max_column

        for col_num in range(2, max_col + 1):
            if not worksheet.cell(1, col_num).value:
                continue

            try:
                packaging = Packaging(
                    quote=quote,
                    packaging_type=worksheet.cell(1, col_num).value or 'pp_box',
                    packaging_length=float(worksheet.cell(2, col_num).value) if worksheet.cell(2, col_num).value else 600,
                    packaging_breadth=float(worksheet.cell(3, col_num).value) if worksheet.cell(3, col_num).value else 400,
                    packaging_height=float(worksheet.cell(4, col_num).value) if worksheet.cell(4, col_num).value else 250,
                    polybag_length=float(worksheet.cell(5, col_num).value) if worksheet.cell(5, col_num).value else 16,
                    polybag_width=float(worksheet.cell(6, col_num).value) if worksheet.cell(6, col_num).value else 20,
                    lifecycle=int(worksheet.cell(7, col_num).value) if worksheet.cell(7, col_num).value else 1,
                    cost=float(worksheet.cell(8, col_num).value) if worksheet.cell(8, col_num).value else 0,
                    maintenance_percentage=float(worksheet.cell(9, col_num).value) if worksheet.cell(9, col_num).value else 0,
                    part_per_polybag=int(worksheet.cell(10, col_num).value) if worksheet.cell(10, col_num).value else 1,
                )
                packagings.append(packaging)
            except Exception as e:
                errors.append(f"Column {get_column_letter(col_num)}: {str(e)}")

        if not errors:
            Packaging.objects.bulk_create(packagings)

        return len(packagings), errors

    @staticmethod
    def _parse_transport_sheet(worksheet, quote):
        """Helper to parse transport sheet (vertical format)"""
        from core.models import Transport

        transports = []
        errors = []

        max_col = worksheet.max_column

        for col_num in range(2, max_col + 1):
            if not worksheet.cell(1, col_num).value:
                continue

            try:
                transport = Transport(
                    quote=quote,
                    transport_length=float(worksheet.cell(1, col_num).value) if worksheet.cell(1, col_num).value else 0,
                    transport_breadth=float(worksheet.cell(2, col_num).value) if worksheet.cell(2, col_num).value else 0,
                    transport_height=float(worksheet.cell(3, col_num).value) if worksheet.cell(3, col_num).value else 0,
                    trip_cost=float(worksheet.cell(4, col_num).value) if worksheet.cell(4, col_num).value else 0,
                    parts_per_box=int(worksheet.cell(5, col_num).value) if worksheet.cell(5, col_num).value else 1,
                )
                transports.append(transport)
            except Exception as e:
                errors.append(f"Column {get_column_letter(col_num)}: {str(e)}")

        if not errors:
            Transport.objects.bulk_create(transports)

        return len(transports), errors

    @staticmethod
    def parse_multiple_quotes_complete(file, project, customer_group, user):
        """Parse multiple complete quotes with all components from Excel file"""
        from core.models import Quote, QuoteTimeline, RawMaterial, MouldingMachineDetail, Assembly, Packaging, Transport

        wb = openpyxl.load_workbook(file)

        # First, parse all quotes from Quote Definition sheet
        if "Quote Definition" not in wb.sheetnames:
            return {'quotes': 0, 'components': {}, 'errors': ["'Quote Definition' sheet not found"]}

        ws_def = wb["Quote Definition"]
        quotes_dict = {}  # Map quote names to Quote objects
        errors = []

        max_col = ws_def.max_column

        # Create all quotes first
        for col_num in range(2, max_col + 1):
            quote_name = ws_def.cell(1, col_num).value
            if not quote_name:
                continue

            try:
                quote = Quote.objects.create(
                    project=project,
                    name=quote_name,
                    client_group=customer_group,
                    client_name=ws_def.cell(2, col_num).value or '',
                    sap_number=ws_def.cell(3, col_num).value or '',
                    part_number=ws_def.cell(4, col_num).value or '',
                    part_name=ws_def.cell(5, col_num).value or '',
                    amendment_number=ws_def.cell(6, col_num).value or '',
                    description=ws_def.cell(7, col_num).value or '',
                    quantity=int(ws_def.cell(8, col_num).value) if ws_def.cell(8, col_num).value else 1,
                    handling_charge=float(ws_def.cell(9, col_num).value) if ws_def.cell(9, col_num).value else 0,
                    profit_percentage=float(ws_def.cell(10, col_num).value) if ws_def.cell(10, col_num).value else 0,
                    notes=ws_def.cell(11, col_num).value or '',
                    created_by=user,
                    quote_definition_complete=True
                )
                quotes_dict[quote_name] = quote

                # Add timeline entry
                QuoteTimeline.add_entry(
                    quote=quote,
                    activity_type='quote_created',
                    description=f'Quote "{quote.name}" created via bulk upload',
                    user=user
                )
            except Exception as e:
                errors.append(f"Quote Definition Column {get_column_letter(col_num)}: {str(e)}")

        # Component counts
        component_counts = {
            'raw_materials': 0,
            'moulding_machines': 0,
            'assemblies': 0,
            'packaging': 0,
            'transport': 0
        }

        # Parse Raw Materials
        if "Raw Materials" in wb.sheetnames:
            ws = wb["Raw Materials"]
            count, sheet_errors = ExcelParser._parse_components_by_quote(
                ws, quotes_dict, 'raw_material', user
            )
            component_counts['raw_materials'] = count
            errors.extend(sheet_errors)

        # Parse Moulding Machines
        if "Moulding Machines" in wb.sheetnames:
            ws = wb["Moulding Machines"]
            count, sheet_errors = ExcelParser._parse_components_by_quote(
                ws, quotes_dict, 'moulding_machine', user
            )
            component_counts['moulding_machines'] = count
            errors.extend(sheet_errors)

        # Parse Assemblies
        if "Assemblies" in wb.sheetnames:
            ws = wb["Assemblies"]
            count, sheet_errors = ExcelParser._parse_components_by_quote(
                ws, quotes_dict, 'assembly', user
            )
            component_counts['assemblies'] = count
            errors.extend(sheet_errors)

        # Parse Packaging
        if "Packaging" in wb.sheetnames:
            ws = wb["Packaging"]
            count, sheet_errors = ExcelParser._parse_components_by_quote(
                ws, quotes_dict, 'packaging', user
            )
            component_counts['packaging'] = count
            errors.extend(sheet_errors)

        # Parse Transport
        if "Transport" in wb.sheetnames:
            ws = wb["Transport"]
            count, sheet_errors = ExcelParser._parse_components_by_quote(
                ws, quotes_dict, 'transport', user
            )
            component_counts['transport'] = count
            errors.extend(sheet_errors)

        return {
            'quotes': len(quotes_dict),
            'components': component_counts,
            'errors': errors
        }

    @staticmethod
    def _parse_components_by_quote(worksheet, quotes_dict, component_type, user):
        """Helper to parse components grouped by quote name"""
        from core.models import RawMaterial, MouldingMachineDetail, Assembly, Packaging, Transport

        count = 0
        errors = []
        max_col = worksheet.max_column

        for col_num in range(2, max_col + 1):
            quote_name = worksheet.cell(1, col_num).value

            if not quote_name:
                continue

            if quote_name not in quotes_dict:
                errors.append(f"{component_type.title()} Column {get_column_letter(col_num)}: Quote '{quote_name}' not found in Quote Definition")
                continue

            quote = quotes_dict[quote_name]

            try:
                if component_type == 'raw_material':
                    RawMaterial.objects.create(
                        quote=quote,
                        material_name=worksheet.cell(2, col_num).value or '',
                        grade=worksheet.cell(3, col_num).value or '',
                        rm_code=worksheet.cell(4, col_num).value or '',
                        unit_of_measurement=worksheet.cell(5, col_num).value or 'kg',
                        rm_rate=float(worksheet.cell(6, col_num).value) if worksheet.cell(6, col_num).value else 0,
                        frozen_rate=float(worksheet.cell(7, col_num).value) if worksheet.cell(7, col_num).value else None,
                        part_weight=float(worksheet.cell(8, col_num).value) if worksheet.cell(8, col_num).value else 0,
                        runner_weight=float(worksheet.cell(9, col_num).value) if worksheet.cell(9, col_num).value else 0,
                        process_losses=float(worksheet.cell(10, col_num).value) if worksheet.cell(10, col_num).value else 0,
                        purging_loss_cost=float(worksheet.cell(11, col_num).value) if worksheet.cell(11, col_num).value else 0,
                        icc_percentage=float(worksheet.cell(12, col_num).value) if worksheet.cell(12, col_num).value else 0,
                        rejection_percentage=float(worksheet.cell(13, col_num).value) if worksheet.cell(13, col_num).value else 0,
                        overhead_percentage=float(worksheet.cell(14, col_num).value) if worksheet.cell(14, col_num).value else 0,
                        maintenance_percentage=float(worksheet.cell(15, col_num).value) if worksheet.cell(15, col_num).value else 0,
                        profit_percentage=float(worksheet.cell(16, col_num).value) if worksheet.cell(16, col_num).value else 0,
                    )
                    quote.increment_version(user, f'Raw material added via bulk upload')

                elif component_type == 'moulding_machine':
                    MouldingMachineDetail.objects.create(
                        quote=quote,
                        cavity=int(worksheet.cell(2, col_num).value) if worksheet.cell(2, col_num).value else 1,
                        machine_tonnage=float(worksheet.cell(3, col_num).value) if worksheet.cell(3, col_num).value else 0,
                        cycle_time=float(worksheet.cell(4, col_num).value) if worksheet.cell(4, col_num).value else 0,
                        efficiency=float(worksheet.cell(5, col_num).value) if worksheet.cell(5, col_num).value else 0,
                        shift_rate=float(worksheet.cell(6, col_num).value) if worksheet.cell(6, col_num).value else 0,
                        shift_rate_for_mtc=float(worksheet.cell(7, col_num).value) if worksheet.cell(7, col_num).value else 0,
                        mtc_count=int(worksheet.cell(8, col_num).value) if worksheet.cell(8, col_num).value else 0,
                        rejection_percentage=float(worksheet.cell(9, col_num).value) if worksheet.cell(9, col_num).value else 0,
                        overhead_percentage=float(worksheet.cell(10, col_num).value) if worksheet.cell(10, col_num).value else 0,
                        maintenance_percentage=float(worksheet.cell(11, col_num).value) if worksheet.cell(11, col_num).value else 0,
                        profit_percentage=float(worksheet.cell(12, col_num).value) if worksheet.cell(12, col_num).value else 0,
                    )
                    quote.increment_version(user, f'Moulding machine added via bulk upload')

                elif component_type == 'assembly':
                    Assembly.objects.create(
                        quote=quote,
                        name=worksheet.cell(2, col_num).value or '',
                        remarks=worksheet.cell(4, col_num).value or '',  # Added remarks
                        manual_cost=float(worksheet.cell(5, col_num).value) if worksheet.cell(5, col_num).value else 0,
                        other_cost=float(worksheet.cell(6, col_num).value) if worksheet.cell(6, col_num).value else 0,
                        profit_percentage=float(worksheet.cell(7, col_num).value) if worksheet.cell(7, col_num).value else 0,
                        rejection_percentage=float(worksheet.cell(8, col_num).value) if worksheet.cell(8, col_num).value else 0,
                        inspection_handling_percentage=float(worksheet.cell(9, col_num).value) if worksheet.cell(9, col_num).value else 0,
                    )
                    quote.increment_version(user, f'Assembly added via bulk upload')

                elif component_type == 'packaging':
                    Packaging.objects.create(
                        quote=quote,
                        packaging_type=worksheet.cell(2, col_num).value or 'pp_box',
                        packaging_length=float(worksheet.cell(3, col_num).value) if worksheet.cell(3, col_num).value else 600,
                        packaging_breadth=float(worksheet.cell(4, col_num).value) if worksheet.cell(4, col_num).value else 400,
                        packaging_height=float(worksheet.cell(5, col_num).value) if worksheet.cell(5, col_num).value else 250,
                        polybag_length=float(worksheet.cell(6, col_num).value) if worksheet.cell(6, col_num).value else 16,
                        polybag_width=float(worksheet.cell(7, col_num).value) if worksheet.cell(7, col_num).value else 20,
                        lifecycle=int(worksheet.cell(8, col_num).value) if worksheet.cell(8, col_num).value else 1,
                        cost=float(worksheet.cell(9, col_num).value) if worksheet.cell(9, col_num).value else 0,
                        maintenance_percentage=float(worksheet.cell(10, col_num).value) if worksheet.cell(10, col_num).value else 0,
                        part_per_polybag=int(worksheet.cell(11, col_num).value) if worksheet.cell(11, col_num).value else 1,
                    )
                    quote.increment_version(user, f'Packaging added via bulk upload')

                elif component_type == 'transport':
                    Transport.objects.create(
                        quote=quote,
                        transport_length=float(worksheet.cell(2, col_num).value) if worksheet.cell(2, col_num).value else 0,
                        transport_breadth=float(worksheet.cell(3, col_num).value) if worksheet.cell(3, col_num).value else 0,
                        transport_height=float(worksheet.cell(4, col_num).value) if worksheet.cell(4, col_num).value else 0,
                        trip_cost=float(worksheet.cell(5, col_num).value) if worksheet.cell(5, col_num).value else 0,
                        parts_per_box=int(worksheet.cell(6, col_num).value) if worksheet.cell(6, col_num).value else 1,
                    )
                    quote.increment_version(user, f'Transport added via bulk upload')

                count += 1

            except Exception as e:
                errors.append(f"{component_type.title()} Column {get_column_letter(col_num)}: {str(e)}")

        return count, errors
