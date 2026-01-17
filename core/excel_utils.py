"""
Excel utilities for quote management system
Handles Excel template generation and parsing for bulk uploads
"""

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from django.utils import timezone


class ExcelTemplateGenerator:
    """Generate Excel templates for various components"""

    @staticmethod
    def _add_vertical_headers(worksheet, headers, color):
        """Helper to add vertical headers with styling"""
        header_fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")

        for row_num, header in enumerate(headers, 1):
            cell = worksheet.cell(row=row_num, column=1)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='left', vertical='center')

        worksheet.column_dimensions['A'].width = 30

    @staticmethod
    def create_raw_materials_template():
        """Create template for raw materials upload (vertical format)"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Raw Materials"

        headers = [
            'Material Name*',
            'Grade',
            'RM Code*',
            'Unit (kg/gm/ton)*',
            'RM Rate (per kg)*',
            'Frozen Rate (per kg)',
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

        ExcelTemplateGenerator._add_vertical_headers(ws, headers, "4472C4")

        # Add sample data for 2 materials
        materials = [
            ['Polypropylene', 'PP-H340R', 'RM-PP-001', 'kg', 125.50, None, 0.0234, 0.0045, 2.5, 1.2, 0.5, 2.0, 5.0, 3.0, 10.0],
            ['ABS Resin', 'ABS-750', 'RM-ABS-002', 'kg', 185.75, 180.00, 0.0456, 0.0089, 3.0, 1.5, 0.75, 2.5, 4.5, 2.5, 12.0],
        ]

        for col_num, material in enumerate(materials, 2):
            for row_num, value in enumerate(material, 1):
                ws.cell(row=row_num, column=col_num, value=value)
            ws.column_dimensions[get_column_letter(col_num)].width = 15

        return wb

    @staticmethod
    def create_moulding_machines_template():
        """Create template for moulding machines upload (vertical format)"""
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

        ExcelTemplateGenerator._add_vertical_headers(ws, headers, "70AD47")

        # Add sample data for 2 machines
        machines = [
            [4, 180, 45.2, 90.0, 6000.00, 5500.00, 3, 2.0, 4.5, 3.0, 12.0],
            [2, 150, 42.8, 92.5, 5800.00, 5400.00, 2, 2.0, 5.0, 3.2, 12.0],
        ]

        for col_num, machine in enumerate(machines, 2):
            for row_num, value in enumerate(machine, 1):
                ws.cell(row=row_num, column=col_num, value=value)
            ws.column_dimensions[get_column_letter(col_num)].width = 15

        return wb

    @staticmethod
    def create_assemblies_template():
        """Create template for assemblies upload (vertical format)"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Assemblies"

        headers = [
            'Assembly Name*',
            'Assembly Type',
            'Remarks',
            'Manual Cost',
            'Other Cost',
            'Inspection & Handling Cost',
            'Profit %',
            'Rejection %'
        ]

        ExcelTemplateGenerator._add_vertical_headers(ws, headers, "FFC000")

        # Add sample data for 2 assemblies
        assemblies = [
            ['Manual Screw Assembly', 'Manual', 'Standard assembly', 12.50, 3.50, 5.00, 10.0, 2.5],
            ['Ultrasonic Welding', 'Automated', 'High precision', 25.00, 5.00, 8.00, 15.0, 3.0],
        ]

        for col_num, assembly in enumerate(assemblies, 2):
            for row_num, value in enumerate(assembly, 1):
                ws.cell(row=row_num, column=col_num, value=value)
            ws.column_dimensions[get_column_letter(col_num)].width = 15

        return wb

    @staticmethod
    def create_packaging_template():
        """Create template for packaging upload (vertical format)"""
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

        ExcelTemplateGenerator._add_vertical_headers(ws, headers, "E26B0A")

        # Add sample data for 2 packaging options
        packagings = [
            ['pp_box', 600, 400, 250, 16, 20, 150, 8.50, 5.0, 50],
            ['cg_box', 700, 450, 300, 18, 22, 200, 15.00, 4.0, 100],
        ]

        for col_num, packaging in enumerate(packagings, 2):
            for row_num, value in enumerate(packaging, 1):
                ws.cell(row=row_num, column=col_num, value=value)
            ws.column_dimensions[get_column_letter(col_num)].width = 15

        return wb

    @staticmethod
    def create_transport_template():
        """Create template for transport upload (vertical format)"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Transport"

        headers = [
            'Transport Length (ft)*',
            'Transport Breadth (ft)*',
            'Transport Height (ft)*',
            'Trip Cost*',
            'Parts per Box*'
        ]

        ExcelTemplateGenerator._add_vertical_headers(ws, headers, "9933FF")

        # Add sample data for 2 transport options
        transports = [
            [12, 8, 6, 6500.00, 120],
            [14, 9, 7, 8500.00, 200],
        ]

        for col_num, transport in enumerate(transports, 2):
            for row_num, value in enumerate(transport, 1):
                ws.cell(row=row_num, column=col_num, value=value)
            ws.column_dimensions[get_column_letter(col_num)].width = 15

        return wb

    @staticmethod
    def create_complete_quote_template():
        """Create comprehensive template for creating a complete quote"""
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

        # Add single quote data
        quote_data = [
            'Sample Quote',
            'Sample Client',
            'SAP-001',
            'PART-001',
            'Widget A',
            'Rev-A',
            'Sample quote description',
            1000,
            100.00,
            15.0,
            'Test notes'
        ]

        for row_num, value in enumerate(quote_data, 1):
            ws_def.cell(row=row_num, column=2, value=value)
        ws_def.column_dimensions['B'].width = 25

        # Sheet 2: Raw Materials
        ws_rm = wb.create_sheet("Raw Materials")
        rm_headers = [
            'Material Name*',
            'Grade',
            'RM Code*',
            'Unit (kg/gm/ton)*',
            'RM Rate (per kg)*',
            'Frozen Rate (per kg)',
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

        materials = [
            ['Polypropylene', 'PP-H340R', 'RM-PP-001', 'kg', 125.50, None, 0.0234, 0.0045, 2.5, 1.2, 0.5, 2.0, 5.0, 3.0, 10.0],
            ['ABS Resin', 'ABS-750', 'RM-ABS-002', 'kg', 185.75, None, 0.0456, 0.0089, 3.0, 1.5, 0.75, 2.5, 4.5, 2.5, 12.0],
        ]

        for col_num, material in enumerate(materials, 2):
            for row_num, value in enumerate(material, 1):
                ws_rm.cell(row=row_num, column=col_num, value=value)
            ws_rm.column_dimensions[get_column_letter(col_num)].width = 15

        # Sheet 3: Moulding Machines
        ws_mm = wb.create_sheet("Moulding Machines")
        mm_headers = [
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

        machines = [
            [4, 180, 45.2, 90.0, 6000.00, 5500.00, 3, 2.0, 4.5, 3.0, 12.0],
        ]

        for col_num, machine in enumerate(machines, 2):
            for row_num, value in enumerate(machine, 1):
                ws_mm.cell(row=row_num, column=col_num, value=value)
            ws_mm.column_dimensions[get_column_letter(col_num)].width = 15

        # Sheet 4: Assemblies
        ws_asm = wb.create_sheet("Assemblies")
        asm_headers = [
            'Assembly Name*',
            'Assembly Type',
            'Remarks',
            'Manual Cost',
            'Other Cost',
            'Inspection & Handling Cost',
            'Profit %',
            'Rejection %'
        ]
        ExcelTemplateGenerator._add_vertical_headers(ws_asm, asm_headers, "FFC000")

        assemblies = [
            ['Manual Screw Assembly', 'Manual', 'Standard assembly', 12.50, 3.50, 5.00, 10.0, 2.5],
        ]

        for col_num, assembly in enumerate(assemblies, 2):
            for row_num, value in enumerate(assembly, 1):
                ws_asm.cell(row=row_num, column=col_num, value=value)
            ws_asm.column_dimensions[get_column_letter(col_num)].width = 15

        # Sheet 5: Packaging
        ws_pkg = wb.create_sheet("Packaging")
        pkg_headers = [
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

        packagings = [
            ['pp_box', 600, 400, 250, 16, 20, 150, 8.50, 5.0, 50],
        ]

        for col_num, packaging in enumerate(packagings, 2):
            for row_num, value in enumerate(packaging, 1):
                ws_pkg.cell(row=row_num, column=col_num, value=value)
            ws_pkg.column_dimensions[get_column_letter(col_num)].width = 15

        # Sheet 6: Transport
        ws_trans = wb.create_sheet("Transport")
        trans_headers = [
            'Transport Length (ft)*',
            'Transport Breadth (ft)*',
            'Transport Height (ft)*',
            'Trip Cost*',
            'Parts per Box*'
        ]
        ExcelTemplateGenerator._add_vertical_headers(ws_trans, trans_headers, "9933FF")

        transports = [
            [12, 8, 6, 6500.00, 120],
        ]

        for col_num, transport in enumerate(transports, 2):
            for row_num, value in enumerate(transport, 1):
                ws_trans.cell(row=row_num, column=col_num, value=value)
            ws_trans.column_dimensions[get_column_letter(col_num)].width = 15

        return wb

    @staticmethod
    def create_multiple_quotes_template():
        """Create comprehensive template for creating multiple complete quotes in a project"""
        wb = Workbook()
        wb.remove(wb.active)

        # Instructions Sheet
        ws_inst = wb.create_sheet("Instructions", 0)
        ws_inst.column_dimensions['A'].width = 100

        instructions = [
            "MULTIPLE COMPLETE QUOTES TEMPLATE",
            "",
            "This file allows you to create multiple quotes with all components:",
            "- Quote definition",
            "- Raw materials (RM Rate in per kg)",
            "- Moulding machines",
            "- Assemblies",
            "- Packaging",
            "- Transport (dimensions in feet)",
            "",
            "STRUCTURE:",
            "- Each sheet uses vertical format (headers in column A)",
            "- Each column (B, C, D...) represents one entry",
            "- Link components to quotes using 'Quote Name' in row 1",
            "",
            "IMPORTANT NOTES:",
            "- RM Rate is entered per kg (will be converted to per gram for calculations)",
            "- Weights are converted to grams: kg→×1000, ton→×1,000,000, gm→×1",
            "- Transport dimensions are in feet (1 ft = 304.8 mm)",
            "- Inspection & Handling is a fixed cost (not percentage)",
            "",
            "Upload this to a project to create multiple quotes with all components.",
        ]

        for row_num, instruction in enumerate(instructions, 1):
            cell = ws_inst.cell(row=row_num, column=1)
            cell.value = instruction
            if row_num == 1:
                cell.font = Font(bold=True, size=14)
            elif instruction.startswith("STRUCTURE") or instruction.startswith("IMPORTANT"):
                cell.font = Font(bold=True)

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

        # Add sample quote data (2 quotes)
        quotes = [
            ['Quote 001', 'Sample Client A', 'SAP-001', 'PART-001', 'Widget A', 'Rev-A', 'Sample quote 1', 1000, 100, 15, 'Test quote'],
            ['Quote 002', 'Sample Client B', 'SAP-002', 'PART-002', 'Widget B', 'Rev-B', 'Sample quote 2', 2000, 150, 18, 'Test quote'],
        ]

        for col_num, quote_data in enumerate(quotes, 2):
            for row_num, value in enumerate(quote_data, 1):
                ws_def.cell(row=row_num, column=col_num, value=value)
            ws_def.column_dimensions[get_column_letter(col_num)].width = 20

        # Sheet 2: Raw Materials
        ws_rm = wb.create_sheet("Raw Materials")
        rm_headers = [
            'Quote Name*',
            'Material Name*',
            'Grade',
            'RM Code*',
            'Unit (kg/gm/ton)*',
            'RM Rate (per kg)*',
            'Frozen Rate (per kg)',
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

        raw_materials = [
            ['Quote 001', 'Polypropylene', 'PP-H340R', 'RM-PP-001', 'kg', 125.50, None, 0.0234, 0.0045, 2.5, 1.2, 0.5, 2.0, 5.0, 3.0, 10.0],
            ['Quote 001', 'ABS Resin', 'ABS-750', 'RM-ABS-002', 'kg', 185.75, None, 0.0456, 0.0089, 3.0, 1.5, 0.75, 2.5, 4.5, 2.5, 12.0],
            ['Quote 002', 'Nylon 6', 'PA6-GF30', 'RM-PA6-003', 'kg', 215.00, None, 0.0678, 0.0123, 4.5, 2.0, 1.0, 3.0, 6.0, 4.0, 15.0],
        ]

        for col_num, material in enumerate(raw_materials, 2):
            for row_num, value in enumerate(material, 1):
                ws_rm.cell(row=row_num, column=col_num, value=value)
            ws_rm.column_dimensions[get_column_letter(col_num)].width = 18

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

        machines = [
            ['Quote 001', 4, 180, 45.2, 90.0, 6000.00, 5500.00, 3, 2.0, 4.5, 3.0, 12.0],
            ['Quote 002', 2, 150, 42.8, 92.5, 5800.00, 5400.00, 2, 2.0, 5.0, 3.2, 12.0],
        ]

        for col_num, machine in enumerate(machines, 2):
            for row_num, value in enumerate(machine, 1):
                ws_mm.cell(row=row_num, column=col_num, value=value)
            ws_mm.column_dimensions[get_column_letter(col_num)].width = 18

        # Sheet 4: Assemblies
        ws_asm = wb.create_sheet("Assemblies")
        asm_headers = [
            'Quote Name*',
            'Assembly Name*',
            'Assembly Type',
            'Remarks',
            'Manual Cost',
            'Other Cost',
            'Inspection & Handling Cost',
            'Profit %',
            'Rejection %'
        ]
        ExcelTemplateGenerator._add_vertical_headers(ws_asm, asm_headers, "FFC000")

        assemblies = [
            ['Quote 001', 'Manual Screw Assembly', 'Manual', 'Standard assembly', 12.50, 3.50, 5.00, 10.0, 2.5],
            ['Quote 002', 'Ultrasonic Welding', 'Automated', 'High precision', 25.00, 5.00, 8.00, 15.0, 3.0],
        ]

        for col_num, assembly in enumerate(assemblies, 2):
            for row_num, value in enumerate(assembly, 1):
                ws_asm.cell(row=row_num, column=col_num, value=value)
            ws_asm.column_dimensions[get_column_letter(col_num)].width = 18

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

        packagings = [
            ['Quote 001', 'pp_box', 600, 400, 250, 16, 20, 150, 8.50, 5.0, 50],
            ['Quote 002', 'cg_box', 700, 450, 300, 18, 22, 200, 15.00, 4.0, 100],
        ]

        for col_num, packaging in enumerate(packagings, 2):
            for row_num, value in enumerate(packaging, 1):
                ws_pkg.cell(row=row_num, column=col_num, value=value)
            ws_pkg.column_dimensions[get_column_letter(col_num)].width = 18

        # Sheet 6: Transport
        ws_trans = wb.create_sheet("Transport")
        trans_headers = [
            'Quote Name*',
            'Transport Length (ft)*',
            'Transport Breadth (ft)*',
            'Transport Height (ft)*',
            'Trip Cost*',
            'Parts per Box*'
        ]
        ExcelTemplateGenerator._add_vertical_headers(ws_trans, trans_headers, "9933FF")

        transports = [
            ['Quote 001', 12, 8, 6, 6500.00, 120],
            ['Quote 002', 14, 9, 7, 8500.00, 200],
        ]

        for col_num, transport in enumerate(transports, 2):
            for row_num, value in enumerate(transport, 1):
                ws_trans.cell(row=row_num, column=col_num, value=value)
            ws_trans.column_dimensions[get_column_letter(col_num)].width = 18

        return wb


class ExcelParser:
    """Parse Excel files and create database objects"""

    @staticmethod
    def parse_complete_quote(file, project, customer_group, user):
        """Parse a complete quote from Excel file"""
        from core.models import Quote, QuoteTimeline

        wb = openpyxl.load_workbook(file)

        # Parse Quote Definition
        if "Quote Definition" not in wb.sheetnames:
            return None, ['Quote Definition sheet not found']

        ws = wb["Quote Definition"]

        try:
            quote = Quote.objects.create(
                project=project,
                name=ws.cell(1, 2).value or '',
                client_group=customer_group,
                client_name=ws.cell(2, 2).value or '',
                sap_number=ws.cell(3, 2).value or '',
                part_number=ws.cell(4, 2).value or '',
                part_name=ws.cell(5, 2).value or '',
                amendment_number=ws.cell(6, 2).value or '',
                description=ws.cell(7, 2).value or '',
                quantity=int(ws.cell(8, 2).value) if ws.cell(8, 2).value else 1,
                handling_charge=float(ws.cell(9, 2).value) if ws.cell(9, 2).value else 0,
                profit_percentage=float(ws.cell(10, 2).value) if ws.cell(10, 2).value else 0,
                notes=ws.cell(11, 2).value or '',
                created_by=user,
                quote_definition_complete=True
            )

            # Add timeline entry
            QuoteTimeline.add_entry(
                quote=quote,
                activity_type='quote_created',
                description=f'Quote "{quote.name}" created via complete upload',
                user=user
            )

            # Parse components
            errors = []
            component_counts = {}

            # Raw Materials
            if "Raw Materials" in wb.sheetnames:
                count, comp_errors = ExcelParser._parse_components(wb["Raw Materials"], quote, 'raw_material', user)
                component_counts['raw_materials'] = count
                errors.extend(comp_errors)

            # Moulding Machines
            if "Moulding Machines" in wb.sheetnames:
                count, comp_errors = ExcelParser._parse_components(wb["Moulding Machines"], quote, 'moulding_machine', user)
                component_counts['moulding_machines'] = count
                errors.extend(comp_errors)

            # Assemblies
            if "Assemblies" in wb.sheetnames:
                count, comp_errors = ExcelParser._parse_components(wb["Assemblies"], quote, 'assembly', user)
                component_counts['assemblies'] = count
                errors.extend(comp_errors)

            # Packaging
            if "Packaging" in wb.sheetnames:
                count, comp_errors = ExcelParser._parse_components(wb["Packaging"], quote, 'packaging', user)
                component_counts['packaging'] = count
                errors.extend(comp_errors)

            # Transport
            if "Transport" in wb.sheetnames:
                count, comp_errors = ExcelParser._parse_components(wb["Transport"], quote, 'transport', user)
                component_counts['transport'] = count
                errors.extend(comp_errors)

            return quote, errors, component_counts

        except Exception as e:
            return None, [f'Error creating quote: {str(e)}']

    @staticmethod
    def _parse_components(worksheet, quote, component_type, user):
        """Parse components from worksheet"""
        from core.models import RawMaterial, MouldingMachineDetail, Assembly, Packaging, Transport

        count = 0
        errors = []
        max_col = worksheet.max_column

        for col_num in range(2, max_col + 1):
            try:
                if component_type == 'raw_material':
                    # Check if first cell has data
                    if not worksheet.cell(1, col_num).value:
                        continue

                    RawMaterial.objects.create(
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
                    count += 1

                elif component_type == 'moulding_machine':
                    if not worksheet.cell(1, col_num).value:
                        continue

                    MouldingMachineDetail.objects.create(
                        quote=quote,
                        cavity=int(worksheet.cell(1, col_num).value) if worksheet.cell(1, col_num).value else 1,
                        machine_tonnage=float(worksheet.cell(2, col_num).value) if worksheet.cell(2, col_num).value else 0,
                        cycle_time=float(worksheet.cell(3, col_num).value) if worksheet.cell(3, col_num).value else 0,
                        efficiency_percentage=float(worksheet.cell(4, col_num).value) if worksheet.cell(4, col_num).value else 0,
                        shift_rate=float(worksheet.cell(5, col_num).value) if worksheet.cell(5, col_num).value else 0,
                        shift_rate_for_mtc=float(worksheet.cell(6, col_num).value) if worksheet.cell(6, col_num).value else 0,
                        mtc_count=int(worksheet.cell(7, col_num).value) if worksheet.cell(7, col_num).value else 0,
                        rejection_percentage=float(worksheet.cell(8, col_num).value) if worksheet.cell(8, col_num).value else 0,
                        overhead_percentage=float(worksheet.cell(9, col_num).value) if worksheet.cell(9, col_num).value else 0,
                        maintenance_percentage=float(worksheet.cell(10, col_num).value) if worksheet.cell(10, col_num).value else 0,
                        profit_percentage=float(worksheet.cell(11, col_num).value) if worksheet.cell(11, col_num).value else 0,
                    )
                    count += 1

                elif component_type == 'assembly':
                    if not worksheet.cell(1, col_num).value:
                        continue

                    Assembly.objects.create(
                        quote=quote,
                        name=worksheet.cell(1, col_num).value or '',
                        remarks=worksheet.cell(3, col_num).value or '',
                        manual_cost=float(worksheet.cell(4, col_num).value) if worksheet.cell(4, col_num).value else 0,
                        other_cost=float(worksheet.cell(5, col_num).value) if worksheet.cell(5, col_num).value else 0,
                        inspection_handling_cost=float(worksheet.cell(6, col_num).value) if worksheet.cell(6, col_num).value else 0,
                        profit_percentage=float(worksheet.cell(7, col_num).value) if worksheet.cell(7, col_num).value else 0,
                        rejection_percentage=float(worksheet.cell(8, col_num).value) if worksheet.cell(8, col_num).value else 0,
                    )
                    count += 1

                elif component_type == 'packaging':
                    if not worksheet.cell(1, col_num).value:
                        continue

                    Packaging.objects.create(
                        quote=quote,
                        packaging_type=worksheet.cell(1, col_num).value or 'pp_box',
                        packaging_length=float(worksheet.cell(2, col_num).value) if worksheet.cell(2, col_num).value else 0,
                        packaging_breadth=float(worksheet.cell(3, col_num).value) if worksheet.cell(3, col_num).value else 0,
                        packaging_height=float(worksheet.cell(4, col_num).value) if worksheet.cell(4, col_num).value else 0,
                        polybag_length=float(worksheet.cell(5, col_num).value) if worksheet.cell(5, col_num).value else 0,
                        polybag_width=float(worksheet.cell(6, col_num).value) if worksheet.cell(6, col_num).value else 0,
                        lifecycle=int(worksheet.cell(7, col_num).value) if worksheet.cell(7, col_num).value else 0,
                        cost=float(worksheet.cell(8, col_num).value) if worksheet.cell(8, col_num).value else 0,
                        maintenance_percentage=float(worksheet.cell(9, col_num).value) if worksheet.cell(9, col_num).value else 0,
                        parts_per_polybag=int(worksheet.cell(10, col_num).value) if worksheet.cell(10, col_num).value else 0,
                    )
                    count += 1

                elif component_type == 'transport':
                    if not worksheet.cell(1, col_num).value:
                        continue

                    Transport.objects.create(
                        quote=quote,
                        transport_length=float(worksheet.cell(1, col_num).value) if worksheet.cell(1, col_num).value else 0,
                        transport_breadth=float(worksheet.cell(2, col_num).value) if worksheet.cell(2, col_num).value else 0,
                        transport_height=float(worksheet.cell(3, col_num).value) if worksheet.cell(3, col_num).value else 0,
                        trip_cost=float(worksheet.cell(4, col_num).value) if worksheet.cell(4, col_num).value else 0,
                        parts_per_box=int(worksheet.cell(5, col_num).value) if worksheet.cell(5, col_num).value else 0,
                    )
                    count += 1

            except Exception as e:
                errors.append(f"{component_type.title()} Column {get_column_letter(col_num)}: {str(e)}")

        if count > 0:
            quote.increment_version(user, f'{count} {component_type}(s) added via upload')

        return count, errors

    @staticmethod
    def parse_multiple_quotes_complete(file, project, customer_group, user):
        """Parse multiple complete quotes with all components from Excel file"""
        from core.models import Quote, QuoteTimeline

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

        # Parse components for each sheet
        if "Raw Materials" in wb.sheetnames:
            ws = wb["Raw Materials"]
            count, sheet_errors = ExcelParser._parse_components_by_quote(
                ws, quotes_dict, 'raw_material', user
            )
            component_counts['raw_materials'] = count
            errors.extend(sheet_errors)

        if "Moulding Machines" in wb.sheetnames:
            ws = wb["Moulding Machines"]
            count, sheet_errors = ExcelParser._parse_components_by_quote(
                ws, quotes_dict, 'moulding_machine', user
            )
            component_counts['moulding_machines'] = count
            errors.extend(sheet_errors)

        if "Assemblies" in wb.sheetnames:
            ws = wb["Assemblies"]
            count, sheet_errors = ExcelParser._parse_components_by_quote(
                ws, quotes_dict, 'assembly', user
            )
            component_counts['assemblies'] = count
            errors.extend(sheet_errors)

        if "Packaging" in wb.sheetnames:
            ws = wb["Packaging"]
            count, sheet_errors = ExcelParser._parse_components_by_quote(
                ws, quotes_dict, 'packaging', user
            )
            component_counts['packaging'] = count
            errors.extend(sheet_errors)

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
        """Parse components that belong to specific quotes"""
        from core.models import RawMaterial, MouldingMachineDetail, Assembly, Packaging, Transport

        count = 0
        errors = []
        max_col = worksheet.max_column

        for col_num in range(2, max_col + 1):
            quote_name = worksheet.cell(1, col_num).value

            if not quote_name:
                continue

            if quote_name not in quotes_dict:
                errors.append(f"{component_type.title()} Column {get_column_letter(col_num)}: Quote '{quote_name}' not found")
                continue

            quote = quotes_dict[quote_name]

            try:
                if component_type == 'raw_material':
                    # Material name is in row 2 for multiple quotes format
                    if not worksheet.cell(2, col_num).value:
                        continue

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
                    count += 1

                elif component_type == 'moulding_machine':
                    if not worksheet.cell(2, col_num).value:
                        continue

                    MouldingMachineDetail.objects.create(
                        quote=quote,
                        cavity=int(worksheet.cell(2, col_num).value) if worksheet.cell(2, col_num).value else 1,
                        machine_tonnage=float(worksheet.cell(3, col_num).value) if worksheet.cell(3, col_num).value else 0,
                        cycle_time=float(worksheet.cell(4, col_num).value) if worksheet.cell(4, col_num).value else 0,
                        efficiency_percentage=float(worksheet.cell(5, col_num).value) if worksheet.cell(5, col_num).value else 0,
                        shift_rate=float(worksheet.cell(6, col_num).value) if worksheet.cell(6, col_num).value else 0,
                        shift_rate_for_mtc=float(worksheet.cell(7, col_num).value) if worksheet.cell(7, col_num).value else 0,
                        mtc_count=int(worksheet.cell(8, col_num).value) if worksheet.cell(8, col_num).value else 0,
                        rejection_percentage=float(worksheet.cell(9, col_num).value) if worksheet.cell(9, col_num).value else 0,
                        overhead_percentage=float(worksheet.cell(10, col_num).value) if worksheet.cell(10, col_num).value else 0,
                        maintenance_percentage=float(worksheet.cell(11, col_num).value) if worksheet.cell(11, col_num).value else 0,
                        profit_percentage=float(worksheet.cell(12, col_num).value) if worksheet.cell(12, col_num).value else 0,
                    )
                    count += 1

                elif component_type == 'assembly':
                    if not worksheet.cell(2, col_num).value:
                        continue

                    Assembly.objects.create(
                        quote=quote,
                        name=worksheet.cell(2, col_num).value or '',
                        remarks=worksheet.cell(4, col_num).value or '',
                        manual_cost=float(worksheet.cell(5, col_num).value) if worksheet.cell(5, col_num).value else 0,
                        other_cost=float(worksheet.cell(6, col_num).value) if worksheet.cell(6, col_num).value else 0,
                        inspection_handling_cost=float(worksheet.cell(7, col_num).value) if worksheet.cell(7, col_num).value else 0,
                        profit_percentage=float(worksheet.cell(8, col_num).value) if worksheet.cell(8, col_num).value else 0,
                        rejection_percentage=float(worksheet.cell(9, col_num).value) if worksheet.cell(9, col_num).value else 0,
                    )
                    count += 1

                elif component_type == 'packaging':
                    if not worksheet.cell(2, col_num).value:
                        continue

                    Packaging.objects.create(
                        quote=quote,
                        packaging_type=worksheet.cell(2, col_num).value or 'pp_box',
                        packaging_length=float(worksheet.cell(3, col_num).value) if worksheet.cell(3, col_num).value else 0,
                        packaging_breadth=float(worksheet.cell(4, col_num).value) if worksheet.cell(4, col_num).value else 0,
                        packaging_height=float(worksheet.cell(5, col_num).value) if worksheet.cell(5, col_num).value else 0,
                        polybag_length=float(worksheet.cell(6, col_num).value) if worksheet.cell(6, col_num).value else 0,
                        polybag_width=float(worksheet.cell(7, col_num).value) if worksheet.cell(7, col_num).value else 0,
                        lifecycle=int(worksheet.cell(8, col_num).value) if worksheet.cell(8, col_num).value else 0,
                        cost=float(worksheet.cell(9, col_num).value) if worksheet.cell(9, col_num).value else 0,
                        maintenance_percentage=float(worksheet.cell(10, col_num).value) if worksheet.cell(10, col_num).value else 0,
                        parts_per_polybag=int(worksheet.cell(11, col_num).value) if worksheet.cell(11, col_num).value else 0,
                    )
                    count += 1

                elif component_type == 'transport':
                    if not worksheet.cell(2, col_num).value:
                        continue

                    Transport.objects.create(
                        quote=quote,
                        transport_length=float(worksheet.cell(2, col_num).value) if worksheet.cell(2, col_num).value else 0,
                        transport_breadth=float(worksheet.cell(3, col_num).value) if worksheet.cell(3, col_num).value else 0,
                        transport_height=float(worksheet.cell(4, col_num).value) if worksheet.cell(4, col_num).value else 0,
                        trip_cost=float(worksheet.cell(5, col_num).value) if worksheet.cell(5, col_num).value else 0,
                        parts_per_box=int(worksheet.cell(6, col_num).value) if worksheet.cell(6, col_num).value else 0,
                    )
                    count += 1

            except Exception as e:
                errors.append(f"{component_type.title()} Column {get_column_letter(col_num)}: {str(e)}")

        return count, errors

class ExcelExporter:
    """Export quotes and projects to Excel with all calculated fields in vertical format"""

    @staticmethod
    def export_quote(quote):
        """Export a single quote with all data and calculated fields (vertical format)"""
        wb = Workbook()
        wb.remove(wb.active)

        # Sheet 1: Quote Definition
        ws_def = wb.create_sheet("Quote Definition")

        # Get version string
        version_str = f"{quote.version_major}.{quote.version_minor}" if hasattr(quote, 'version_major') else 'N/A'

        # Define headers and data
        definition_data = [
            ['Quote Name', quote.name],
            ['Client Name', quote.client_name],
            ['Client Group', quote.client_group.name if quote.client_group else ''],
            ['SAP Number', quote.sap_number],
            ['Part Number', quote.part_number],
            ['Part Name', quote.part_name],
            ['Amendment Number', quote.amendment_number],
            ['Description', quote.description],
            ['Quantity', quote.quantity],
            ['Handling Charge', float(quote.handling_charge)],
            ['Profit %', float(quote.profit_percentage)],
            ['Notes', quote.notes],
            ['Status', quote.get_status_display()],
            ['Version', version_str],
            ['Created By', quote.created_by.username if quote.created_by else ''],
            ['Created At', quote.created_at.strftime('%Y-%m-%d %H:%M:%S')],
            ['Updated At', quote.updated_at.strftime('%Y-%m-%d %H:%M:%S')],
        ]

        ExcelExporter._add_vertical_data(ws_def, definition_data, "2E75B6")

        # Sheet 2: Raw Materials (vertical format with calculated fields)
        if quote.raw_materials.exists():
            ws_rm = wb.create_sheet("Raw Materials")
            rm_headers = [
                'Material Name',
                'Grade',
                'RM Code',
                'Unit',
                'RM Rate (per kg)',
                'Frozen Rate (per kg)',
                'Effective Rate (per kg)',
                'Part Weight',
                'Runner Weight',
                'Gross Weight (grams)',
                'Process Losses',
                'Purging Loss Cost',
                'ICC %',
                'Rejection %',
                'Overhead %',
                'Maintenance %',
                'Profit %',
                'Base RM Cost',
                'Total Cost'
            ]

            # Add headers in column A
            ExcelExporter._add_vertical_headers(ws_rm, rm_headers, "4472C4")

            # Add data for each raw material in subsequent columns
            for col_num, rm in enumerate(quote.raw_materials.all(), 2):
                ws_rm.cell(row=1, column=col_num, value=rm.material_name)
                ws_rm.cell(row=2, column=col_num, value=rm.grade)
                ws_rm.cell(row=3, column=col_num, value=rm.rm_code)
                ws_rm.cell(row=4, column=col_num, value=rm.unit_of_measurement)
                ws_rm.cell(row=5, column=col_num, value=float(rm.rm_rate))
                ws_rm.cell(row=6, column=col_num, value=float(rm.frozen_rate) if rm.frozen_rate else None)
                ws_rm.cell(row=7, column=col_num, value=float(rm.effective_rate_per_kg))
                ws_rm.cell(row=8, column=col_num, value=float(rm.part_weight))
                ws_rm.cell(row=9, column=col_num, value=float(rm.runner_weight))
                ws_rm.cell(row=10, column=col_num, value=float(rm.gross_weight_in_grams))
                ws_rm.cell(row=11, column=col_num, value=float(rm.process_losses))
                ws_rm.cell(row=12, column=col_num, value=float(rm.purging_loss_cost))
                ws_rm.cell(row=13, column=col_num, value=float(rm.icc_percentage))
                ws_rm.cell(row=14, column=col_num, value=float(rm.rejection_percentage))
                ws_rm.cell(row=15, column=col_num, value=float(rm.overhead_percentage))
                ws_rm.cell(row=16, column=col_num, value=float(rm.maintenance_percentage))
                ws_rm.cell(row=17, column=col_num, value=float(rm.profit_percentage))
                ws_rm.cell(row=18, column=col_num, value=float(rm.base_rm_cost))
                ws_rm.cell(row=19, column=col_num, value=float(rm.rm_cost))
                ws_rm.column_dimensions[get_column_letter(col_num)].width = 18

        # Sheet 3: Moulding Machines (vertical format with calculated fields)
        if quote.moulding_machines.exists():
            ws_mm = wb.create_sheet("Moulding Machines")
            mm_headers = [
                'Cavity',
                'Machine Tonnage',
                'Cycle Time (s)',
                'Efficiency %',
                'Shift Rate',
                'Shift Rate for MTC',
                'MTC Count',
                'Rejection %',
                'Overhead %',
                'Maintenance %',
                'Profit %',
                'Parts per Shift',
                'MTC Cost',
                'Total Cost'
            ]

            ExcelExporter._add_vertical_headers(ws_mm, mm_headers, "70AD47")

            for col_num, mm in enumerate(quote.moulding_machines.all(), 2):
                ws_mm.cell(row=1, column=col_num, value=mm.cavity)
                ws_mm.cell(row=2, column=col_num, value=float(mm.machine_tonnage))
                ws_mm.cell(row=3, column=col_num, value=float(mm.cycle_time))
                ws_mm.cell(row=4, column=col_num, value=float(mm.efficiency))
                ws_mm.cell(row=5, column=col_num, value=float(mm.shift_rate))
                ws_mm.cell(row=6, column=col_num, value=float(mm.shift_rate_for_mtc))
                ws_mm.cell(row=7, column=col_num, value=mm.mtc_count)
                ws_mm.cell(row=8, column=col_num, value=float(mm.rejection_percentage))
                ws_mm.cell(row=9, column=col_num, value=float(mm.overhead_percentage))
                ws_mm.cell(row=10, column=col_num, value=float(mm.maintenance_percentage))
                ws_mm.cell(row=11, column=col_num, value=float(mm.profit_percentage))
                ws_mm.cell(row=13, column=col_num, value=float(mm.number_of_parts_per_shift))
                ws_mm.cell(row=14, column=col_num, value=float(mm.mtc_cost))
                ws_mm.cell(row=16, column=col_num, value=float(mm.conversion_cost))
                ws_mm.column_dimensions[get_column_letter(col_num)].width = 18

        # Sheet 4: Assemblies (vertical format with calculated fields)
        if quote.assemblies.exists():
            ws_asm = wb.create_sheet("Assemblies")
            asm_headers = [
                'Assembly Name',
                'Remarks',
                'Manual Cost',
                'Other Cost',
                'Inspection & Handling Cost',
                'Profit %',
                'Rejection %',
                'Base Cost',
                'Profit Cost',
                'Rejection Cost',
                'Total Cost'
            ]

            ExcelExporter._add_vertical_headers(ws_asm, asm_headers, "FFC000")

            for col_num, asm in enumerate(quote.assemblies.all(), 2):
                costs = asm.calculate_costs()
                ws_asm.cell(row=1, column=col_num, value=asm.name)
                ws_asm.cell(row=2, column=col_num, value=asm.remarks)
                ws_asm.cell(row=3, column=col_num, value=float(asm.manual_cost))
                ws_asm.cell(row=4, column=col_num, value=float(asm.other_cost))
                ws_asm.cell(row=5, column=col_num, value=float(asm.inspection_handling_cost))
                ws_asm.cell(row=6, column=col_num, value=float(asm.profit_percentage))
                ws_asm.cell(row=7, column=col_num, value=float(asm.rejection_percentage))
                ws_asm.cell(row=8, column=col_num, value=float(costs['base_cost']))
                ws_asm.cell(row=9, column=col_num, value=float(costs['profit_cost']))
                ws_asm.cell(row=10, column=col_num, value=float(costs['rejection_cost']))
                ws_asm.cell(row=11, column=col_num, value=float(asm.total_assembly_cost))
                ws_asm.column_dimensions[get_column_letter(col_num)].width = 18

            # Sheet 4a: Assembly Raw Materials (vertical format)
            has_assembly_rms = any(
                hasattr(asm, 'assembly_raw_materials') and asm.assembly_raw_materials.exists()
                for asm in quote.assemblies.all()
            )

            if has_assembly_rms:
                ws_arm = wb.create_sheet("Assembly Raw Materials")
                arm_headers = [
                    'Assembly Name',
                    'Description',
                    'Production Quantity',
                    'Production Weight',
                    'Unit',
                    'Cost per Unit',
                    'Total Cost'
                ]

                ExcelExporter._add_vertical_headers(ws_arm, arm_headers, "FF6B6B")

                col_num = 2
                for asm in quote.assemblies.all():
                    if hasattr(asm, 'assembly_raw_materials'):
                        for arm in asm.assembly_raw_materials.all():
                            ws_arm.cell(row=1, column=col_num, value=asm.name)
                            ws_arm.cell(row=2, column=col_num, value=arm.description)
                            ws_arm.cell(row=3, column=col_num, value=arm.production_quantity)
                            ws_arm.cell(row=4, column=col_num, value=arm.production_weight)
                            ws_arm.cell(row=5, column=col_num, value=arm.unit)
                            ws_arm.cell(row=6, column=col_num, value=float(arm.cost_per_unit))
                            ws_arm.cell(row=7, column=col_num, value=float(arm.total_cost))
                            ws_arm.column_dimensions[get_column_letter(col_num)].width = 18
                            col_num += 1

            # Sheet 4b: Manufacturing Costs (vertical format)
            has_manufacturing_costs = any(
                hasattr(asm, 'manufacturing_printing_costs') and asm.manufacturing_printing_costs.exists()
                for asm in quote.assemblies.all()
            )

            if has_manufacturing_costs:
                ws_mpc = wb.create_sheet("Manufacturing Costs")
                mpc_headers = [
                    'Assembly Name',
                    'Process',
                    'Per Cost'
                ]

                ExcelExporter._add_vertical_headers(ws_mpc, mpc_headers, "9B59B6")

                col_num = 2
                for asm in quote.assemblies.all():
                    if hasattr(asm, 'manufacturing_printing_costs'):
                        for mpc in asm.manufacturing_printing_costs.all():
                            ws_mpc.cell(row=1, column=col_num, value=asm.name)
                            ws_mpc.cell(row=2, column=col_num, value=mpc.process)
                            ws_mpc.cell(row=3, column=col_num, value=float(mpc.per_cost))
                            ws_mpc.column_dimensions[get_column_letter(col_num)].width = 18
                            col_num += 1

        # Sheet 5: Packaging (vertical format with calculated fields)
        if quote.packagings.exists():
            ws_pkg = wb.create_sheet("Packaging")
            pkg_headers = [
                'Packaging Type',
                'Length (mm)',
                'Breadth (mm)',
                'Height (mm)',
                'Polybag Length',
                'Polybag Width',
                'Lifecycle',
                'Cost',
                'Maintenance %',
                'Parts per Polybag',
                'Maintenance Cost',
                'Total Cost',
                'Cost per Part'
            ]

            ExcelExporter._add_vertical_headers(ws_pkg, pkg_headers, "E26B0A")

            for col_num, pkg in enumerate(quote.packagings.all(), 2):
                ws_pkg.cell(row=1, column=col_num, value=pkg.packaging_type.name if pkg.packaging_type else 'Custom')
                ws_pkg.cell(row=2, column=col_num, value=float(pkg.packaging_length))
                ws_pkg.cell(row=3, column=col_num, value=float(pkg.packaging_breadth))
                ws_pkg.cell(row=4, column=col_num, value=float(pkg.packaging_height))
                ws_pkg.cell(row=5, column=col_num, value=float(pkg.polybag_length))
                ws_pkg.cell(row=6, column=col_num, value=float(pkg.polybag_width))
                ws_pkg.cell(row=7, column=col_num, value=pkg.lifecycle)
                ws_pkg.cell(row=8, column=col_num, value=float(pkg.cost))
                ws_pkg.cell(row=9, column=col_num, value=float(pkg.maintenance_percentage))
                ws_pkg.cell(row=10, column=col_num, value=pkg.parts_per_polybag)
                ws_pkg.cell(row=11, column=col_num, value=float(pkg.maintenance_cost))
                ws_pkg.cell(row=12, column=col_num, value=float(pkg.total_cost))
                ws_pkg.cell(row=13, column=col_num, value=float(pkg.cost_per_part))
                ws_pkg.column_dimensions[get_column_letter(col_num)].width = 18

        # Sheet 6: Transport (vertical format with calculated fields)
        if quote.transports.exists():
            ws_trans = wb.create_sheet("Transport")
            trans_headers = [
                'Length (ft)',
                'Breadth (ft)',
                'Height (ft)',
                'Length (mm)',
                'Breadth (mm)',
                'Height (mm)',
                'Trip Cost',
                'Parts per Box',
                'Total Boxes',
                'Cost per Part'
            ]

            ExcelExporter._add_vertical_headers(ws_trans, trans_headers, "9933FF")

            for col_num, trans in enumerate(quote.transports.all(), 2):
                ws_trans.cell(row=1, column=col_num, value=float(trans.transport_length))
                ws_trans.cell(row=2, column=col_num, value=float(trans.transport_breadth))
                ws_trans.cell(row=3, column=col_num, value=float(trans.transport_height))
                ws_trans.cell(row=4, column=col_num, value=float(trans.transport_length_mm))
                ws_trans.cell(row=5, column=col_num, value=float(trans.transport_breadth_mm))
                ws_trans.cell(row=6, column=col_num, value=float(trans.transport_height_mm))
                ws_trans.cell(row=7, column=col_num, value=float(trans.trip_cost))
                ws_trans.cell(row=8, column=col_num, value=trans.parts_per_box)
                ws_trans.cell(row=9, column=col_num, value=float(trans.total_boxes))
                ws_trans.cell(row=10, column=col_num, value=float(trans.trip_cost_per_part))
                ws_trans.column_dimensions[get_column_letter(col_num)].width = 18

        # Sheet 7: Summary (vertical format)
        ws_summary = wb.create_sheet("Summary")

        # Calculate costs safely
        try:
            rm_cost = float(quote.get_total_rm_cost())
        except:
            rm_cost = 0

        try:
            moulding_cost = float(quote.get_total_moulding_cost())
        except:
            moulding_cost = 0

        try:
            assembly_cost = float(quote.get_total_assembly_cost())
        except:
            assembly_cost = 0

        try:
            packaging_cost = float(quote.get_total_packaging_cost())
        except:
            packaging_cost = 0

        try:
            transport_cost = float(quote.get_total_transport_cost())
        except:
            transport_cost = 0

        try:
            profit_amount = float(quote.get_profit_amount())
        except:
            profit_amount = 0

        try:
            grand_total = float(quote.get_grand_total())
        except:
            grand_total = 0

        # Calculate cost per part
        cost_per_part = grand_total / quote.quantity if quote.quantity > 0 else 0

        summary_data = [
            ['Raw Material Cost', rm_cost],
            ['Moulding Machine Cost', moulding_cost],
            ['Assembly Cost', assembly_cost],
            ['Packaging Cost', packaging_cost],
            ['Transport Cost', transport_cost],
            ['Handling Charge', float(quote.handling_charge)],
            ['Profit Amount', profit_amount],
            ['Grand Total', grand_total],
            ['Quantity', quote.quantity],
            ['Cost per Part', cost_per_part],
        ]

        ExcelExporter._add_vertical_data(ws_summary, summary_data, "FF0000")

        return wb

    @staticmethod
    def export_project(project):
        """Export all quotes in a project (vertical format)"""
        wb = Workbook()
        wb.remove(wb.active)

        # Sheet 1: Project Information
        ws_proj = wb.create_sheet("Project Information")

        proj_data = [
            ['Project Name', project.name],
            ['Description', project.description],
            ['Created By', project.created_by.username if project.created_by else ''],
            ['Created At', project.created_at.strftime('%Y-%m-%d %H:%M:%S')],
            ['Total Quotes', project.quotes.count()],
        ]

        ExcelExporter._add_vertical_data(ws_proj, proj_data, "2E75B6")

        # Sheet 2: Quotes Summary (vertical format)
        ws_quotes = wb.create_sheet("Quotes Summary")
        quotes_headers = [
            'Quote Name',
            'Status',
            'Version',
            'Client Name',
            'Part Number',
            'Quantity',
            'Grand Total'
        ]

        ExcelExporter._add_vertical_headers(ws_quotes, quotes_headers, "70AD47")

        for col_num, quote in enumerate(project.quotes.all(), 2):
            # Get version string (major.minor format)
            version_str = f"{quote.version_major}.{quote.version_minor}" if hasattr(quote, 'version_major') else 'N/A'

            # Calculate grand total safely
            try:
                grand_total = float(quote.get_grand_total())
            except:
                grand_total = 0

            ws_quotes.cell(row=1, column=col_num, value=quote.name)
            ws_quotes.cell(row=2, column=col_num, value=quote.get_status_display())
            ws_quotes.cell(row=3, column=col_num, value=version_str)
            ws_quotes.cell(row=4, column=col_num, value=quote.client_name)
            ws_quotes.cell(row=5, column=col_num, value=quote.part_number)
            ws_quotes.cell(row=6, column=col_num, value=quote.quantity)
            ws_quotes.cell(row=7, column=col_num, value=grand_total)
            ws_quotes.column_dimensions[get_column_letter(col_num)].width = 20

        # Add each quote as separate sheets
        for quote in project.quotes.all():
            try:
                quote_wb = ExcelExporter.export_quote(quote)
                for sheet_name in quote_wb.sheetnames:
                    source_sheet = quote_wb[sheet_name]
                    # Limit sheet name to 31 characters (Excel limit) and make it unique
                    safe_sheet_name = f"{quote.name[:10]} - {sheet_name}"[:31]
                    target_sheet = wb.create_sheet(safe_sheet_name)

                    # Copy all cells with formatting
                    for row in source_sheet.iter_rows():
                        for cell in row:
                            target_cell = target_sheet[cell.coordinate]
                            target_cell.value = cell.value
                            if cell.has_style:
                                target_cell.font = cell.font.copy()
                                target_cell.fill = cell.fill.copy()
                                target_cell.alignment = cell.alignment.copy()

                    # Copy column widths
                    for col_letter in source_sheet.column_dimensions:
                        if col_letter in source_sheet.column_dimensions:
                            target_sheet.column_dimensions[col_letter].width = source_sheet.column_dimensions[col_letter].width

            except Exception as e:
                # If a quote export fails, continue with others
                print(f"Error exporting quote {quote.name}: {str(e)}")
                continue

        return wb

    @staticmethod
    def _add_vertical_headers(worksheet, headers, color):
        """Add vertical headers in column A with styling"""
        header_fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")

        for row_num, header in enumerate(headers, 1):
            cell = worksheet.cell(row=row_num, column=1)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='left', vertical='center')

        worksheet.column_dimensions['A'].width = 30

    @staticmethod
    def _add_vertical_data(worksheet, data_rows, color):
        """Add vertical data (headers in column A, values in column B)"""
        header_fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")

        for row_num, (header, value) in enumerate(data_rows, 1):
            # Header in column A
            header_cell = worksheet.cell(row=row_num, column=1)
            header_cell.value = header
            header_cell.fill = header_fill
            header_cell.font = header_font
            header_cell.alignment = Alignment(horizontal='left', vertical='center')

            # Value in column B
            value_cell = worksheet.cell(row=row_num, column=2)
            value_cell.value = value

        worksheet.column_dimensions['A'].width = 30
        worksheet.column_dimensions['B'].width = 40