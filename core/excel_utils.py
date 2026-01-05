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
        """Create template for raw materials"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Raw Materials"
        
        # Headers
        headers = [
            'Material Name*', 'Grade', 'RM Code*', 'Unit (kg/gm/ton)*',
            'RM Rate*', 'Frozen Rate', 'Part Weight*', 'Runner Weight*',
            'Process Losses', 'Purging Loss Cost', 'ICC %',
            'Rejection %', 'Overhead %', 'Maintenance %', 'Profit %'
        ]
        
        # Style headers
        header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.column_dimensions[get_column_letter(col_num)].width = 18
        
        # Add sample data
        sample_data = [
            ['PP Compound', 'Grade A', 'RM001', 'kg', 150.50, '', 0.025, 0.005, 
             0, 0, 0, 2, 5, 3, 10],
            ['ABS Plastic', 'Grade B', 'RM002', 'kg', 200.00, 195.00, 0.030, 0.006,
             5, 2, 1, 2.5, 4.5, 2.5, 12]
        ]
        
        for row_num, row_data in enumerate(sample_data, 2):
            for col_num, value in enumerate(row_data, 1):
                ws.cell(row=row_num, column=col_num, value=value)
        
        return wb
    
    @staticmethod
    def create_moulding_machines_template():
        """Create template for moulding machines"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Moulding Machines"
        
        headers = [
            'Cavity*', 'Machine Tonnage*', 'Cycle Time (s)*', 'Efficiency %*',
            'Shift Rate*', 'Shift Rate for MTC*', 'MTC Count*',
            'Rejection %', 'Overhead %', 'Maintenance %', 'Profit %'
        ]
        
        header_fill = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.column_dimensions[get_column_letter(col_num)].width = 18
        
        sample_data = [
            [2, 150, 45, 85, 5000, 4500, 500, 3, 5, 4, 15],
            [4, 250, 60, 90, 7000, 6500, 750, 2.5, 4.5, 3.5, 12]
        ]
        
        for row_num, row_data in enumerate(sample_data, 2):
            for col_num, value in enumerate(row_data, 1):
                ws.cell(row=row_num, column=col_num, value=value)
        
        return wb
    
    @staticmethod
    def create_assemblies_template():
        """Create template for assemblies"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Assemblies"
        
        headers = [
            'Assembly Type*', 'Assembly RM Cost', 'Manufacturing Cost',
            'Profit Cost', 'Rejection Cost', 'Inspection Cost'
        ]
        
        header_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.column_dimensions[get_column_letter(col_num)].width = 20
        
        sample_data = [
            ['Manual Assembly', 10.5, 25.0, 5.0, 2.5, 3.0],
            ['Automated Assembly', 15.0, 40.0, 8.0, 3.0, 4.5]
        ]
        
        for row_num, row_data in enumerate(sample_data, 2):
            for col_num, value in enumerate(row_data, 1):
                ws.cell(row=row_num, column=col_num, value=value)
        
        return wb
    
    @staticmethod
    def create_packaging_template():
        """Create template for packaging"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Packaging"
        
        headers = [
            'Packaging Type*', 'Lifecycle*', 'Cost*', 'Parts per Package*',
            'Polybag Cost'
        ]
        
        header_fill = PatternFill(start_color="E26B0A", end_color="E26B0A", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.column_dimensions[get_column_letter(col_num)].width = 20
        
        sample_data = [
            ['Cardboard Box', 100, 5.50, 50, 0.25],
            ['Plastic Container', 200, 12.00, 100, 0.30]
        ]
        
        for row_num, row_data in enumerate(sample_data, 2):
            for col_num, value in enumerate(row_data, 1):
                ws.cell(row=row_num, column=col_num, value=value)
        
        return wb
    
    @staticmethod
    def create_transport_template():
        """Create template for transport"""
        wb = Workbook()
        ws = wb.active
        ws.title = "Transport"
        
        headers = [
            'Total Boxes*', 'Parts per Box*', 'Trip Cost*'
        ]
        
        header_fill = PatternFill(start_color="9933FF", end_color="9933FF", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            ws.column_dimensions[get_column_letter(col_num)].width = 20
        
        sample_data = [
            [50, 100, 5000],
            [100, 200, 8000]
        ]
        
        for row_num, row_data in enumerate(sample_data, 2):
            for col_num, value in enumerate(row_data, 1):
                ws.cell(row=row_num, column=col_num, value=value)
        
        return wb
    
    @staticmethod
    def create_complete_quote_template():
        """Create complete template with all sheets"""
        wb = Workbook()
        wb.remove(wb.active)  # Remove default sheet
        
        # Add Raw Materials sheet
        ws_rm = wb.create_sheet("Raw Materials")
        rm_headers = [
            'Material Name*', 'Grade', 'RM Code*', 'Unit (kg/gm/ton)*',
            'RM Rate*', 'Frozen Rate', 'Part Weight*', 'Runner Weight*',
            'Process Losses', 'Purging Loss Cost', 'ICC %',
            'Rejection %', 'Overhead %', 'Maintenance %', 'Profit %'
        ]
        ExcelTemplateGenerator._add_headers(ws_rm, rm_headers, "4472C4")
        
        # Add Moulding Machines sheet
        ws_mm = wb.create_sheet("Moulding Machines")
        mm_headers = [
            'Cavity*', 'Machine Tonnage*', 'Cycle Time (s)*', 'Efficiency %*',
            'Shift Rate*', 'Shift Rate for MTC*', 'MTC Count*',
            'Rejection %', 'Overhead %', 'Maintenance %', 'Profit %'
        ]
        ExcelTemplateGenerator._add_headers(ws_mm, mm_headers, "70AD47")
        
        # Add Assemblies sheet
        ws_asm = wb.create_sheet("Assemblies")
        asm_headers = [
            'Assembly Type*', 'Assembly RM Cost', 'Manufacturing Cost',
            'Profit Cost', 'Rejection Cost', 'Inspection Cost'
        ]
        ExcelTemplateGenerator._add_headers(ws_asm, asm_headers, "FFC000")
        
        # Add Packaging sheet
        ws_pkg = wb.create_sheet("Packaging")
        pkg_headers = [
            'Packaging Type*', 'Lifecycle*', 'Cost*', 'Parts per Package*',
            'Polybag Cost'
        ]
        ExcelTemplateGenerator._add_headers(ws_pkg, pkg_headers, "E26B0A")
        
        # Add Transport sheet
        ws_trans = wb.create_sheet("Transport")
        trans_headers = [
            'Total Boxes*', 'Parts per Box*', 'Trip Cost*'
        ]
        ExcelTemplateGenerator._add_headers(ws_trans, trans_headers, "9933FF")
        
        return wb
    
    @staticmethod
    def _add_headers(worksheet, headers, color):
        """Helper to add headers to worksheet"""
        header_fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        
        for col_num, header in enumerate(headers, 1):
            cell = worksheet.cell(row=1, column=col_num)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            worksheet.column_dimensions[get_column_letter(col_num)].width = 18


class ExcelParser:
    """Parse uploaded Excel files"""
    
    @staticmethod
    def parse_raw_materials(file, quote):
        """Parse raw materials from Excel"""
        from core.models import RawMaterial
        
        wb = openpyxl.load_workbook(file)
        ws = wb.active
        
        raw_materials = []
        errors = []
        
        for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not any(row):  # Skip empty rows
                continue
            
            try:
                rm = RawMaterial(
                    quote=quote,
                    material_name=row[0] or '',
                    grade=row[1] or '',
                    rm_code=row[2] or '',
                    unit_of_measurement=row[3] or 'kg',
                    rm_rate=float(row[4]) if row[4] else 0,
                    frozen_rate=float(row[5]) if row[5] else None,
                    part_weight=float(row[6]) if row[6] else 0,
                    runner_weight=float(row[7]) if row[7] else 0,
                    process_losses=float(row[8]) if row[8] else 0,
                    purging_loss_cost=float(row[9]) if row[9] else 0,
                    icc_percentage=float(row[10]) if row[10] else 0,
                    rejection_percentage=float(row[11]) if row[11] else 0,
                    overhead_percentage=float(row[12]) if row[12] else 0,
                    maintenance_percentage=float(row[13]) if row[13] else 0,
                    profit_percentage=float(row[14]) if row[14] else 0,
                )
                raw_materials.append(rm)
            except Exception as e:
                errors.append(f"Row {row_num}: {str(e)}")
        
        if not errors:
            RawMaterial.objects.bulk_create(raw_materials)
        
        return len(raw_materials), errors
    
    @staticmethod
    def parse_moulding_machines(file, quote):
        """Parse moulding machines from Excel"""
        from core.models import MouldingMachineDetail
        
        wb = openpyxl.load_workbook(file)
        ws = wb.active
        
        machines = []
        errors = []
        
        for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if not any(row):
                continue
            
            try:
                mm = MouldingMachineDetail(
                    quote=quote,
                    cavity=int(row[0]) if row[0] else 1,
                    machine_tonnage=float(row[1]) if row[1] else 0,
                    cycle_time=float(row[2]) if row[2] else 0,
                    efficiency=float(row[3]) if row[3] else 0,
                    shift_rate=float(row[4]) if row[4] else 0,
                    shift_rate_for_mtc=float(row[5]) if row[5] else 0,
                    mtc_count=int(row[6]) if row[6] else 0,
                    rejection_percentage=float(row[7]) if row[7] else 0,
                    overhead_percentage=float(row[8]) if row[8] else 0,
                    maintenance_percentage=float(row[9]) if row[9] else 0,
                    profit_percentage=float(row[10]) if row[10] else 0,
                )
                machines.append(mm)
            except Exception as e:
                errors.append(f"Row {row_num}: {str(e)}")
        
        if not errors:
            MouldingMachineDetail.objects.bulk_create(machines)
        
        return len(machines), errors
    
    @staticmethod
    def parse_complete_quote(file, quote):
        """Parse complete quote from multi-sheet Excel"""
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
        
        # Add similar parsing for other sheets...
        
        return results
    
    @staticmethod
    def _parse_raw_materials_sheet(worksheet, quote):
        """Helper to parse raw materials sheet"""
        from core.models import RawMaterial
        
        raw_materials = []
        errors = []
        
        for row_num, row in enumerate(worksheet.iter_rows(min_row=2, values_only=True), start=2):
            if not any(row):
                continue
            
            try:
                rm = RawMaterial(
                    quote=quote,
                    material_name=row[0] or '',
                    grade=row[1] or '',
                    rm_code=row[2] or '',
                    unit_of_measurement=row[3] or 'kg',
                    rm_rate=float(row[4]) if row[4] else 0,
                    frozen_rate=float(row[5]) if row[5] else None,
                    part_weight=float(row[6]) if row[6] else 0,
                    runner_weight=float(row[7]) if row[7] else 0,
                    process_losses=float(row[8]) if row[8] else 0,
                    purging_loss_cost=float(row[9]) if row[9] else 0,
                    icc_percentage=float(row[10]) if row[10] else 0,
                    rejection_percentage=float(row[11]) if row[11] else 0,
                    overhead_percentage=float(row[12]) if row[12] else 0,
                    maintenance_percentage=float(row[13]) if row[13] else 0,
                    profit_percentage=float(row[14]) if row[14] else 0,
                )
                raw_materials.append(rm)
            except Exception as e:
                errors.append(f"Row {row_num}: {str(e)}")
        
        if not errors:
            RawMaterial.objects.bulk_create(raw_materials)
        
        return len(raw_materials), errors
    
    @staticmethod
    def _parse_moulding_machines_sheet(worksheet, quote):
        """Helper to parse moulding machines sheet"""
        from core.models import MouldingMachineDetail
        
        machines = []
        errors = []
        
        for row_num, row in enumerate(worksheet.iter_rows(min_row=2, values_only=True), start=2):
            if not any(row):
                continue
            
            try:
                mm = MouldingMachineDetail(
                    quote=quote,
                    cavity=int(row[0]) if row[0] else 1,
                    machine_tonnage=float(row[1]) if row[1] else 0,
                    cycle_time=float(row[2]) if row[2] else 0,
                    efficiency=float(row[3]) if row[3] else 0,
                    shift_rate=float(row[4]) if row[4] else 0,
                    shift_rate_for_mtc=float(row[5]) if row[5] else 0,
                    mtc_cost=float(row[6]) if row[6] else 0,
                    rejection_percentage=float(row[7]) if row[7] else 0,
                    overhead_percentage=float(row[8]) if row[8] else 0,
                    maintenance_percentage=float(row[9]) if row[9] else 0,
                    profit_percentage=float(row[10]) if row[10] else 0,
                )
                machines.append(mm)
            except Exception as e:
                errors.append(f"Row {row_num}: {str(e)}")
        
        if not errors:
            MouldingMachineDetail.objects.bulk_create(machines)
        
        return len(machines), errors