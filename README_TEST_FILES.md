# Test Files for Quote Management System

## Generating Test Files

Run the following command from the project root directory:
```bash
python generate_test_files.py
```

This will create a `test_files/` directory with 7 Excel files for testing.

## Generated Test Files

### 1. Component-Level Test Files (for existing quotes)

#### test_raw_materials.xlsx
- **Contains**: 5 raw materials
- **Materials**: PP, ABS, Nylon 6, Polycarbonate, PVC
- **Use**: Upload to add raw materials to an existing quote

#### test_moulding_machines.xlsx
- **Contains**: 4 moulding machines
- **Configurations**: Various cavities (1, 2, 4, 8) and tonnages (80-250)
- **Use**: Upload to add moulding machines to an existing quote

#### test_assemblies.xlsx
- **Contains**: 3 assemblies
- **Types**: Manual, Automated, Semi-Auto
- **Use**: Upload to add assemblies to an existing quote

#### test_packaging.xlsx
- **Contains**: 3 packaging options
- **Types**: PP Box, CG Box, Bin
- **Use**: Upload to add packaging to an existing quote

#### test_transport.xlsx
- **Contains**: 2 transport options
- **Use**: Upload to add transport to an existing quote

### 2. Complete Quote Test Files (for projects)

#### test_complete_quote.xlsx
- **Contains**: 1 complete quote
- **Quote Name**: Test Complete Quote
- **Components**: 2 raw materials, 1 machine, 1 assembly, 1 packaging, 1 transport
- **Use**: Upload to create a complete quote with all components in a single upload

#### test_multiple_quotes.xlsx
- **Contains**: 3 complete quotes with all components
- **Quotes**:
  1. Auto Dashboard 2024 (3 materials, 2 machines)
  2. Electronics Housing Pro (2 materials, 1 machine)
  3. Medical Device Shell (2 materials, 1 machine)
- **Use**: Upload to create multiple complete quotes in a project

## Testing Workflow

### Test 1: Component Upload to Existing Quote
1. Create a project manually
2. Create a quote manually with basic definition
3. Navigate to the quote detail page
4. Upload `test_raw_materials.xlsx` → should add 5 raw materials
5. Upload `test_moulding_machines.xlsx` → should add 4 machines
6. Upload `test_assemblies.xlsx` → should add 3 assemblies
7. Upload `test_packaging.xlsx` → should add 3 packaging options
8. Upload `test_transport.xlsx` → should add 2 transport options
9. Check quote summary for all components

### Test 2: Complete Quote Upload
1. Create a project manually
2. Select customer group
3. Navigate to project detail
4. Click "Bulk Upload Quotes"
5. Upload `test_complete_quote.xlsx`
6. Should create 1 quote with all components
7. Verify all sections are populated

### Test 3: Multiple Quotes Upload
1. Create a project manually
2. Select customer group
3. Navigate to project detail
4. Click "Bulk Upload Quotes"
5. Upload `test_multiple_quotes.xlsx`
6. Should create 3 quotes:
   - Auto Dashboard 2024 (high volume)
   - Electronics Housing Pro (medium volume)
   - Medical Device Shell (low volume, high value)
7. Verify each quote has all components
8. Check quote summaries for correct calculations

## Data Characteristics

### Raw Materials
- Realistic material types (PP, ABS, Nylon, PC, PVC)
- Varied weights (0.02 - 0.07 kg)
- Different cost percentages (rejection: 1.5-3%, overhead: 3.5-6%, profit: 8-15%)

### Moulding Machines
- Multiple cavity configurations (1, 2, 4, 8)
- Realistic cycle times (28-58 seconds)
- High efficiency rates (85-92%)
- Varied shift rates (3200-9500)

### Assemblies
- Mix of manual and automated processes
- Realistic costs (8-45 per unit)
- Different complexity levels

### Complete Quotes
- **Auto Dashboard**: High volume (10,000 units), multiple materials
- **Electronics**: Medium volume (25,000 units), technical specs
- **Medical**: Low volume (5,000 units), high quality standards

## Prerequisites

Before uploading test files:
1. Create at least one Customer Group in Configuration
2. Create at least one Project
3. Ensure you're logged in as a user with appropriate permissions

## Expected Results

After uploading test files, you should see:
- Version increments in quote timeline
- All components visible in respective tabs
- Correct cost calculations in quote summary
- No errors in upload process
- Success messages showing item counts

## Troubleshooting

If uploads fail:
- Verify customer group is selected
- Check that quote exists (for component uploads)
- Ensure quote is in "in_progress" status
- Check error messages for specific issues
- Verify Excel file format is .xlsx

## Notes

- All test data uses realistic industry values
- Decimal precision is up to 8 places as configured
- Quote names in multiple quotes file are unique
- Material codes follow a consistent naming pattern
- All required fields are populated