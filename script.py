import pandas as pd
import sqlite3
import json
import os
import re
from datetime import datetime, date
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import coordinate_from_string

# ============================================================================
# Configuration
# ============================================================================
excel_files = ["Deposits Data Lite.xlsx", "Form X Report  Main Lite.xlsx", "Loans Data Lite.xlsx"]
report_sheets = {"Part I", "Part II", "Part III", "MIS-Report"}
exclude_sheets = {"Pivot-Borrowings"}
db_filename = "excel_data.db"
output_dir = "output"
new_base_path = ""  # Will be set at runtime

# Output directory for recreated files
# (Directory creation moved to main function)

# ============================================================================
# Helper Functions for Formula Path Handling
# ============================================================================
def fix_external_references(formula, excel_file_map):
    """
    Fix external references in Excel formulas.
    
    Args:
        formula: The Excel formula string
        excel_file_map: Dictionary mapping original filenames to new filenames
        
    Returns:
        Updated formula with fixed paths
    """
    if not formula or not isinstance(formula, str):
        return formula
    
    # Log the formula for debugging
    if "xlsx" in formula or "[" in formula and "]" in formula:
        print(f"Analyzing formula with potential external reference: {formula}")
    
    # Pattern to match standard external references like: 'Path\[Filename.xlsx]SheetName'!Range
    standard_pattern = r"'?([^']*\[([^]]+)\]([^!']*))'?!([A-Z0-9:$]+)"
    
    # Pattern to match references with numeric workbook indices like: [1]Deposits!L:L
    indexed_pattern = r"\[(\d+)\]([^!]+)!([A-Z0-9:$]+)"
    
    # Pattern to match non-conventional references that might include sheet names or file paths
    sheet_reference_pattern = r"([\w\s-]+\.xlsx)([\w\s-]+)"
    
    def replace_standard_match(match):
        full_path = match.group(1)  # The whole path with filename and sheet
        filename = match.group(2)   # Just the filename
        sheet = match.group(3)      # Just the sheet name
        cell_ref = match.group(4)   # The cell reference
        
        print(f"Found standard external reference - File: {filename}, Sheet: {sheet}, Cell: {cell_ref}")
        
        # Check if this filename matches any of our known files (possibly with slight differences)
        target_filename = None
        for file_key, file_value in excel_file_map.items():
            # Case-insensitive comparison, ignore spaces and parentheses
            clean_filename = filename.lower().replace(" ", "").replace("(", "").replace(")", "")
            clean_key = file_key.lower().replace(" ", "").replace("(", "").replace(")", "")
            if clean_filename in clean_key or clean_key in clean_filename:
                target_filename = file_value
                break
        
        if target_filename:
            # Construct the new reference with the new path and matched filename
            new_ref = f"'{new_base_path}[{target_filename}]{sheet}'!{cell_ref}"
            print(f"Updated to: {new_ref}")
            return new_ref
        else:
            # If no match found, keep the original reference
            return match.group(0)
    
    def replace_indexed_match(match):
        """Handle references with numeric workbook indices like [1]Deposits!L:L"""
        workbook_index = match.group(1)  # The numeric index (e.g., '1')
        sheet_name = match.group(2)      # The sheet name (e.g., 'Deposits')
        cell_ref = match.group(3)        # The range reference (e.g., 'L:L')
        
        print(f"Found indexed external reference - Index: [{workbook_index}], Sheet: {sheet_name}, Cell: {cell_ref}")
        
        # Map workbook indices to file names - customize based on your specific workbooks
        index_to_file = {
            '1': 'Deposits Data Lite.xlsx',
            '2': 'Loans Data Lite.xlsx',
            '3': 'Form X Report  Main Lite.xlsx'
        }
        
        # Preserve the indexed reference format - Excel will resolve these correctly
        # if it can find the referenced workbooks
        if workbook_index in index_to_file and new_base_path:
            # For new base path, create a full path reference
            filename = index_to_file[workbook_index]
            new_ref = f"'{new_base_path}[{filename}]{sheet_name}'!{cell_ref}"
            print(f"Mapped indexed reference to: {new_ref}")
            return new_ref
        else:
            # Keep the original indexed reference if no base path change or unknown index
            print(f"Preserving original indexed reference: [{workbook_index}]{sheet_name}!{cell_ref}")
            return match.group(0)
    
    def replace_sheet_reference(match):
        file = match.group(1)
        sheet = match.group(3)
        print(f"Found non-standard reference - File: {file}, Content: {sheet}")
        
        # Try to match with our known files
        target_filename = None
        for file_key, file_value in excel_file_map.items():
            clean_file = file.lower().replace(" ", "").replace("(", "").replace(")", "")
            clean_key = file_key.lower().replace(" ", "").replace("(", "").replace(")", "")
            if clean_file in clean_key or clean_key in clean_file:
                target_filename = file_value
                break
        
        if target_filename and new_base_path:
            return f"{new_base_path}{target_filename}{sheet}"
        return match.group(0)
    
    # Apply different patterns in sequence
    updated_formula = re.sub(standard_pattern, replace_standard_match, formula)
    updated_formula = re.sub(indexed_pattern, replace_indexed_match, updated_formula)
    updated_formula = re.sub(sheet_reference_pattern, replace_sheet_reference, updated_formula)
    
    # Debug output when we change formulas
    if updated_formula != formula:
        print(f"Formula updated:\nFrom: {formula}\nTo:   {updated_formula}")
        
    return updated_formula

# ============================================================================
# Phase 1 Helper Functions: Data Identification
# ============================================================================
def extract_cell_formatting_phase1(cell):
    """Extract formatting details from a cell for storage in JSON."""
    return {
        "number_format": cell.number_format,
        "font": {
            "name": cell.font.name,
            "size": cell.font.size,
            "bold": cell.font.bold,
            "italic": cell.font.italic,
            "underline": cell.font.underline,
            "color": str(cell.font.color.rgb) if cell.font.color and cell.font.color.rgb else None
        },
        "fill": {
            "fill_type": cell.fill.fill_type,
            "start_color": str(cell.fill.start_color.rgb) if cell.fill.start_color and cell.fill.start_color.rgb else None,
            "end_color": str(cell.fill.end_color.rgb) if cell.fill.end_color and cell.fill.end_color.rgb else None
        },
        "alignment": {
            "horizontal": cell.alignment.horizontal,
            "vertical": cell.alignment.vertical,
            "wrap_text": cell.alignment.wrap_text
        },
        "border": {
            "top": {
                "style": cell.border.top.style if cell.border.top else None,
                "color": str(cell.border.top.color.rgb) if cell.border.top and cell.border.top.color and cell.border.top.color.rgb else None
            },
            "bottom": {
                "style": cell.border.bottom.style if cell.border.bottom else None,
                "color": str(cell.border.bottom.color.rgb) if cell.border.bottom and cell.border.bottom.color and cell.border.bottom.color.rgb else None
            },
            "left": {
                "style": cell.border.left.style if cell.border.left else None,
                "color": str(cell.border.left.color.rgb) if cell.border.left and cell.border.left.color and cell.border.left.color.rgb else None
            },
            "right": {
                "style": cell.border.right.style if cell.border.right else None,
                "color": str(cell.border.right.color.rgb) if cell.border.right and cell.border.right.color and cell.border.right.color.rgb else None
            }
        }
    }

# Custom JSON encoder to handle datetime objects
class DateTimeEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, (datetime, date)):
            return obj.isoformat()
        return super(DateTimeEncoder, self).default(obj)

# ============================================================================
# Phase 2 Helper Functions: Data Storage
# ============================================================================
def setup_database(db_path):
    """Create database and necessary tables."""
    # Create a new database connection
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    print(f"Creating database schema in: {db_path}")
    
    # Create workbooks table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS workbooks (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        filename TEXT UNIQUE,
        properties TEXT
    )
    """)
    
    # Create sheets table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS sheets (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        workbook_id INTEGER,
        sheet_name TEXT,
        sheet_type TEXT,
        max_row INTEGER,
        max_column INTEGER,
        merged_cells TEXT,
        column_dimensions TEXT,
        row_dimensions TEXT,
        FOREIGN KEY (workbook_id) REFERENCES workbooks (id),
        UNIQUE (workbook_id, sheet_name)
    )
    """)
    
    # Create cells table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS cells (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        sheet_id INTEGER,
        coordinate TEXT,
        value TEXT,
        is_formula BOOLEAN,
        formatting TEXT,
        FOREIGN KEY (sheet_id) REFERENCES sheets (id),
        UNIQUE (sheet_id, coordinate)
    )
    """)
    
    # Create tabular_data table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS tabular_data (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        workbook TEXT,
        sheet TEXT,
        table_name TEXT UNIQUE,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    """)
    
    conn.commit()
    return conn

def extract_cell_formatting_phase2(cell):
    """Extract formatting details from a cell for database storage with enhanced color handling."""
    # Simplified color handling - avoid direct string conversion
    font_color = None
    try:
        if cell.font.color:
            if hasattr(cell.font.color, 'rgb') and cell.font.color.rgb:
                # For RGB colors, use a normalized format
                if isinstance(cell.font.color.rgb, str):
                    font_color = {'type': 'rgb', 'value': cell.font.color.rgb}
                else:
                    font_color = {'type': 'rgb', 'value': 'FF000000'}  # Default to black if not a string
            elif hasattr(cell.font.color, 'theme') and cell.font.color.theme is not None:
                # For theme colors
                font_color = {'type': 'theme', 'value': cell.font.color.theme}
            elif hasattr(cell.font.color, 'indexed') and cell.font.color.indexed is not None:
                # For indexed colors
                font_color = {'type': 'indexed', 'value': cell.font.color.indexed}
    except Exception as e:
        print(f"Error extracting font color: {e}")
        font_color = {'type': 'rgb', 'value': 'FF000000'}  # Default to black on error
    
    # Extract fill color with similar approach
    fill_color = None
    try:
        if cell.fill and cell.fill.start_color:
            if hasattr(cell.fill.start_color, 'rgb') and cell.fill.start_color.rgb:
                if isinstance(cell.fill.start_color.rgb, str):
                    fill_color = {'type': 'rgb', 'value': cell.fill.start_color.rgb}
                else:
                    fill_color = {'type': 'rgb', 'value': 'FFFFFFFF'}  # Default to white
            elif hasattr(cell.fill.start_color, 'theme') and cell.fill.start_color.theme is not None:
                fill_color = {'type': 'theme', 'value': cell.fill.start_color.theme}
            elif hasattr(cell.fill.start_color, 'indexed') and cell.fill.start_color.indexed is not None:
                fill_color = {'type': 'indexed', 'value': cell.fill.start_color.indexed}
    except Exception as e:
        print(f"Error extracting fill color: {e}")
        fill_color = None
    
    # Build formatting dictionary
    fmt = {
        "number_format": cell.number_format,
        "font": {
            "name": cell.font.name,
            "size": cell.font.size,
            "bold": cell.font.bold,
            "italic": cell.font.italic,
            "underline": cell.font.underline,
            "color": font_color
        },
        "fill": {
            "fill_type": cell.fill.fill_type,
            "start_color": fill_color,
            "end_color": None  # Simplified to avoid similar issues with end color
        },
        "alignment": {
            "horizontal": cell.alignment.horizontal,
            "vertical": cell.alignment.vertical,
            "wrap_text": cell.alignment.wrap_text
        },
        "border": {
            "top": {
                "style": cell.border.top.style if cell.border.top else None,
                "color": None  # Simplified border color handling
            },
            "bottom": {
                "style": cell.border.bottom.style if cell.border.bottom else None,
                "color": None
            },
            "left": {
                "style": cell.border.left.style if cell.border.left else None,
                "color": None
            },
            "right": {
                "style": cell.border.right.style if cell.border.right else None,
                "color": None
            }
        }
    }
    return fmt

# ============================================================================
# Phase 3 Helper Functions: Data Recreation
# ============================================================================
def apply_formatting(cell, formatting_json):
    """Apply formatting details from JSON to the cell."""
    try:
        # Parse the formatting JSON
        fmt = json.loads(formatting_json)
        
        # Apply number format
        if "number_format" in fmt:
            cell.number_format = fmt["number_format"]
        
        # Apply font formatting
        if "font" in fmt:
            font_data = fmt["font"]
            
            # Handle font color with specific approach for each color type
            font_color = None
            if "color" in font_data and font_data["color"]:
                color_info = font_data["color"]
                if color_info["type"] == "rgb":
                    # Use the RGB value directly, defaulting to black if issues
                    if color_info["value"] and isinstance(color_info["value"], str):
                        # Clean up the value to ensure proper format
                        font_color = color_info["value"].strip()
                        # Default to black for any problematic values
                        if not font_color or not all(c in '0123456789ABCDEFabcdef' for c in font_color.replace('FF', '', 1)):
                            font_color = '000000'
                    else:
                        font_color = '000000'  # Default to black
                elif color_info["type"] == "theme":
                    # Theme colors aren't directly supported in this simple approach
                    # Default to black for theme colors
                    font_color = '000000'
                elif color_info["type"] == "indexed":
                    # Indexed colors aren't directly supported in this simple approach
                    # Default to black for indexed colors
                    font_color = '000000'
            
            # Create and apply the font
            try:
                cell.font = Font(
                    name=font_data.get("name"),
                    size=font_data.get("size"),
                    bold=font_data.get("bold"),
                    italic=font_data.get("italic"),
                    underline=font_data.get("underline"),
                    color=font_color
                )
            except Exception as e:
                print(f"Error setting font for cell {cell.coordinate}: {e}")
                # Apply font without color if there was an error
                cell.font = Font(
                    name=font_data.get("name"),
                    size=font_data.get("size"),
                    bold=font_data.get("bold"),
                    italic=font_data.get("italic"),
                    underline=font_data.get("underline")
                )
        
        # Apply fill formatting
        if "fill" in fmt:
            fill_data = fmt["fill"]
            fill_type = fill_data.get("fill_type")
            
            # Apply fill only if it's a valid fill type
            if fill_type and fill_type not in ['none', None]:
                # Get fill color
                fill_color = None
                if "start_color" in fill_data and fill_data["start_color"]:
                    color_info = fill_data["start_color"]
                    if color_info["type"] == "rgb" and color_info["value"]:
                        fill_color = color_info["value"]
                
                # Apply the fill if we have a valid color
                if fill_color:
                    try:
                        cell.fill = PatternFill(
                            fill_type=fill_type,
                            start_color=fill_color,
                            end_color=fill_color  # Use same color for end
                        )
                    except Exception as e:
                        print(f"Error applying fill for cell {cell.coordinate}: {e}")
        
        # Apply alignment formatting
        if "alignment" in fmt:
            align_data = fmt["alignment"]
            cell.alignment = Alignment(
                horizontal=align_data.get("horizontal"),
                vertical=align_data.get("vertical"),
                wrap_text=align_data.get("wrap_text")
            )
        
        # Apply border formatting (simplified - just style without color)
        if "border" in fmt:
            border_data = fmt["border"]
            
            # Helper function to create a Side object
            def create_side(side_data):
                if not side_data or not side_data.get("style"):
                    return None
                return Side(style=side_data["style"])
            
            # Create border sides
            top = create_side(border_data.get("top"))
            bottom = create_side(border_data.get("bottom"))
            left = create_side(border_data.get("left"))
            right = create_side(border_data.get("right"))
            
            # Set the border if any sides are defined
            if any([top, bottom, left, right]):
                cell.border = Border(top=top or Side(), bottom=bottom or Side(),
                                    left=left or Side(), right=right or Side())
    
    except Exception as e:
        print(f"Error applying formatting to cell {cell.coordinate}: {e}")

def apply_merged_cells(worksheet, merged_cells_json):
    """Apply merged cell ranges to the worksheet."""
    try:
        merged_ranges = json.loads(merged_cells_json)
        for merged_range in merged_ranges:
            worksheet.merge_cells(merged_range)
    except Exception as e:
        print(f"Error applying merged cells: {e}")

def apply_dimensions(worksheet, column_dimensions_json, row_dimensions_json):
    """Apply column and row dimensions to the worksheet."""
    try:
        column_dimensions = json.loads(column_dimensions_json)
        for col, properties in column_dimensions.items():
            if col in worksheet.column_dimensions and properties.get("width"):
                worksheet.column_dimensions[col].width = properties["width"]
    except Exception as e:
        print(f"Error applying column dimensions: {e}")
    
    try:
        row_dimensions = json.loads(row_dimensions_json)
        for row_str, properties in row_dimensions.items():
            row = int(row_str)
            if row in worksheet.row_dimensions and properties.get("height"):
                worksheet.row_dimensions[row].height = properties["height"]
    except Exception as e:
        print(f"Error applying row dimensions: {e}")

# ============================================================================
# Phase 4 Helper Functions: Font Color Correction
# ============================================================================
def fix_font_colors(excel_file):
    """
    Open the Excel file, set all font colors to black, and save as a fixed version.
    """
    print(f"Processing file: {excel_file}")
    
    # Load the workbook
    wb = load_workbook(excel_file)
    
    # Process each sheet
    for sheet_name in wb.sheetnames:
        print(f"  Processing sheet: {sheet_name}")
        ws = wb[sheet_name]
        
        # Get the sheet dimensions
        max_row = ws.max_row
        max_col = ws.max_column
        
        # Process all cells
        cells_modified = 0
        for row in range(1, max_row + 1):
            for col in range(1, max_col + 1):
                cell = ws.cell(row=row, column=col)
                
                # Skip empty cells
                if cell.value is None:
                    continue
                
                # Get current font properties
                current_font = cell.font
                
                # Create a new font with the same properties but black color
                new_font = Font(
                    name=current_font.name,
                    size=current_font.size,
                    bold=current_font.bold,
                    italic=current_font.italic,
                    underline=current_font.underline,
                    color='000000'  # Set to black
                )
                
                # Apply the new font
                cell.font = new_font
                cells_modified += 1
        
        print(f"    Modified {cells_modified} cells")
    
    # Create output filename
    base_name, ext = os.path.splitext(excel_file)
    output_file = f"{base_name}_fixed{ext}"
    
    # Save the modified workbook
    wb.save(output_file)
    print(f"Saved fixed file: {output_file}")
    return output_file

# ============================================================================
# Phase 1: Data Identification
# ============================================================================
def identify_data():
    """Extract and identify data from Excel files."""
    print("\n" + "="*70)
    print("PHASE 1: DATA IDENTIFICATION")
    print("="*70)
    
    # Access global configuration variables
    global excel_files, report_sheets, exclude_sheets
    
    # Dictionary to store all workbook data and metadata
    workbook_data = {}
    
    # Create a list to store potential references for later analysis
    potential_references = []
    
    # Process each Excel file
    for file in excel_files:
        print(f"Processing file: {file}")
        workbook_data[file] = {"sheets": {}}
        
        # Load the workbook with openpyxl (data_only=False to capture formulas)
        wb = load_workbook(file, data_only=False)
        
        # Get workbook-level properties
        workbook_data[file]["properties"] = {
            "title": wb.properties.title,
            "creator": wb.properties.creator,
            "created": str(wb.properties.created) if wb.properties.created else None,
            "sheet_names": wb.sheetnames
        }
        
        # Process each sheet in the workbook
        for sheet_name in wb.sheetnames:
            # Skip excluded sheets
            if sheet_name in exclude_sheets:
                print(f"Skipping excluded sheet: {sheet_name}")
                continue
            
            # Get the worksheet
            ws = wb[sheet_name]
            
            # Initialize sheet data
            sheet_data = {
                "type": "report" if sheet_name in report_sheets else "non_report",
                "max_row": ws.max_row,
                "max_column": ws.max_column,
                "merged_cells": [str(merged_range) for merged_range in ws.merged_cells.ranges],
                "column_dimensions": {col: {"width": ws.column_dimensions[col].width} 
                                     for col in ws.column_dimensions},
                "row_dimensions": {row: {"height": ws.row_dimensions[row].height} 
                                  for row in ws.row_dimensions},
                "cells": {}
            }
            
            # Process all cells in the sheet
            print(f"  Processing cells in sheet: {sheet_name}")
            for row in range(1, ws.max_row + 1):
                for col in range(1, ws.max_column + 1):
                    cell = ws.cell(row=row, column=col)
                    
                    # Skip empty cells to save space
                    if cell.value is None:
                        continue
                    
                    # Enhanced detection for references and formulas
                    cell_value = cell.value
                    is_formula = False
                    
                    if isinstance(cell_value, str):
                        # Check for standard formula
                        if cell_value.startswith('='):
                            is_formula = True
                        
                        # Check for potential external references
                        if '.xlsx' in cell_value or '.xls' in cell_value or ('[' in cell_value and ']' in cell_value):
                            potential_references.append({
                                'file': file,
                                'sheet': sheet_name,
                                'cell': cell.coordinate,
                                'value': cell_value
                            })
                            print(f"  Potential external reference found in {file}, sheet {sheet_name}, cell {cell.coordinate}: {cell_value}")
                    
                    # Handle datetime objects for JSON serialization
                    if isinstance(cell_value, (datetime, date)):
                        cell_value = cell_value.isoformat()
                    
                    # Store cell data
                    sheet_data["cells"][cell.coordinate] = {
                        "value": cell_value,
                        "is_formula": is_formula,
                        "formatting": extract_cell_formatting_phase1(cell)
                    }
            
            # Store sheet data in the workbook dictionary
            workbook_data[file]["sheets"][sheet_name] = sheet_data
        
        print(f"Completed processing file: {file}")
    
    # Output summary of potential references
    if potential_references:
        print("\nPotential External References Found:")
        for ref in potential_references:
            print(f"  {ref['file']} - Sheet: {ref['sheet']} - Cell: {ref['cell']} - Value: {ref['value']}")
    
    # Output summary
    print("\nData Identification Summary:")
    for file, data in workbook_data.items():
        print(f"\nFile: {file}")
        print(f"  Total sheets: {len(data['sheets'])}")
        for sheet_name, sheet_data in data['sheets'].items():
            cell_count = len(sheet_data['cells'])
            sheet_type = sheet_data['type']
            print(f"  Sheet: {sheet_name} ({sheet_type}) - {cell_count} non-empty cells")
    
    # Save the identification results to a JSON file for reference
    with open('workbook_identification.json', 'w') as f:
        json.dump(workbook_data, f, indent=2, cls=DateTimeEncoder)
    
    print("\nData identification complete. Results saved to workbook_identification.json.")
    return workbook_data

# ============================================================================
# Phase 2: Data Storage
# ============================================================================
def store_data(workbook_data=None):
    """Store identified data in SQLite database."""
    print("\n" + "="*70)
    print("PHASE 2: DATA STORAGE")
    print("="*70)
    
    # Access global configuration variables
    global excel_files, report_sheets, exclude_sheets, db_filename, new_base_path
    
    # Create mapping dictionary for Excel filenames
    excel_file_map = {}
    for file in excel_files:
        base_name = os.path.basename(file)
        # Add variations of the filename (with/without spaces, with/without extension)
        name_without_ext = os.path.splitext(base_name)[0]
        clean_name = base_name.replace(" ", "")
        clean_name_without_ext = name_without_ext.replace(" ", "")
        
        excel_file_map[base_name] = base_name
        excel_file_map[name_without_ext] = base_name
        excel_file_map[clean_name] = base_name
        excel_file_map[clean_name_without_ext] = base_name
    
    # Set up database - ensure we start with a clean database
    if os.path.exists(db_filename):
        try:
            os.remove(db_filename)
            print(f"Removed existing database: {db_filename}")
        except PermissionError:
            print(f"Could not remove existing database. Make sure it's not in use by another program.")
            raise
    
    conn = setup_database(db_filename)
    cursor = conn.cursor()
    
    # Process each Excel file
    for file in excel_files:
        print(f"Processing file: {file}")
        
        # Insert workbook metadata
        wb = load_workbook(file, data_only=False)
        properties = {
            "title": wb.properties.title,
            "creator": wb.properties.creator,
            "created": str(wb.properties.created) if wb.properties.created else None,
            "sheet_names": wb.sheetnames
        }
        
        cursor.execute(
            "INSERT OR REPLACE INTO workbooks (filename, properties) VALUES (?, ?)",
            (file, json.dumps(properties))
        )
        conn.commit()
        
        # Get the workbook_id
        cursor.execute("SELECT id FROM workbooks WHERE filename = ?", (file,))
        workbook_id = cursor.fetchone()[0]
        
        # Process each sheet in the workbook
        for sheet_name in wb.sheetnames:
            if sheet_name in exclude_sheets:
                print(f"  Skipping excluded sheet: {sheet_name}")
                continue
            
            ws = wb[sheet_name]
            sheet_type = "report" if sheet_name in report_sheets else "non_report"
            
            # Store sheet metadata
            sheet_metadata = {
                "merged_cells": [str(merged_range) for merged_range in ws.merged_cells.ranges],
                "column_dimensions": {col: {"width": ws.column_dimensions[col].width} 
                                      for col in ws.column_dimensions},
                "row_dimensions": {row: {"height": ws.row_dimensions[row].height}
                                  for row in ws.row_dimensions}
            }
            
            cursor.execute(
                """INSERT OR REPLACE INTO sheets 
                   (workbook_id, sheet_name, sheet_type, max_row, max_column, merged_cells, column_dimensions, row_dimensions) 
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                (workbook_id, sheet_name, sheet_type, ws.max_row, ws.max_column, 
                 json.dumps(sheet_metadata["merged_cells"]),
                 json.dumps(sheet_metadata["column_dimensions"]),
                 json.dumps(sheet_metadata["row_dimensions"]))
            )
            conn.commit()
            
            # Get the sheet_id
            cursor.execute("SELECT id FROM sheets WHERE workbook_id = ? AND sheet_name = ?", 
                           (workbook_id, sheet_name))
            sheet_id = cursor.fetchone()[0]
            
            # Store cell data and formatting
            print(f"  Processing cells in sheet: {sheet_name}")
            for row in range(1, ws.max_row + 1):
                for col in range(1, ws.max_column + 1):
                    cell = ws.cell(row=row, column=col)
                    
                    # Skip empty cells to save space
                    if cell.value is None:
                        continue
                    
                    # Determine if the cell contains a formula
                    is_formula = isinstance(cell.value, str) and cell.value.startswith('=')
                    
                    # Convert cell value to string for storage
                    if isinstance(cell.value, (int, float, bool)):
                        cell_value = str(cell.value)
                    elif cell.value is None:
                        cell_value = ""
                    else:
                        cell_value = str(cell.value)
                    
                    # Fix external references in formulas if needed
                    if is_formula and new_base_path:
                        cell_value = fix_external_references(cell_value, excel_file_map)
                    
                    # Extract formatting
                    formatting_dict = extract_cell_formatting_phase2(cell)
                    
                    # Store cell data
                    cursor.execute(
                        "INSERT OR REPLACE INTO cells (sheet_id, coordinate, value, is_formula, formatting) VALUES (?, ?, ?, ?, ?)",
                        (sheet_id, cell.coordinate, cell_value, is_formula, json.dumps(formatting_dict))
                    )
            
            # For non-report sheets, also store as tabular data
            if sheet_type == "non_report":
                try:
                    # Generate a unique table name
                    base_name = os.path.splitext(os.path.basename(file))[0]
                    table_name = f"{base_name}_{sheet_name}".replace(" ", "_").replace("-", "_")
                    
                    # Read the sheet as a DataFrame
                    df = pd.read_excel(file, sheet_name=sheet_name)
                    
                    # Store the DataFrame in the database
                    df.to_sql(table_name, conn, if_exists='replace', index=False)
                    
                    # Record this table in the tabular_data table
                    cursor.execute(
                        "INSERT OR REPLACE INTO tabular_data (workbook, sheet, table_name) VALUES (?, ?, ?)",
                        (file, sheet_name, table_name)
                    )
                    
                    print(f"  Stored tabular data for sheet '{sheet_name}' in table '{table_name}'")
                except Exception as e:
                    print(f"  Error storing tabular data for sheet '{sheet_name}': {e}")
            
            conn.commit()
        
        print(f"Completed processing file: {file}")
    
    # Generate summary
    cursor.execute("""
        SELECT w.filename, COUNT(DISTINCT s.id) as sheet_count, 
               SUM(CASE WHEN s.sheet_type = 'report' THEN 1 ELSE 0 END) as report_sheets,
               SUM(CASE WHEN s.sheet_type = 'non_report' THEN 1 ELSE 0 END) as non_report_sheets,
               COUNT(c.id) as cell_count
        FROM workbooks w
        JOIN sheets s ON w.id = s.workbook_id
        LEFT JOIN cells c ON s.id = c.sheet_id
        GROUP BY w.filename
    """)
    summary = cursor.fetchall()
    
    print("\nData Storage Summary:")
    for row in summary:
        filename, sheet_count, report_sheets, non_report_sheets, cell_count = row
        print(f"\nFile: {filename}")
        print(f"  Total sheets: {sheet_count} ({report_sheets} report, {non_report_sheets} non-report)")
        print(f"  Total cells stored: {cell_count}")
    
    # Close the database connection
    conn.commit()
    conn.close()
    print("\nData storage complete.")

# ============================================================================
# Phase 3: Data Recreation
# ============================================================================
def recreate_workbooks():
    """Recreate workbooks from stored data."""
    print("\n" + "="*70)
    print("PHASE 3: DATA RECREATION")
    print("="*70)
    
    # Access global configuration variables
    global excel_files, exclude_sheets, db_filename, output_dir, new_base_path
    
    print("Formula handling settings:")
    print(f"  External reference path: {new_base_path if new_base_path else 'ORIGINAL PATHS'}")
    
    # Create mapping dictionary for Excel filenames
    excel_file_map = {}
    for file in excel_files:
        base_name = os.path.basename(file)
        # Add variations of the filename (with/without spaces, with/without extension)
        name_without_ext = os.path.splitext(base_name)[0]
        clean_name = base_name.replace(" ", "")
        clean_name_without_ext = name_without_ext.replace(" ", "")
        
        excel_file_map[base_name] = base_name
        excel_file_map[name_without_ext] = base_name
        excel_file_map[clean_name] = base_name
        excel_file_map[clean_name_without_ext] = base_name
    
    # Create a mapping for numeric index references
    index_to_filename = {
        '1': 'Deposits Data Lite.xlsx',
        '2': 'Loans Data Lite.xlsx',
        '3': 'Form X Report  Main Lite.xlsx'
    }
    print("Workbook index mapping:")
    for idx, filename in index_to_filename.items():
        print(f"  [{idx}] -> {filename}")
        
    # Open database connection
    conn = sqlite3.connect(db_filename)
    cursor = conn.cursor()
    
    # Get recreated workbook paths for later use in font fixing
    recreated_files = []
    
    # Process each workbook for recreation
    for file in excel_files:
        print(f"Recreating workbook: {file}")
        base_name = os.path.splitext(os.path.basename(file))[0]
        
        # Get workbook ID
        cursor.execute("SELECT id FROM workbooks WHERE filename = ?", (file,))
        workbook_id_result = cursor.fetchone()
        
        if not workbook_id_result:
            print(f"Workbook {file} not found in database. Skipping.")
            continue
        
        workbook_id = workbook_id_result[0]
        
        # Get sheet information for this workbook
        cursor.execute("""
            SELECT id, sheet_name, sheet_type, max_row, max_column, merged_cells, column_dimensions, row_dimensions
            FROM sheets
            WHERE workbook_id = ?
        """, (workbook_id,))
        sheets_info = cursor.fetchall()
        
        # Create a new workbook
        new_wb = Workbook()
        # Remove the default sheet
        if len(sheets_info) > 0:
            default_sheet = new_wb.active
            new_wb.remove(default_sheet)
        
        # Process each sheet
        for sheet_info in sheets_info:
            sheet_id, sheet_name, sheet_type, max_row, max_column, merged_cells, column_dimensions, row_dimensions = sheet_info
            
            print(f"  Recreating sheet: {sheet_name} (type: {sheet_type})")
            
            # Create the sheet
            ws = new_wb.create_sheet(title=sheet_name)
            
            # Apply merged cells
            if merged_cells:
                apply_merged_cells(ws, merged_cells)
            
            # Apply dimensions
            if column_dimensions and row_dimensions:
                apply_dimensions(ws, column_dimensions, row_dimensions)
            
            # Populate sheet data
            cursor.execute("""
                SELECT coordinate, value, is_formula, formatting
                FROM cells
                WHERE sheet_id = ?
            """, (sheet_id,))
            cells_data = cursor.fetchall()
            
            for cell_data in cells_data:
                coordinate, value, is_formula, formatting = cell_data
                
                # Set the cell value
                if is_formula:
                    # For formula cells, apply any path updates if needed
                    if new_base_path:
                        formula_value = fix_external_references(value, excel_file_map)
                    else:
                        formula_value = value
                    
                    # Check if this is a sheet with special formula handling needs
                    is_special_sheet = (sheet_name == "MIS-Report" or sheet_name == "Part I" or 
                                        sheet_name == "Part II" or sheet_name == "Part III")
                    
                    # Check if this is an indexed reference formula
                    has_indexed_ref = ('[1]' in value) or ('[2]' in value) or ('[3]' in value)
                    
                    try:
                        # Special handling for indexed references like [1]Deposits!L:L
                        if is_special_sheet and has_indexed_ref:
                            print(f"  Preserving indexed reference in {coordinate}: {value}")
                            # For indexed references, it's better to set the value directly
                            # Excel will resolve these correctly if workbooks are available
                            ws[coordinate].value = formula_value
                        else:
                            # Standard formula handling via openpyxl's formula property
                            if formula_value.startswith('='):
                                print(f"  Setting formula in {coordinate}: {formula_value}")
                                # Clear the cell first to avoid any existing value interference
                                ws[coordinate].value = None
                                # Set the formula without the equals sign (openpyxl requirement)
                                ws[coordinate].formula = formula_value[1:]
                            else:
                                # If formula doesn't start with equals but is marked as formula
                                print(f"  Setting non-standard formula in {coordinate}: {formula_value}")
                                ws[coordinate].value = None
                                ws[coordinate].formula = formula_value
                    except Exception as e:
                        print(f"  Error setting formula in {coordinate}: {e}")
                        # Fallback - set as text if formula application fails
                        ws[coordinate].value = formula_value
                else:
                    # For non-formula cells, try to convert to appropriate data type
                    if value.lower() == 'true':
                        ws[coordinate] = True
                    elif value.lower() == 'false':
                        ws[coordinate] = False
                    else:
                        try:
                            # Try to convert to number (int or float)
                            if value.isdigit():
                                ws[coordinate] = int(value)
                            else:
                                ws[coordinate] = float(value)
                        except (ValueError, TypeError):
                            # If not a number, keep as string
                            ws[coordinate] = value
                
                # Apply formatting
                if formatting:
                    apply_formatting(ws[coordinate], formatting)
        
        # Save the new workbook
        output_file = os.path.join(output_dir, f"{base_name}_recreated.xlsx")
        
        try:
            # Before saving, add a special linked workbook index list for MIS-Report formulas
            # This helps Excel properly resolve [1], [2], etc. references
            if 'Form X Report' in file:
                print(f"  Adding workbook links for indexed references...")
                
                # If the workbook doesn't have a defined name for external references,
                # we need to create one to help Excel resolve [1], [2] references
                try:
                    # Try adding the workbooks to the existing links
                    # (This is an approximation - full linking requires Excel's COM interface)
                    links_sheet = new_wb.create_sheet(title="_Links", index=0)
                    links_sheet["A1"] = "Workbook Index References"
                    links_sheet["A2"] = "[1] = Deposits Data Lite.xlsx"
                    links_sheet["A3"] = "[2] = Loans Data Lite.xlsx"
                    links_sheet["A4"] = "[3] = Form X Report  Main Lite.xlsx"
                    links_sheet["A6"] = "Note: These links help resolve formulas with [1], [2] references."
                    links_sheet["A7"] = "You may need to update links manually in Excel: Data > Edit Links"
                except Exception as e:
                    print(f"  Could not create links helper sheet: {e}")
            
            new_wb.save(output_file)
            print(f"Created new workbook: {output_file}")
            
            # Add a note about opening the file with Excel
            print(f"NOTE: When opening {output_file}, Excel may prompt to update links.")
            print("      Select 'Update' to refresh external references.")
            
            recreated_files.append(output_file)
        except Exception as e:
            print(f"ERROR saving workbook {output_file}: {str(e)}")
            print("This might be due to formula issues or locked file access.")
    
    # Close the database connection
    conn.close()
    print("Data recreation complete.")
    
    return recreated_files

# ============================================================================
# Phase 4: Font Color Correction
# ============================================================================
def fix_workbook_fonts(recreated_files):
    """Fix font colors in recreated workbooks."""
    print("\n" + "="*70)
    print("PHASE 4: FONT COLOR CORRECTION")
    print("="*70)
    
    # Access global configuration variables
    global output_dir
    
    fixed_files = []
    
    # Process each recreated file
    for file in recreated_files:
        try:
            if os.path.exists(file):
                fixed_file = fix_font_colors(file)
                fixed_files.append(fixed_file)
            else:
                print(f"File not found: {file}")
        except Exception as e:
            print(f"Error processing {file}: {e}")
    
    print("\nFont Color Correction Summary:")
    print(f"Processed {len(fixed_files)} files:")
    for file in fixed_files:
        print(f"  - {file}")
    
    return fixed_files

# ============================================================================
# Main Function
# ============================================================================
def main():
    """Main function to execute the entire Excel processing workflow."""
    global new_base_path
    
    print("\n" + "="*70)
    print("EXCEL PROCESSING SYSTEM")
    print("="*70)
    print(f"Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    # Get the new base path for external references
    new_base_path = input("Enter new base path for external references (leave empty to keep original): ")
    if new_base_path and not new_base_path.endswith('\\'):
        new_base_path += '\\'  # Ensure path ends with backslash
    
    print(f"External reference path will be updated to: '{new_base_path}'" if new_base_path else "External references will keep original paths")
    
    # Make output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Created output directory: {output_dir}")
        
    # Check if the Excel files actually exist
    for file in excel_files:
        if not os.path.exists(file):
            print(f"WARNING: Input file '{file}' does not exist in the current directory.")
            print(f"Current directory: {os.path.abspath('.')}")
            confirm = input("Continue anyway? (y/n): ")
            if confirm.lower() != 'y':
                print("Process aborted.")
                return
    
    try:
        # Phase 1: Data Identification
        workbook_data = identify_data()
        
        # Phase 2: Data Storage
        store_data(workbook_data)
        
        # Phase 3: Data Recreation
        recreated_files = recreate_workbooks()
        
        # Phase 4: Font Color Correction
        fixed_files = fix_workbook_fonts(recreated_files)
        
        # Final Summary
        print("\n" + "="*70)
        print("PROCESS COMPLETED SUCCESSFULLY")
        print("="*70)
        print(f"Input Files: {len(excel_files)}")
        print(f"Recreated Files: {len(recreated_files)}")
        print(f"Fixed Files: {len(fixed_files)}")
        print(f"Output Directory: {os.path.abspath(output_dir)}")
        print(f"Completed at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        print("\nIMPORTANT NOTES FOR USING RECREATED FILES:")
        print("1. When opening workbooks with external references, Excel will prompt to update links")
        print("2. Select 'Update' when prompted to ensure formulas display correct values")
        print("3. If formulas show #VALUE! or appear blank, press F9 to recalculate")
        print("4. For persistent issues, check Data > Edit Links to verify paths") 
        print("5. Ensure all referenced files exist in the correct locations")
    
    except Exception as e:
        print("\n" + "="*70)
        print("ERROR ENCOUNTERED")
        print("="*70)
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        print("\nProcess terminated with errors.")

if __name__ == "__main__":
    main()