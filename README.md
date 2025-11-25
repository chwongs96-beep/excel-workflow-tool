# Excel Workflow Tool

A simple, offline n8n-like visual workflow automation tool focused on Excel processing. No API keys, no database - just drag-and-drop Excel workflow automation.

## Features

- **Visual Node Editor**: Drag and drop nodes to create workflows
- **Excel Focus**: Built specifically for Excel/spreadsheet operations
- **Offline**: Works completely offline, no internet required
- **Save/Load**: Save your workflows and reload them later
- **Data Preview**: See your data at each step

## Available Nodes

### 📥 Input/Output
| Node | Description |
|------|-------------|
| **Read Excel** | Load data from Excel files (.xlsx, .xls) |
| **Write Excel** | Save data to Excel files |
| **Data Preview** | View data in the preview panel |

### 📑 Sheet Operations
| Node | Description |
|------|-------------|
| **Read All Sheets** | Read all sheets from an Excel file |
| **Get Sheet** | Extract a specific sheet from sheets dictionary |
| **List Sheets** | List all sheet names in an Excel file |
| **Copy Sheet to File** | Copy data to a specific sheet in another file |
| **Copy Sheet Between Files** | Copy sheet directly between two Excel files |
| **Write Multi-Sheet Excel** | Write multiple datasets to different sheets |
| **Merge Sheets** | Combine multiple sheets into one dataset |

### 🔄 Transform
| Node | Description |
|------|-------------|
| **Filter Rows** | Filter data based on conditions |
| **Select Columns** | Keep only specific columns |
| **Rename Columns** | Rename column headers |
| **Sort Data** | Sort by one or more columns |
| **Remove Duplicates** | Remove duplicate rows |
| **Add Column** | Add new columns with values or formulas |
| **Sample Data** | Take random or systematic samples |
| **Format Numbers** | Format numbers (currency, percentage, etc.) |
| **Add Rank** | Add ranking column based on values |

### 🧹 Clean
| Node | Description |
|------|-------------|
| **Trim Whitespace** | Remove leading/trailing spaces |
| **Remove Empty Rows** | Remove rows with null/empty values |
| **Find & Replace** | Find and replace values |
| **Change Data Type** | Convert column data types |
| **Fill Missing Values** | Handle empty/null values |

### ✏️ Text Processing
| Node | Description |
|------|-------------|
| **Change Text Case** | UPPER, lower, Title Case |
| **Split Column** | Split one column into multiple |
| **Combine Columns** | Merge multiple columns into one |
| **Extract Text** | Extract text using regex or position |

### 📅 Date/Time
| Node | Description |
|------|-------------|
| **Extract Date Parts** | Extract year, month, day, etc. |
| **Format Date** | Convert date to specific format |
| **Date Difference** | Calculate days/months/years between dates |

### 🔀 Logic
| Node | Description |
|------|-------------|
| **IF Condition** | Create column based on IF/ELSE |
| **Multiple Conditions** | CASE/SWITCH style conditions |

### 📊 Aggregate
| Node | Description |
|------|-------------|
| **Group By** | Group and aggregate data |
| **Pivot Table** | Create pivot table summary |
| **Statistics Summary** | Generate statistical summary |

### 🔗 Combine
| Node | Description |
|------|-------------|
| **Merge Data** | Join datasets (like SQL JOIN) |
| **VLOOKUP** | Look up values from another table |
| **Concatenate Data** | Stack datasets vertically |

### ✅ Validate
| Node | Description |
|------|-------------|
| **Find Duplicates** | Find and report duplicate rows |
| **Data Validation** | Validate data with rules (email, range, etc.) |

## Installation

1. Make sure you have Python 3.9+ installed
2. Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

Run the application:

```bash
python main.py
```


### Creating a Workflow

1. **Add Nodes**: Double-click nodes in the left panel to add them to the canvas
2. **Connect Nodes**: Click and drag from an output port (pink) to an input port (green)
3. **Configure Nodes**: Select a node to see its configuration in the right panel
4. **Execute**: Click the "Execute" button to run the workflow

### Controls

- **Pan**: Middle mouse button drag
- **Zoom**: Mouse wheel
- **Delete Node**: Select and press Delete key
- **Move Node**: Left click and drag

### Example Workflow

1. Add a "Read Excel" node and configure the file path
2. Add a "Filter Rows" node and connect it
3. Add a "Select Columns" node to keep only needed columns
4. Add a "Write Excel" node to save the result
5. Click Execute!

## File Format

Workflows are saved as `.workflow.json` files which can be shared and reloaded.

## Requirements

- Python 3.9+
- PyQt6
- pandas
- openpyxl (for Excel support)
