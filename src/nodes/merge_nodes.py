import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from copy import copy as copy_obj
from pathlib import Path
from typing import Any, Dict, List, Union
from .base_node import BaseNode
from .node_registry import register_node

import shutil

class StyledSheet:
    """Wrapper to hold sheet info for style-preserved copying"""
    def __init__(self, file_path, sheet_name, df_filtered=None, header_row=0, is_full_copy=False):
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.df_filtered = df_filtered
        self.header_row = header_row
        self.is_full_copy = is_full_copy

# ============================================================================
# 批量合并节点 (Batch Merge)
# ============================================================================

@register_node
class MergeExcelFilesNode(BaseNode):
    """Node to merge multiple Excel files into one"""
    
    node_type = "merge_excel_files"
    node_name = "批量合并Excel"
    node_category = "批量处理"
    node_description = "将多个Excel文件合并到一个文件中"
    node_color = "#8b5cf6"  # Violet
    
    def _setup_ports(self):
        self.add_output("file_path")
    
    def get_config_ui_schema(self) -> List[Dict[str, Any]]:
        return [
            {
                "key": "base_file",
                "label": "基础文件 (File 1)",
                "type": "file",
                "file_filter": "Excel文件 (*.xlsx *.xls)",
                "required": True
            },
            {
                "key": "files_to_merge",
                "label": "要合并的文件 (File 2, 3, 4...)",
                "type": "file_multiple",
                "file_filter": "Excel文件 (*.xlsx *.xls)",
                "required": True
            },
            {
                "key": "sheet_mode",
                "label": "工作表选择模式",
                "type": "select",
                "options": [
                    {"value": "all", "label": "所有工作表"},
                    {"value": "first", "label": "仅第一个工作表"},
                    {"value": "name", "label": "指定工作表名称"}
                ],
                "default": "all"
            },
            {
                "key": "sheet_name",
                "label": "指定工作表名称 (如果选择)",
                "type": "text",
                "default": "",
                "placeholder": "例如: Sheet1"
            },
            {
                "key": "output_file",
                "label": "输出文件路径",
                "type": "file_save",
                "file_filter": "Excel文件 (*.xlsx)",
                "required": True
            }
        ]
    
    def validate(self) -> tuple[bool, str]:
        base_file = self.get_param("base_file", "")
        if not base_file or not Path(base_file).exists():
            return False, "基础文件是必需的"
            
        files = self.get_param("files_to_merge", "")
        if not files:
            return False, "要合并的文件是必需的"
            
        output_file = self.get_param("output_file", "")
        if not output_file:
            return False, "输出文件路径是必需的"
            
        return True, ""
    
    def execute(self, input_data: Dict[str, Any]) -> Dict[str, Any]:
        base_file = self.get_param("base_file")
        files_str = self.get_param("files_to_merge", "")
        files_to_merge = [f.strip() for f in files_str.split('\n') if f.strip()]
        output_file = self.get_param("output_file")
        
        sheet_mode = self.get_param("sheet_mode", "all")
        target_sheet_name = self.get_param("sheet_name", "")
        
        # Helper to read sheets based on mode
        def read_sheets(file_path):
            if sheet_mode == "all":
                return pd.read_excel(file_path, sheet_name=None)
            elif sheet_mode == "first":
                df = pd.read_excel(file_path, sheet_name=0)
                return {"Sheet1": df} # Use generic name, will be renamed
            elif sheet_mode == "name":
                if not target_sheet_name:
                    # Fallback to all if name not specified
                    return pd.read_excel(file_path, sheet_name=None)
                try:
                    df = pd.read_excel(file_path, sheet_name=target_sheet_name)
                    return {target_sheet_name: df}
                except Exception:
                    print(f"Warning: Sheet '{target_sheet_name}' not found in {file_path}")
                    return {}
            return {}

        # Read base file sheets
        try:
            # Base file always reads all sheets usually, or should it follow the rule?
            # Let's assume base file is the "template" so we keep all its sheets usually.
            # But if user wants to merge specific sheets from ALL files including base...
            # Let's keep base file intact (all sheets) as it is the "Base".
            base_dfs = pd.read_excel(base_file, sheet_name=None)
        except Exception as e:
            raise ValueError(f"读取基础文件失败: {e}")
            
        merged_sheets = {}
        
        # Add base sheets first
        for sheet_name, df in base_dfs.items():
            merged_sheets[sheet_name] = df
            
        # Process other files
        for i, file_path in enumerate(files_to_merge):
            try:
                dfs = read_sheets(file_path)
                
                for sheet_name, df in dfs.items():
                    # Create a unique sheet name
                    # If mode is 'first', we might want to name it after the file
                    if sheet_mode == "first":
                        p = Path(file_path)
                        new_sheet_name = p.stem
                    else:
                        new_sheet_name = sheet_name
                        
                    # Conflict resolution
                    if new_sheet_name in merged_sheets:
                        p = Path(file_path)
                        new_sheet_name = f"{p.stem}_{sheet_name}"
                        
                    if new_sheet_name in merged_sheets:
                        counter = 1
                        while f"{new_sheet_name}_{counter}" in merged_sheets:
                            counter += 1
                        new_sheet_name = f"{new_sheet_name}_{counter}"
                    
                    merged_sheets[new_sheet_name] = df
                    
            except Exception as e:
                print(f"Warning: Failed to read {file_path}: {e}")
        
        # Write to output file
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name, df in merged_sheets.items():
                # Excel sheet name limit is 31 chars
                safe_name = sheet_name[:31]
                df.to_excel(writer, sheet_name=safe_name, index=False)
                
        return {"file_path": output_file}


# ============================================================================
# 灵活工作流节点 (Flexible Workflow)
# ============================================================================

@register_node
class WorkbookCreateNode(BaseNode):
    """Node to start a workbook workflow"""
    
    node_type = "workbook_create"
    node_name = "创建工作簿(输入)"
    node_category = "灵活合并"
    node_description = "开始一个新的工作簿，可选从现有文件加载"
    node_color = "#22c55e"  # Green
    
    def _setup_ports(self):
        self.add_output("workbook")
    
    def get_config_ui_schema(self) -> List[Dict[str, Any]]:
        return [
            {
                "key": "base_file",
                "label": "基础文件 (可选)",
                "type": "file",
                "file_filter": "Excel文件 (*.xlsx *.xls)",
                "placeholder": "留空则创建一个空工作簿"
            }
        ]
    
    def execute(self, input_data: Dict[str, Any]) -> Dict[str, Any]:
        base_file = self.get_param("base_file", "")
        workbook = {}
        
        if base_file and Path(base_file).exists():
            try:
                # Read all sheets as DataFrames
                dfs = pd.read_excel(base_file, sheet_name=None)
                
                # Wrap them in StyledSheet to preserve original styles
                for sheet_name, df in dfs.items():
                    workbook[sheet_name] = StyledSheet(base_file, sheet_name, df, is_full_copy=True)
                    
            except Exception as e:
                raise ValueError(f"读取基础文件失败: {e}")
        
        return {"workbook": workbook}


@register_node
class WorkbookAppendNode(BaseNode):
    """Node to append a sheet from another file"""
    
    node_type = "workbook_append"
    node_name = "追加工作表"
    node_category = "灵活合并"
    node_description = "从另一个Excel文件读取工作表并添加到当前工作簿"
    node_color = "#8b5cf6"  # Violet
    
    def _setup_ports(self):
        self.add_input("workbook")
        self.add_input("file_path")  # Optional input for dynamic source file
        self.add_output("workbook")
    
    def get_config_ui_schema(self) -> List[Dict[str, Any]]:
        return [
            {
                "key": "source_type",
                "label": "来源类型",
                "type": "select",
                "options": [
                    {"value": "file", "label": "指定文件"},
                    {"value": "search", "label": "搜索文件夹"}
                ],
                "default": "file"
            },
            {
                "key": "file_path",
                "label": "文件路径 (指定文件时)",
                "type": "file",
                "file_filter": "Excel/CSV文件 (*.xlsx *.xls *.csv)",
                "placeholder": "请选择要追加的Excel或CSV文件 (或连接输入节点)"
            },
            {
                "key": "folder_path",
                "label": "文件夹路径 (搜索时)",
                "type": "directory",
                "placeholder": "请选择要搜索的文件夹"
            },
            {
                "key": "keyword",
                "label": "文件名关键字 (搜索时)",
                "type": "text",
                "placeholder": "例如: 2023年报 (留空则匹配所有文件)"
            },
            {
                "key": "sheet_mode",
                "label": "选择模式",
                "type": "select",
                "options": [
                    {"value": "first", "label": "第一个工作表"},
                    {"value": "name", "label": "指定名称"},
                    {"value": "all", "label": "所有工作表"}
                ],
                "default": "first"
            },
            {
                "key": "sheet_name",
                "label": "源工作表名称 (指定名称时)",
                "type": "sheet_selector",
                "dependency": "file_path",
                "default": ""
            },
            {
                "key": "target_name",
                "label": "目标工作表名称 (可选)",
                "type": "text",
                "default": "",
                "placeholder": "留空则自动命名 (文件名/原名)"
            }
        ]
    
    def validate(self) -> tuple[bool, str]:
        source_type = self.get_param("source_type", "file")
        # Relax validation to allow dynamic input via connection
        if source_type == "search":
            if not self.get_param("folder_path"):
                return False, "搜索文件夹是必需的"
        return True, ""
    
    def execute(self, input_data: Dict[str, Any]) -> Dict[str, Any]:
        workbook = input_data.get("workbook")
        if workbook is None:
            workbook = {} # Start fresh if no input
        else:
            workbook = workbook.copy() # Shallow copy dict
            
        # Determine file path
        source_type = self.get_param("source_type", "file")
        file_path = ""
        
        # Check input port first for dynamic file path
        if "file_path" in input_data and input_data["file_path"]:
            file_path = input_data["file_path"]
        elif source_type == "file":
            file_path = self.get_param("file_path")
        else:
            folder_path = self.get_param("folder_path")
            keyword = self.get_param("keyword", "")
            
            if not folder or not Path(folder).exists():
                raise ValueError(f"文件夹不存在: {folder}")
                
            p = Path(folder)
            # Find excel/csv files
            files = list(p.glob("*.xlsx")) + list(p.glob("*.xls")) + list(p.glob("*.csv"))
            
            # Filter by keyword
            if keyword:
                files = [f for f in files if keyword in f.name]
            
            if not files:
                raise ValueError(f"在 {folder} 中未找到匹配 '{keyword}' 的Excel/CSV文件")
            
            # Sort by name and take first
            files.sort(key=lambda f: f.name)
            file_path = str(files[0])
            print(f"Found file by keyword '{keyword}': {file_path}")

        sheet_mode = self.get_param("sheet_mode", "first")
        src_sheet_name = self.get_param("sheet_name", "")
        target_name = self.get_param("target_name", "")
        
        try:
            is_csv = str(file_path).lower().endswith('.csv')
            
            if is_csv:
                # CSV handling
                try:
                    # Try reading with default encoding first
                    df = pd.read_csv(file_path)
                except UnicodeDecodeError:
                    try:
                        # Try GBK (common for Chinese CSVs)
                        df = pd.read_csv(file_path, encoding='gbk')
                    except UnicodeDecodeError:
                        # Try UTF-8-SIG (Excel CSV)
                        df = pd.read_csv(file_path, encoding='utf-8-sig')
                
                default_name = Path(file_path).stem
                
                # Determine target name
                if target_name:
                    t_name = target_name
                else:
                    t_name = default_name
                
                # Ensure unique
                base_t_name = t_name
                counter = 1
                while t_name in workbook:
                    t_name = f"{base_t_name}_{counter}"
                    counter += 1
                
                workbook[t_name] = df
                
            elif sheet_mode == "all":
                dfs = pd.read_excel(file_path, sheet_name=None)
                for name, df in dfs.items():
                    # Determine target name
                    if target_name:
                        # If target name provided for ALL sheets, we must append index or something
                        # But usually target_name is for single sheet.
                        # Let's just use original name + conflict resolution
                        t_name = name
                    else:
                        t_name = name
                        
                    # Conflict resolution
                    if t_name in workbook:
                        p = Path(file_path)
                        t_name = f"{p.stem}_{name}"
                    
                    # Ensure unique
                    base_t_name = t_name
                    counter = 1
                    while t_name in workbook:
                        t_name = f"{base_t_name}_{counter}"
                        counter += 1
                        
                    workbook[t_name] = df
                    
            else:
                # Single sheet
                if sheet_mode == "first":
                    df = pd.read_excel(file_path, sheet_name=0)
                    # Try to get actual name? pd.read_excel(sheet_name=0) returns dataframe, not dict
                    # To get name, we need to load workbook object or read_excel(sheet_name=None) and take first key
                    # Optimization: just read first sheet
                    default_name = Path(file_path).stem # Use filename as default sheet name
                else: # name
                    if not src_sheet_name:
                        raise ValueError("未指定源工作表名称")
                    df = pd.read_excel(file_path, sheet_name=src_sheet_name)
                    default_name = src_sheet_name
                
                # Determine target name
                if target_name:
                    t_name = target_name
                else:
                    t_name = default_name
                
                # Ensure unique
                base_t_name = t_name
                counter = 1
                while t_name in workbook:
                    t_name = f"{base_t_name}_{counter}"
                    counter += 1
                
                workbook[t_name] = df
                
        except Exception as e:
            raise ValueError(f"读取文件失败 {file_path}: {e}")
            
        return {"workbook": workbook}


@register_node
class SheetCopyNode(BaseNode):
    """Node to copy data between sheets with advanced modes"""
    
    node_type = "sheet_copy"
    node_name = "复制/合并数据"
    node_category = "灵活合并"
    node_description = "将Excel/CSV数据复制到工作簿，支持列映射和空值检查"
    node_color = "#f59e0b"  # Amber
    
    def _setup_ports(self):
        self.add_input("workbook")
        self.add_input("file_path")  # Optional input for dynamic source file
        self.add_output("workbook")
    
    def get_config_ui_schema(self) -> List[Dict[str, Any]]:
        return [
            {
                "key": "file_path",
                "label": "来源文件",
                "type": "file",
                "file_filter": "Excel/CSV文件 (*.xlsx *.xls *.csv)",
                "required": True,
                "placeholder": "请选择来源文件 (或连接输入节点)"
            },
            {
                "key": "sheet_name",
                "label": "来源工作表 (Excel)",
                "type": "sheet_selector",
                "dependency": "file_path",
                "default": "",
                "placeholder": "CSV文件可忽略此项"
            },
            {
                "key": "header_row",
                "label": "标题所在行 (从0开始)",
                "type": "number",
                "default": 0,
                "min": 0
            },
            {
                "key": "target_sheet",
                "label": "目标工作表名称",
                "type": "sheet_selector",
                "dependency": "__upstream__",
                "required": True,
                "placeholder": "选择或输入目标工作表名称"
            },
            {
                "key": "copy_mode",
                "label": "复制模式",
                "type": "select",
                "options": [
                    {"value": "whole", "label": "整页复制"},
                    {"value": "columns", "label": "指定列到列"},
                    {"value": "no_blank", "label": "自动检查值(无空白)"}
                ],
                "default": "whole"
            },
            {
                "key": "column_mapping",
                "label": "列映射 (仅指定列模式)",
                "type": "text",
                "placeholder": "格式: 源列A=目标列B; 源列C=目标列D"
            },
            {
                "key": "filter_query",
                "label": "行过滤条件 (Pandas Query)",
                "type": "text",
                "placeholder": "例如: 状态 == '完成' and 金额 > 1000"
            },
            {
                "key": "remove_duplicates",
                "label": "去除重复行",
                "type": "checkbox",
                "default": False
            },
            {
                "key": "strip_whitespace",
                "label": "去除文本首尾空格",
                "type": "checkbox",
                "default": True
            },
            {
                "key": "preserve_formatting",
                "label": "保留原始Excel格式 (较慢)",
                "type": "checkbox",
                "default": True
            },
            {
                "key": "write_mode",
                "label": "写入方式",
                "type": "select",
                "options": [
                    {"value": "overwrite", "label": "覆盖目标表"},
                    {"value": "append", "label": "追加到末尾"}
                ],
                "default": "overwrite"
            }
        ]
    
    def validate(self) -> tuple[bool, str]:
        # Relax validation to allow dynamic input via connection
        # if not self.get_param("file_path"):
        #     return False, "来源文件是必需的"
        if not self.get_param("target_sheet"):
            return False, "目标工作表名称是必需的"
        
        mode = self.get_param("copy_mode")
        if mode == "columns" and not self.get_param("column_mapping"):
            return False, "指定列模式下需要填写列映射"
            
        return True, ""
    
    def execute(self, input_data: Dict[str, Any]) -> Dict[str, Any]:
        workbook = input_data.get("workbook")
        if workbook is None:
            workbook = {}
        else:
            workbook = workbook.copy()
            
        # Check input port first for dynamic file path
        if "file_path" in input_data and input_data["file_path"]:
            file_path = input_data["file_path"]
        else:
            file_path = self.get_param("file_path")
            
        src_sheet_name = self.get_param("sheet_name")
        target_sheet = self.get_param("target_sheet")
        copy_mode = self.get_param("copy_mode", "whole")
        write_mode = self.get_param("write_mode", "overwrite")
        col_mapping_str = self.get_param("column_mapping", "")
        
        header_row = self.get_param("header_row", 0)
        filter_query = self.get_param("filter_query", "")
        remove_duplicates = self.get_param("remove_duplicates", False)
        strip_whitespace = self.get_param("strip_whitespace", True)
        preserve_formatting = self.get_param("preserve_formatting", True)
        
        # 1. Read Source Data
        try:
            is_csv = str(file_path).lower().endswith('.csv')
            if is_csv:
                try:
                    df = pd.read_csv(file_path, header=header_row)
                except UnicodeDecodeError:
                    try:
                        df = pd.read_csv(file_path, encoding='gbk', header=header_row)
                    except UnicodeDecodeError:
                        df = pd.read_csv(file_path, encoding='utf-8-sig', header=header_row)
            else:
                # Excel
                if not src_sheet_name:
                    # Default to first sheet if not specified
                    df = pd.read_excel(file_path, sheet_name=0, header=header_row)
                else:
                    df = pd.read_excel(file_path, sheet_name=src_sheet_name, header=header_row)
        except Exception as e:
            raise ValueError(f"读取来源文件失败: {e}")
            
        # Check row count limit
        if len(df) > 20000:
            print(f"Skipping file {file_path} because it has {len(df)} rows (limit 20000)")
            return {"workbook": workbook}

        try:
            # 2. Pre-process Data (Cleaning & Filtering)
            
            # Strip whitespace from string columns
            if strip_whitespace:
            try:
                df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)
                # Also strip column names if they are strings
                df.columns = df.columns.map(lambda x: x.strip() if isinstance(x, str) else x)
            except Exception as e:
                raise ValueError(f"去除空格失败: {e}")
            
        # Filter rows
        if filter_query:
            try:
                # Support simple syntax like: 状态 == '完成'
                # Pandas query uses the dataframe's columns
                df = df.query(filter_query)
            except Exception as e:
                raise ValueError(f"行过滤条件错误: {e}")
        
        # Remove duplicates
        if remove_duplicates:
            try:
                df = df.drop_duplicates()
            except Exception as e:
                raise ValueError(f"去除重复行失败: {e}")
                df = df.dropna(how='all')
                # Also remove rows where all elements are empty strings (if any)
                # df = df[df.astype(bool).any(axis=1)] # This might be too aggressive for 0 values
                
            elif copy_mode == "columns":
                # Parse mapping: "A=B; C=D" or "Name=Name"
                # If just "A,B,C", assume same names? Let's support "Src=Tgt"
                mappings = [m.strip() for m in col_mapping_str.split(';') if m.strip()]
                
                new_df = pd.DataFrame()
                
                for m in mappings:
                    if '=' in m:
                        src_col, tgt_col = m.split('=', 1)
                        src_col = src_col.strip()
                        tgt_col = tgt_col.strip()
                    else:
                        # If no =, assume src = tgt
                        src_col = m.strip()
                        tgt_col = m.strip()
                    
                    # Check if src_col exists (by name)
                    if src_col in df.columns:
                        new_df[tgt_col] = df[src_col]
                    else:
                        # Try by index if integer?
                        # For now, assume column names. 
                        # If user inputs "A", "B", pandas uses names. 
                        # If excel has no header, columns are 0, 1, 2...
                        # Let's try to handle integer indices if src_col is digit
                        if src_col.isdigit() and int(src_col) < len(df.columns):
                            new_df[tgt_col] = df.iloc[:, int(src_col)]
                        else:
                            print(f"Warning: Column '{src_col}' not found in source.")
                
                df = new_df

            # 4. Write to Target
            if preserve_formatting and not is_csv:
                # Use StyledSheet wrapper
                # We pass the processed dataframe so we know what data to write (cleaned/filtered)
                # But we also pass file path to read styles from
                
                # If appending, we need to handle that in WorkbookSaveNode or here?
                # WorkbookSaveNode handles the final write.
                # If we have multiple StyledSheets for the same target sheet (append), 
                # we might need a list of them.
                
                if target_sheet in workbook:
                    existing = workbook[target_sheet]
                    if isinstance(existing, list):
                        existing.append(StyledSheet(file_path, src_sheet_name, df, header_row))
                    else:
                        # Convert to list if appending
                        if write_mode == "append":
                            workbook[target_sheet] = [existing, StyledSheet(file_path, src_sheet_name, df, header_row)]
                        else:
                            # Overwrite
                            workbook[target_sheet] = StyledSheet(file_path, src_sheet_name, df, header_row)
                else:
                    workbook[target_sheet] = StyledSheet(file_path, src_sheet_name, df, header_row)
                    
            else:
                # Standard DataFrame mode
                if target_sheet in workbook and write_mode == "append":
                    target_data = workbook[target_sheet]
                    # If target is StyledSheet, we can't easily append DataFrame to it without breaking style logic
                    # For now, if mixing, convert everything to DataFrame (lose styles)
                    if isinstance(target_data, StyledSheet):
                        target_data = target_data.df_filtered
                    elif isinstance(target_data, list) and len(target_data) > 0 and isinstance(target_data[0], StyledSheet):
                        # Concatenate all StyledSheets DFs
                        dfs = [s.df_filtered for s in target_data]
                        target_data = pd.concat(dfs, ignore_index=True)
                    
                    if isinstance(target_data, pd.DataFrame):
                        workbook[target_sheet] = pd.concat([target_data, df], ignore_index=True)
                    elif isinstance(target_data, list):
                        # Append to list of StyledSheets? No, this branch is for DataFrame
                        # If we are here, preserve_formatting is False.
                        # So we should probably convert everything to DF.
                        pass
                else:
                    # Overwrite or Create new
                    workbook[target_sheet] = df
        
        except Exception as e:
            raise ValueError(f"处理文件失败 [{file_path}]: {e}")
            
        return {"workbook": workbook}


@register_node
class WorkbookSaveNode(BaseNode):
    """Node to save the workbook"""
    
    node_type = "workbook_save"
    node_name = "保存工作簿(输出)"
    node_category = "灵活合并"
    node_description = "将工作簿保存到文件"
    node_color = "#ef4444"  # Red
    
    def _setup_ports(self):
        self.add_input("workbook")
        self.add_output("file_path")
    
    def get_config_ui_schema(self) -> List[Dict[str, Any]]:
        return [
            {
                "key": "output_file",
                "label": "保存路径",
                "type": "file_save",
                "file_filter": "Excel文件 (*.xlsx)",
                "required": True
            }
        ]
    
    def validate(self) -> tuple[bool, str]:
        if not self.get_param("output_file"):
            return False, "保存路径是必需的"
        return True, ""
    
    def execute(self, input_data: Dict[str, Any]) -> Dict[str, Any]:
        workbook = input_data.get("workbook")
        if workbook is None:
            raise ValueError("没有接收到工作簿数据")
            
        output_file = self.get_param("output_file")
        
        # Ensure file extension is .xlsx
        if not str(output_file).lower().endswith('.xlsx'):
            output_file = str(output_file) + '.xlsx'
        
        if not workbook:
            workbook = {"Sheet1": pd.DataFrame()}
            
        # Check if we have any StyledSheets
        has_styles = False
        for val in workbook.values():
            if isinstance(val, StyledSheet) or (isinstance(val, list) and len(val) > 0 and isinstance(val[0], StyledSheet)):
                has_styles = True
                break
        
        if has_styles:
            self._save_with_styles(output_file, workbook)
        else:
            self._save_standard(output_file, workbook)
                
        return {"file_path": output_file}

    def _save_standard(self, output_file, workbook):
        """Save using standard Pandas (no style preservation)"""
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name, df in workbook.items():
                if isinstance(df, pd.DataFrame):
                    safe_name = str(sheet_name)[:31]
                    df.to_excel(writer, sheet_name=safe_name, index=False)

    def _save_with_styles(self, output_file, workbook):
        """Save using OpenPyXL to preserve styles"""
        
        # Optimization: Check if we can use a template file (Base File)
        # This avoids slow cell-by-cell copying for untouched sheets
        template_file = self._find_template_file(workbook)
        
        if template_file:
            try:
                # Copy template to output
                shutil.copy2(template_file, output_file)
                wb = openpyxl.load_workbook(output_file)
                
                # Track which sheets we have handled (updated or verified as existing)
                handled_sheets = set()
                
                # Update/Add sheets
                for sheet_name, data in workbook.items():
                    handled_sheets.add(sheet_name)
                    
                    # Check if this is the original sheet from template (unmodified)
                    is_original = False
                    if isinstance(data, StyledSheet):
                        if (data.file_path == template_file and 
                            data.sheet_name == sheet_name and 
                            data.is_full_copy):
                            is_original = True
                    
                    if is_original:
                        # It's already in the file, skip writing
                        continue
                    
                    # If we are here, we need to write this sheet.
                    # If it exists in template (but we are overwriting it), remove it first.
                    if sheet_name in wb.sheetnames:
                        # Remove existing sheet to overwrite
                        wb.remove(wb[sheet_name])
                    
                    # Create new sheet
                    safe_name = str(sheet_name)[:31]
                    target_ws = wb.create_sheet(title=safe_name)
                    
                    # Write data
                    self._write_items_to_sheet(data, target_ws)
                
                # Remove sheets that are in template but not in workbook (deleted)
                for sheet in list(wb.sheetnames):
                    if sheet not in handled_sheets:
                        wb.remove(wb[sheet])
                
                wb.save(output_file)
                return
                
            except Exception as e:
                print(f"Template optimization failed, falling back to slow copy: {e}")
                # Fallback to standard creation if optimization fails
        
        # Create new workbook (Fallback or if no template)
        wb = openpyxl.Workbook()
        # Remove default sheet
        if "Sheet" in wb.sheetnames:
            wb.remove(wb["Sheet"])
            
        for sheet_name, data in workbook.items():
            safe_name = str(sheet_name)[:31]
            target_ws = wb.create_sheet(title=safe_name)
            self._write_items_to_sheet(data, target_ws)
        
        wb.save(output_file)

    def _find_template_file(self, workbook):
        """Find a potential template file (Base File) from workbook data"""
        # We look for the most common file path among StyledSheets that are full copies
        # Or just the first one if it covers most sheets.
        # Simple heuristic: If there is at least one StyledSheet with is_full_copy=True,
        # use its file path.
        
        for data in workbook.values():
            if isinstance(data, StyledSheet) and data.is_full_copy:
                if Path(data.file_path).exists():
                    return data.file_path
            elif isinstance(data, list) and len(data) > 0:
                if isinstance(data[0], StyledSheet) and data[0].is_full_copy:
                    if Path(data[0].file_path).exists():
                        return data[0].file_path
        return None

    def _write_items_to_sheet(self, data, target_ws):
        """Write data items (StyledSheet or DataFrame) to target worksheet"""
        items = data if isinstance(data, list) else [data]
        current_row = 1
        
        for item in items:
            if isinstance(item, StyledSheet):
                current_row = self._copy_styled_sheet(item, target_ws, current_row)
            elif isinstance(item, pd.DataFrame):
                # Write dataframe values (no styles)
                # Optimization: Use append() for much faster writing
                
                # Write header if it's the first item
                if current_row == 1:
                    target_ws.append(list(item.columns))
                    current_row += 1
                
                # Convert to list of lists for faster iteration
                # Using itertuples is faster than iterrows
                try:
                    for row_idx, row in enumerate(item.itertuples(index=False), 1):
                        try:
                            target_ws.append(list(row))
                            current_row += 1
                        except Exception as e:
                            raise ValueError(f"写入数据失败，位置: 第 {row_idx} 行. 错误: {e}")
                except Exception as e:
                    if "写入数据失败" in str(e):
                        raise e
                    raise ValueError(f"处理数据失败: {e}")

    def _copy_styled_sheet(self, styled: StyledSheet, target_ws, start_row):
        """Copy data and styles from StyledSheet to target worksheet"""
        try:
            src_wb = openpyxl.load_workbook(styled.file_path, data_only=False)
            if styled.sheet_name and styled.sheet_name in src_wb.sheetnames:
                src_ws = src_wb[styled.sheet_name]
            else:
                src_ws = src_wb.active
                
            df = styled.df_filtered
            header_row_idx = styled.header_row + 1 # 1-based
            
            # Check if data is sequential (unfiltered)
            # If index is RangeIndex(0, N, 1), it means no rows were dropped/reordered
            is_sequential = isinstance(df.index, pd.RangeIndex) and df.index.step == 1 and df.index.start == 0
            
            # 1. Copy Header
            if start_row == 1:
                # Copy header row from source
                for col in range(1, src_ws.max_column + 1):
                    src_cell = src_ws.cell(row=header_row_idx, column=col)
                    tgt_cell = target_ws.cell(row=start_row, column=col)
                    
                    tgt_cell.value = src_cell.value
                    if src_cell.has_style:
                        tgt_cell.font = copy_obj(src_cell.font)
                        tgt_cell.border = copy_obj(src_cell.border)
                        tgt_cell.fill = copy_obj(src_cell.fill)
                        tgt_cell.number_format = copy_obj(src_cell.number_format)
                        tgt_cell.protection = copy_obj(src_cell.protection)
                        tgt_cell.alignment = copy_obj(src_cell.alignment)
                
                # Copy column dimensions
                for col_letter, dim in src_ws.column_dimensions.items():
                    target_ws.column_dimensions[col_letter] = copy_obj(dim)
                
                # Copy merged cells in header area
                for range_ in src_ws.merged_cells.ranges:
                    if range_.max_row <= header_row_idx:
                        # Shift to target header row (start_row)
                        # Source header is at header_row_idx
                        # Target header is at start_row
                        offset = start_row - header_row_idx
                        
                        # We need to shift the range
                        min_row = range_.min_row + offset
                        max_row = range_.max_row + offset
                        target_ws.merge_cells(start_row=min_row, start_column=range_.min_col,
                                            end_row=max_row, end_column=range_.max_col)
                                            
                start_row += 1
            
            # 2. Copy Data Rows
            for idx, row_data in df.iterrows():
                try:
                    if isinstance(idx, int):
                        src_row_idx = header_row_idx + 1 + idx
                    else:
                        src_row_idx = None
                    
                    for col_pos, (col_name, value) in enumerate(row_data.items()):
                        tgt_col_idx = col_pos + 1
                        
                        # Try to find source cell for style
                        src_cell = None
                        if src_row_idx:
                            # Assuming 1:1 column mapping for simplicity in style copying
                            # If columns were reordered, this might pick wrong style source column
                            # But usually acceptable for "Whole" copy mode
                            if tgt_col_idx <= src_ws.max_column:
                                src_cell = src_ws.cell(row=src_row_idx, column=tgt_col_idx)
                        
                        tgt_cell = target_ws.cell(row=start_row, column=tgt_col_idx)
                        tgt_cell.value = value
                        
                        if src_cell and src_cell.has_style:
                            tgt_cell.font = copy_obj(src_cell.font)
                            tgt_cell.border = copy_obj(src_cell.border)
                            tgt_cell.fill = copy_obj(src_cell.fill)
                            tgt_cell.number_format = copy_obj(src_cell.number_format)
                            tgt_cell.protection = copy_obj(src_cell.protection)
                            tgt_cell.alignment = copy_obj(src_cell.alignment)
                    
                    # Copy row dimensions
                    if src_row_idx and src_row_idx in src_ws.row_dimensions:
                        target_ws.row_dimensions[start_row] = copy_obj(src_ws.row_dimensions[src_row_idx])
                    
                    start_row += 1
                except Exception as e:
                    raise ValueError(f"复制带格式数据失败，位置: 第 {idx} 行 (源文件行号: {src_row_idx if src_row_idx else '未知'}). 错误: {e}")
            
            # 3. Copy Merged Cells in Data Area (Only if sequential/unfiltered)
            if is_sequential:
                # Calculate offset between source data start and target data start
                # Source data starts at: header_row_idx + 1
                # Target data started at: (original start_row passed to func) + 1 (if header copied)
                # Current start_row is at the end.
                
                # Let's recalculate target data start row
                # If we copied header, target data started at (original_start_row + 1)
                # If we didn't, target data started at original_start_row
                
                # We can just iterate ranges and check if they are in data area
                for range_ in src_ws.merged_cells.ranges:
                    if range_.min_row > header_row_idx:
                        # This is a data area merge
                        
                        # Calculate offset
                        # Source row R maps to Target row R'
                        # R' = R - (header_row_idx + 1) + target_data_start_row
                        
                        # Wait, simpler:
                        # We know we wrote `len(df)` rows.
                        # The data block in source is `header_row_idx + 1` to `header_row_idx + len(df)`
                        # The data block in target is `target_data_start` to `target_data_start + len(df) - 1`
                        
                        # Let's find target_data_start
                        # It is `start_row - len(df)` (since start_row was incremented in loop)
                        target_data_start = start_row - len(df)
                        
                        offset = target_data_start - (header_row_idx + 1)
                        
                        min_row = range_.min_row + offset
                        max_row = range_.max_row + offset
                        
                        target_ws.merge_cells(start_row=min_row, start_column=range_.min_col,
                                            end_row=max_row, end_column=range_.max_col)

            return start_row
            
        except Exception as e:
            print(f"Error copying styles: {e}")
            return start_row
