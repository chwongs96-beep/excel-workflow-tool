import pandas as pd
from pathlib import Path
from typing import Any, Dict, List
from .base_node import BaseNode
from .node_registry import register_node

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
                workbook = pd.read_excel(base_file, sheet_name=None)
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
                "file_filter": "Excel文件 (*.xlsx *.xls)",
                "placeholder": "请选择要追加的Excel文件"
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
                "placeholder": "例如: 2023年报 (留空则匹配所有Excel)"
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
        if source_type == "file":
            if not self.get_param("file_path"):
                return False, "来源文件是必需的"
        else:
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
        
        if source_type == "file":
            file_path = self.get_param("file_path")
        else:
            folder = self.get_param("folder_path")
            keyword = self.get_param("keyword", "")
            
            if not folder or not Path(folder).exists():
                raise ValueError(f"文件夹不存在: {folder}")
                
            p = Path(folder)
            # Find excel files
            files = list(p.glob("*.xlsx")) + list(p.glob("*.xls"))
            
            # Filter by keyword
            if keyword:
                files = [f for f in files if keyword in f.name]
            
            if not files:
                raise ValueError(f"在 {folder} 中未找到匹配 '{keyword}' 的Excel文件")
            
            # Sort by name and take first
            files.sort(key=lambda f: f.name)
            file_path = str(files[0])
            print(f"Found file by keyword '{keyword}': {file_path}")

        sheet_mode = self.get_param("sheet_mode", "first")
        src_sheet_name = self.get_param("sheet_name", "")
        target_name = self.get_param("target_name", "")
        
        try:
            if sheet_mode == "all":
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
        
        if not workbook:
            # Empty workbook, create empty file? Or error?
            # Create empty dataframe
            workbook = {"Sheet1": pd.DataFrame()}
            
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            for sheet_name, df in workbook.items():
                # Excel sheet name limit is 31 chars
                safe_name = str(sheet_name)[:31]
                # Ensure unique in case truncation caused duplicates
                # (Simple check, writer might handle or throw error)
                df.to_excel(writer, sheet_name=safe_name, index=False)
                
        return {"file_path": output_file}
