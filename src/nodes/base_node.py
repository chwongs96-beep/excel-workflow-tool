"""
Base Node class for all workflow nodes
"""

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional
from enum import Enum
import pandas as pd


class PortType(Enum):
    INPUT = "input"
    OUTPUT = "output"


@dataclass
class Port:
    """Represents an input or output port on a node"""
    name: str
    port_type: PortType
    data_type: str = "dataframe"  # dataframe, value, any
    connected_to: Optional[str] = None  # node_id:port_name
    

@dataclass
class NodeConfig:
    """Configuration for a node"""
    params: Dict[str, Any] = field(default_factory=dict)


class BaseNode:
    """Base class for all nodes in the workflow"""
    
    # Class attributes - override in subclasses
    node_type: str = "base"
    node_name: str = "Base Node"
    node_category: str = "General"
    node_description: str = "Base node description"
    # 简短说明常见端口接线（可选），显示在节点配置面板
    node_wiring_hint: str = ""
    node_color: str = "#6366f1"  # Default indigo color
    
    def __init__(self, node_id: str):
        self.node_id = node_id
        self.inputs: List[Port] = []
        self.outputs: List[Port] = []
        self.config = NodeConfig()
        self.position = (0, 0)
        self._output_data: Dict[str, Any] = {}
        self._context: Dict[str, Any] = {}  # Execution context (e.g. global params)
        self._progress_callback = None # Callback for reporting internal progress
        
        # Initialize ports
        self._setup_ports()
    
    def set_context(self, context: Dict[str, Any]):
        """Set execution context"""
        self._context = context

    def set_progress_callback(self, callback):
        """Set callback for reporting internal progress"""
        self._progress_callback = callback
        
    def report_progress(self, message: str):
        """Report progress if callback is set"""
        if self._progress_callback:
            self._progress_callback(message)

    def _setup_ports(self):
        """Override in subclasses to set up input/output ports"""
        pass
    
    def add_input(self, name: str, data_type: str = "dataframe"):
        """Add an input port"""
        self.inputs.append(Port(name=name, port_type=PortType.INPUT, data_type=data_type))
    
    def add_output(self, name: str, data_type: str = "dataframe"):
        """Add an output port"""
        self.outputs.append(Port(name=name, port_type=PortType.OUTPUT, data_type=data_type))
    
    def get_input_port(self, name: str) -> Optional[Port]:
        """Get input port by name"""
        for port in self.inputs:
            if port.name == name:
                return port
        return None
    
    def get_output_port(self, name: str) -> Optional[Port]:
        """Get output port by name"""
        for port in self.outputs:
            if port.name == name:
                return port
        return None
    
    def set_param(self, key: str, value: Any):
        """Set a configuration parameter"""
        self.config.params[key] = value
    
    def get_param(self, key: str, default: Any = None) -> Any:
        """Get a configuration parameter with variable substitution"""
        val = self.config.params.get(key, default)
        
        # Perform substitution if value is string and context is available
        if isinstance(val, str) and self._context:
            try:
                original_val = val
                # Simple substitution for {var} - case insensitive
                # We iterate over context keys to replace
                for k, v in self._context.items():
                    # Support both {KEY} and {key}
                    placeholder_upper = "{" + k.upper() + "}"
                    placeholder_lower = "{" + k.lower() + "}"
                    placeholder_exact = "{" + k + "}"
                    
                    # Replace all variations
                    if placeholder_upper in val:
                        val = val.replace(placeholder_upper, str(v))
                    if placeholder_lower in val:
                        val = val.replace(placeholder_lower, str(v))
                    if placeholder_exact in val and placeholder_exact not in [placeholder_upper, placeholder_lower]:
                        val = val.replace(placeholder_exact, str(v))
                
                # Debug: report if substitution occurred
                if val != original_val and self._progress_callback:
                    self._progress_callback(f"参数替换: {original_val} -> {val}")
                    
            except Exception as e:
                # Log substitution errors
                if self._progress_callback:
                    self._progress_callback(f"参数替换失败: {e}")
                
        return val
    
    def execute(self, input_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Execute the node's operation.
        
        Args:
            input_data: Dictionary mapping input port names to their data
            
        Returns:
            Dictionary mapping output port names to their data
        """
        raise NotImplementedError("Subclasses must implement execute()")
    
    def validate(self) -> tuple[bool, str]:
        """
        Validate node configuration.
        
        Returns:
            Tuple of (is_valid, error_message)
        """
        return True, ""
    
    def get_config_ui_schema(self) -> List[Dict[str, Any]]:
        """
        Return UI schema for node configuration.
        Override in subclasses to provide custom configuration UI.
        
        Returns:
            List of field definitions for the config UI
        """
        return []
    
    def to_dict(self) -> Dict[str, Any]:
        """Serialize node to dictionary"""
        return {
            "node_id": self.node_id,
            "node_type": self.node_type,
            "position": self.position,
            "config": self.config.params,
            "inputs": [
                {"name": p.name, "connected_to": p.connected_to}
                for p in self.inputs
            ],
            "outputs": [
                {"name": p.name, "connected_to": p.connected_to}
                for p in self.outputs
            ]
        }
    
    def from_dict(self, data: Dict[str, Any]):
        """Deserialize node from dictionary"""
        self.position = tuple(data.get("position", (0, 0)))
        self.config.params = data.get("config", {})
        
        # Restore connections
        for input_data in data.get("inputs", []):
            port = self.get_input_port(input_data["name"])
            if port:
                port.connected_to = input_data.get("connected_to")
        
        for output_data in data.get("outputs", []):
            port = self.get_output_port(output_data["name"])
            if port:
                port.connected_to = output_data.get("connected_to")
