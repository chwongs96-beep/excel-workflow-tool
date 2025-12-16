"""
Workflow Engine - executes the node workflow
"""

from typing import Dict, List, Any, Optional, Set
from collections import deque
import json
from pathlib import Path

from ..nodes.base_node import BaseNode
from ..nodes.node_registry import NodeRegistry


class Connection:
    """Represents a connection between two ports"""
    
    def __init__(self, from_node: str, from_port: str, to_node: str, to_port: str):
        self.from_node = from_node
        self.from_port = from_port
        self.to_node = to_node
        self.to_port = to_port
    
    def to_dict(self) -> Dict[str, str]:
        return {
            "from_node": self.from_node,
            "from_port": self.from_port,
            "to_node": self.to_node,
            "to_port": self.to_port
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, str]) -> "Connection":
        return cls(
            data["from_node"],
            data["from_port"],
            data["to_node"],
            data["to_port"]
        )


class Workflow:
    """Represents a complete workflow with nodes and connections"""
    
    def __init__(self, name: str = "Untitled Workflow"):
        self.name = name
        self.nodes: Dict[str, BaseNode] = {}
        self.connections: List[Connection] = []
        self.global_params: Dict[str, str] = {}  # Global parameters for substitution
        self._node_counter = 0
    
    def generate_node_id(self) -> str:
        """Generate a unique node ID"""
        self._node_counter += 1
        return f"node_{self._node_counter}"
    
    def add_node(self, node_type: str, position: tuple = (0, 0)) -> BaseNode:
        """Add a new node to the workflow"""
        node_id = self.generate_node_id()
        node = NodeRegistry.create_node(node_type, node_id)
        node.position = position
        self.nodes[node_id] = node
        return node
    
    def remove_node(self, node_id: str):
        """Remove a node and all its connections"""
        if node_id in self.nodes:
            # Remove all connections involving this node
            self.connections = [
                c for c in self.connections
                if c.from_node != node_id and c.to_node != node_id
            ]
            del self.nodes[node_id]
    
    def add_connection(self, from_node: str, from_port: str, 
                       to_node: str, to_port: str) -> Optional[Connection]:
        """Add a connection between two nodes"""
        # Validate nodes exist
        if from_node not in self.nodes or to_node not in self.nodes:
            return None
        
        # Validate ports exist
        source = self.nodes[from_node]
        target = self.nodes[to_node]
        
        if not source.get_output_port(from_port):
            return None
        if not target.get_input_port(to_port):
            return None
        
        # Check for existing connection to the same input port
        for conn in self.connections:
            if conn.to_node == to_node and conn.to_port == to_port:
                # Remove existing connection
                self.connections.remove(conn)
                break
        
        # Create new connection
        conn = Connection(from_node, from_port, to_node, to_port)
        self.connections.append(conn)
        return conn
    
    def remove_connection(self, connection: Connection):
        """Remove a connection"""
        if connection in self.connections:
            self.connections.remove(connection)
    
    def get_execution_order(self) -> List[str]:
        """Get the topological order of nodes for execution"""
        # Build adjacency list and in-degree count
        in_degree: Dict[str, int] = {node_id: 0 for node_id in self.nodes}
        adjacency: Dict[str, List[str]] = {node_id: [] for node_id in self.nodes}
        
        for conn in self.connections:
            adjacency[conn.from_node].append(conn.to_node)
            in_degree[conn.to_node] += 1
        
        # Kahn's algorithm
        queue = deque([node_id for node_id, deg in in_degree.items() if deg == 0])
        result = []
        
        while queue:
            node_id = queue.popleft()
            result.append(node_id)
            
            for neighbor in adjacency[node_id]:
                in_degree[neighbor] -= 1
                if in_degree[neighbor] == 0:
                    queue.append(neighbor)
        
        if len(result) != len(self.nodes):
            raise ValueError("Workflow contains a cycle!")
        
        return result
    
    def get_ancestors(self, node_id: str) -> Set[str]:
        """Get all ancestor nodes of a given node"""
        ancestors = set()
        queue = deque([node_id])
        visited = {node_id}
        
        while queue:
            current = queue.popleft()
            for conn in self.connections:
                if conn.to_node == current:
                    if conn.from_node not in visited:
                        visited.add(conn.from_node)
                        ancestors.add(conn.from_node)
                        queue.append(conn.from_node)
        
        return ancestors

    def execute_node(self, target_node_id: str, progress_callback=None) -> Dict[str, Any]:
        """Execute the workflow up to a specific node"""
        if target_node_id not in self.nodes:
            raise ValueError(f"Node {target_node_id} not found")
            
        # Get all ancestors + target node
        nodes_to_execute = self.get_ancestors(target_node_id)
        nodes_to_execute.add(target_node_id)
        
        # Get full execution order
        full_order = self.get_execution_order()
        
        # Filter order to only include relevant nodes
        execution_order = [nid for nid in full_order if nid in nodes_to_execute]
        
        node_outputs: Dict[str, Dict[str, Any]] = {}
        results = {}
        
        total_nodes = len(execution_order)
        
        for i, node_id in enumerate(execution_order):
            node = self.nodes[node_id]
            
            # Inject global params as context
            node.set_context(self.global_params)
            
            # Validate node
            is_valid, error = node.validate()
            if not is_valid:
                raise ValueError(f"Node '{node.node_name}' ({node_id}): {error}")
            
            # Gather input data from connected nodes
            input_data = {}
            for conn in self.connections:
                if conn.to_node == node_id:
                    # Only consider connections from nodes we are executing
                    # (Though ancestors logic ensures they are in execution_order)
                    source_outputs = node_outputs.get(conn.from_node, {})
                    if conn.from_port in source_outputs:
                        input_data[conn.to_port] = source_outputs[conn.from_port]
            
            # Execute node
            try:
                output = node.execute(input_data)
                node_outputs[node_id] = output
                results[node_id] = {
                    "success": True,
                    "output": output
                }
            except Exception as e:
                results[node_id] = {
                    "success": False,
                    "error": str(e)
                }
                raise RuntimeError(f"Error in node '{node.node_name}': {e}")
            
            # Progress callback
            if progress_callback:
                progress_callback(i + 1, total_nodes, node.node_name, node_id)
        
        return results

    def execute(self, progress_callback=None) -> Dict[str, Any]:
        """Execute the entire workflow"""
        execution_order = self.get_execution_order()
        node_outputs: Dict[str, Dict[str, Any]] = {}
        results = {}
        
        total_nodes = len(execution_order)
        
        for i, node_id in enumerate(execution_order):
            node = self.nodes[node_id]
            
            # Inject global params as context
            node.set_context(self.global_params)
            
            # Validate node
            is_valid, error = node.validate()
            if not is_valid:
                raise ValueError(f"Node '{node.node_name}' ({node_id}): {error}")
            
            # Gather input data from connected nodes
            input_data = {}
            for conn in self.connections:
                if conn.to_node == node_id:
                    source_outputs = node_outputs.get(conn.from_node, {})
                    if conn.from_port in source_outputs:
                        input_data[conn.to_port] = source_outputs[conn.from_port]
            
            # Execute node
            try:
                output = node.execute(input_data)
                node_outputs[node_id] = output
                results[node_id] = {
                    "success": True,
                    "output": output
                }
            except Exception as e:
                results[node_id] = {
                    "success": False,
                    "error": str(e)
                }
                raise RuntimeError(f"Error in node '{node.node_name}': {e}")
            
            # Progress callback
            if progress_callback:
                progress_callback(i + 1, total_nodes, node.node_name, node_id)
        
        return results
    
    def to_dict(self) -> Dict[str, Any]:
        """Serialize workflow to dictionary"""
        return {
            "name": self.name,
            "node_counter": self._node_counter,
            "global_params": self.global_params,
            "nodes": {
                node_id: node.to_dict() 
                for node_id, node in self.nodes.items()
            },
            "connections": [conn.to_dict() for conn in self.connections]
        }
    
    def save(self, file_path: str):
        """Save workflow to file"""
        with open(file_path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, indent=2)
    
    @classmethod
    def load(cls, file_path: str) -> "Workflow":
        """Load workflow from file"""
        with open(file_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        
        workflow = cls(data.get("name", "Untitled"))
        workflow._node_counter = data.get("node_counter", 0)
        workflow.global_params = data.get("global_params", {})
        
        # Recreate nodes
        for node_id, node_data in data.get("nodes", {}).items():
            node_type = node_data["node_type"]
            node = NodeRegistry.create_node(node_type, node_id)
            node.from_dict(node_data)
            workflow.nodes[node_id] = node
        
        # Recreate connections
        for conn_data in data.get("connections", []):
            conn = Connection.from_dict(conn_data)
            workflow.connections.append(conn)
        
        return workflow
