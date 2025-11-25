"""
Node Registry - manages all available node types
"""

from typing import Dict, Type, List
from .base_node import BaseNode


class NodeRegistry:
    """Registry for all available node types"""
    
    _instance = None
    _nodes: Dict[str, Type[BaseNode]] = {}
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    @classmethod
    def register(cls, node_class: Type[BaseNode]):
        """Register a node class"""
        cls._nodes[node_class.node_type] = node_class
        return node_class
    
    @classmethod
    def get_node_class(cls, node_type: str) -> Type[BaseNode]:
        """Get a node class by type"""
        return cls._nodes.get(node_type)
    
    @classmethod
    def create_node(cls, node_type: str, node_id: str) -> BaseNode:
        """Create a new node instance"""
        node_class = cls.get_node_class(node_type)
        if node_class:
            return node_class(node_id)
        raise ValueError(f"Unknown node type: {node_type}")
    
    @classmethod
    def get_all_nodes(cls) -> Dict[str, Type[BaseNode]]:
        """Get all registered nodes"""
        return cls._nodes.copy()
    
    @classmethod
    def get_nodes_by_category(cls) -> Dict[str, List[Type[BaseNode]]]:
        """Get nodes organized by category"""
        categories = {}
        for node_class in cls._nodes.values():
            category = node_class.node_category
            if category not in categories:
                categories[category] = []
            categories[category].append(node_class)
        return categories


# Decorator for easy registration
def register_node(cls: Type[BaseNode]) -> Type[BaseNode]:
    """Decorator to register a node class"""
    NodeRegistry.register(cls)
    return cls
