"""
Worker Thread - Background thread for workflow execution
"""

from typing import Dict, Any, Optional
from PyQt6.QtCore import QThread, pyqtSignal


class WorkflowWorker(QThread):
    """Background worker for executing workflows without blocking UI"""
    
    # Signals
    progress_updated = pyqtSignal(int, int, str, str, str)  # current, total, node_name, node_id, detail_msg
    execution_finished = pyqtSignal(dict)  # results
    execution_failed = pyqtSignal(str)  # error message
    node_status_changed = pyqtSignal(str, str)  # node_id, status
    
    def __init__(self, workflow, external_context: Optional[Dict[str, Any]] = None):
        super().__init__()
        self.workflow = workflow
        self.external_context = external_context or {}
        self._is_cancelled = False
    
    def run(self):
        """Execute the workflow in background thread"""
        try:
            def progress_callback(current, total, node_name, node_id=None, detail_msg=None):
                if self._is_cancelled:
                    raise InterruptedError("Execution cancelled by user")
                
                # Emit progress signal
                detail = detail_msg or ""
                nid = node_id or ""
                self.progress_updated.emit(current, total, node_name, nid, detail)
            
            # Execute workflow
            results = self.workflow.execute(progress_callback, external_context=self.external_context)
            
            # Emit success
            if not self._is_cancelled:
                self.execution_finished.emit(results)
                
        except InterruptedError:
            # User cancelled
            self.execution_failed.emit("执行已被用户取消")
        except Exception as e:
            # Execution error
            self.execution_failed.emit(str(e))
    
    def cancel(self):
        """Cancel the execution"""
        self._is_cancelled = True
