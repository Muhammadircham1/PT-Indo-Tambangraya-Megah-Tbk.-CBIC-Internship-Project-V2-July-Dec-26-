# process.py
from PyQt6.QtCore import QObject, pyqtSignal
import subprocess
import sys

class ProcessWorker(QObject):
    """
    Runs the script at script_path as an external process, forwarding stdout/stderr
    lines to PyQt signals for real-time output handling.
    """
    log = pyqtSignal(str)       # Signal emitted for each line of output from the process
    error = pyqtSignal(str)     # Signal emitted if an internal error occurs in this worker
    finished = pyqtSignal(int)  # Signal emitted when process finishes, including exit code

    def __init__(self, script_path: str):
        """
        Initialize the worker with the script path to execute.

        Args:
            script_path (str): Path to the Python script to run.
        """
        super().__init__()
        self.script_path = script_path

    def run(self):
        """
        Execute the script as a subprocess, emitting output lines and error notifications.

        The standard output and error streams are combined and emitted line by line
        through the `log` signal. If an exception is caught, `error` is emitted.
        Finally, the `finished` signal is emitted with the process's exit code.
        """
        try:
            # Launch the subprocess to run the specified Python script
            proc = subprocess.Popen(
                [sys.executable, self.script_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1,
            )

            # Read the subprocess output line by line and emit it via the log signal
            for line in proc.stdout:
                self.log.emit(line.rstrip())

            # Close the subprocess's stdout pipe
            proc.stdout.close()

            # Wait for subprocess completion
            proc.wait()

            # Emit the exit code when the process finishes
            self.finished.emit(proc.returncode)

        except Exception as e:
            # Emit error signal if an exception occurs during execution
            self.error.emit(str(e))

            # Use -1 as exit code in case of error
            self.finished.emit(-1)