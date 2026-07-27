import os
import sys

class JobLock:
    def __init__(self, lock_file_path: str = "temp/app.lock"):
        self.root_dir = os.path.dirname(os.path.abspath(__file__))
        if not os.path.isabs(lock_file_path):
            self.lock_file = os.path.join(self.root_dir, lock_file_path)
        else:
            self.lock_file = lock_file_path
            
        os.makedirs(os.path.dirname(self.lock_file), exist_ok=True)

    def acquire(self) -> bool:
        """Try to acquire the lock. Returns True if successful, False otherwise."""
        if os.path.exists(self.lock_file):
            # Read PID in lock file
            try:
                with open(self.lock_file, "r") as f:
                    pid_str = f.read().strip()
                if pid_str.isdigit():
                    pid = int(pid_str)
                    if self._is_pid_running(pid):
                        return False
            except Exception:
                pass # If file is unreadable, assume we can overwrite
                
        # Write current PID
        try:
            with open(self.lock_file, "w") as f:
                f.write(str(os.getpid()))
            return True
        except Exception:
            return False

    def release(self):
        """Release the lock by deleting the lock file."""
        if os.path.exists(self.lock_file):
            try:
                # Only delete if it belongs to our PID
                with open(self.lock_file, "r") as f:
                    pid_str = f.read().strip()
                if pid_str.isdigit() and int(pid_str) == os.getpid():
                    os.remove(self.lock_file)
            except Exception:
                pass

    def _is_pid_running(self, pid: int) -> bool:
        """Check if a process with the given PID is running on Windows."""
        import ctypes
        # Get exit code of process
        PROCESS_QUERY_INFORMATION = 0x0400
        STILL_ACTIVE = 259
        kernel32 = ctypes.windll.kernel32
        
        handle = kernel32.OpenProcess(PROCESS_QUERY_INFORMATION, False, pid)
        if handle == 0:
            return False
            
        exit_code = ctypes.c_ulong()
        res = kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code))
        kernel32.CloseHandle(handle)
        
        if res == 0:
            return False
        return exit_code.value == STILL_ACTIVE
