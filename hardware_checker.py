import os
import GPUtil
import subprocess
import re
from tkinter import messagebox # For displaying warnings/info

class HardwareChecker:
    # Checks available GPU VRAM to warn users if it's below recommended levels for OCR/AI tasks.
    def check_gpu_vram(self, required_vram_gb=4):
        # Parameters:
        #   required_vram_gb (float): The minimum recommended VRAM in Gigabytes.

        gpu_vram_gb = 0  # Initialize detected VRAM to 0 GB.
        checked_method = "N/A" # Initialize method used to detect VRAM.

        try:
            # Attempt to get GPU information using GPUtil library.
            gpus = GPUtil.getGPUs()
            if gpus:
                # If GPUs are found, get VRAM of the first GPU.
                gpu_vram_gb = gpus[0].memoryTotal / 1024  # GPUtil's memoryTotal is in MB.
                checked_method = "GPUtil"
            else:
                # If GPUtil runs but finds no GPUs, raise an exception to try nvidia-smi.
                raise Exception("No GPUs found by GPUtil") 
        except Exception as e_gputil:
            # GPUtil failed or found no GPUs, try using nvidia-smi command as a fallback.
            # print(f"GPUtil failed: {e_gputil}, trying nvidia-smi...") # For debugging
            try:
                # Set creation flags to prevent console window from popping up on Windows.
                cflags = subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0
                # Run nvidia-smi command to query total GPU memory.
                result = subprocess.run(
                    ['nvidia-smi', '--query-gpu=memory.total', '--format=csv,noheader,nounits'],
                    capture_output=True, text=True, check=True, creationflags=cflags
                )
                # The output is typically like "8192" (MiB).
                # This handles the first GPU if multiple are present. For more complex multi-GPU setups,
                # parsing might need to be more sophisticated (e.g., sum or average VRAM).
                first_line = result.stdout.strip().split('\n')[0] # Get the first line of output.
                if first_line:
                    gpu_vram_gb = int(first_line) / 1024 # Convert MiB to GB.
                    checked_method = "nvidia-smi"
            except Exception as e_nvidia_smi:
                # nvidia-smi failed or is not installed.
                # print(f"nvidia-smi failed: {e_nvidia_smi}") # For debugging
                pass # Both methods failed to detect GPU VRAM.

        # Check if detected VRAM is positive but below the required threshold.
        if gpu_vram_gb > 0 and gpu_vram_gb < required_vram_gb:
            messagebox.showwarning(
                "Hardware Warning", # Dialog title
                f"Detected {gpu_vram_gb:.1f}GB VRAM using {checked_method}. "
                f"This is below the recommended {required_vram_gb}GB. "
                "Performance might be affected or OCR/AI models may fail."
            )
        elif checked_method == "N/A": # If both GPUtil and nvidia-smi failed.
             messagebox.showinfo(
                "Hardware Info", # Dialog title
                "Could not automatically determine GPU VRAM. "
                "Please ensure your hardware meets model requirements if you encounter issues."
            )
        elif gpu_vram_gb >= required_vram_gb:
            # If VRAM meets or exceeds requirements, print a message to console (for logging/debugging).
            print(f"VRAM check: {gpu_vram_gb:.1f}GB detected via {checked_method}. Meets {required_vram_gb}GB requirement.")
        # Note: If gpu_vram_gb is 0 but one of the methods ran successfully (e.g., nvidia-smi reported 0 for integrated graphics),
        # it will be caught by the first conditional (gpu_vram_gb < required_vram_gb) if required_vram_gb > 0.
