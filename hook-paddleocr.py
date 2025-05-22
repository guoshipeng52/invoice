import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Collect data files from the main paddleocr package
datas = collect_data_files('paddleocr', include_py_files=True)

# Collect all submodules for paddleocr
hiddenimports = collect_submodules('paddleocr')

# Add common hidden imports known for paddleocr and its dependencies
hiddenimports += [
    'paddle',
    'paddle.fluid',
    'paddle.fluid.core',
    'cv2', # OpenCV is a common dependency
    'skimage',
    'skimage.measure',
    'shapely',
    'pyclipper'
]

# Attempt to include PaddleOCR models and necessary files from the typical user directory.
paddleocr_home_path_str = ''
if os.name == 'nt': # Windows
    paddleocr_home_path_str = os.path.join(os.path.expanduser('~'), '.paddleocr', 'whl')
else: # Linux/macOS
    paddleocr_home_path_str = os.path.expanduser('~/.paddleocr/whl')

if os.path.exists(paddleocr_home_path_str):
    datas += [(paddleocr_home_path_str, 'paddleocr/whl')]
    print(f"INFO: Added PaddleOCR model directory {paddleocr_home_path_str} to PyInstaller datas.")
else:
    print(f"WARNING: PaddleOCR model directory not found at {paddleocr_home_path_str}. Packaging might fail or the EXE might not find models.")

# Ensure GPUtil is included
# datas += collect_data_files('GPUtil', include_py_files=True) # GPUtil usually bundles well without this if imported.
hiddenimports += collect_submodules('GPUtil')
hiddenimports += ['GPUtil'] # Explicitly add GPUtil as a hidden import

print(f"PyInstaller Hook (hook-paddleocr.py): datas={datas}")
print(f"PyInstaller Hook (hook-paddleocr.py): hiddenimports={hiddenimports}")
