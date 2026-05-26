"""CAD (DWG) to PDF conversion toolkit.

兩種 backend：
    A. ODA File Converter + ezdxf + matplotlib（免費，需裝 ODA 工具）
    B. AutoCAD COM 自動化（需裝完整版 AutoCAD + pywin32）

對外公開 API：
    # ODA 路線
    from cad2pdf import convert_dwg, convert_folder, ConvertMode
    from cad2pdf import find_oda_converter, OdaNotFoundError

    # AutoCAD 路線
    from cad2pdf import acad_convert_dwg, acad_convert_folder
    from cad2pdf import is_autocad_available, AutoCadNotAvailableError
"""

from .converter import convert_dwg, convert_folder, ConvertMode
from .oda import (
    find_oda_converter,
    find_bundled_installer,
    run_installer,
    OdaNotFoundError,
)
from .autocad import (
    acad_convert_dwg,
    acad_convert_folder,
    is_autocad_available,
    AutoCadNotAvailableError,
)

__all__ = [
    # ODA
    "convert_dwg",
    "convert_folder",
    "ConvertMode",
    "find_oda_converter",
    "find_bundled_installer",
    "run_installer",
    "OdaNotFoundError",
    # AutoCAD
    "acad_convert_dwg",
    "acad_convert_folder",
    "is_autocad_available",
    "AutoCadNotAvailableError",
]
