"""啟動入口。

直接執行：
    python main.py

未來小工具整合時，可改為從外部 import：
    from cad2pdf import convert_dwg, convert_folder, ConvertMode
    from gui import App   # 若要復用同一個 GUI 視窗
"""

from gui import main

if __name__ == "__main__":
    main()
