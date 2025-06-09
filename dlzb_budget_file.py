import pandas as pd
from pathlib import Path
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

def extract_filenames_to_excel(folder_path, output_file="文件名列表.xlsx"):
    """
    提取指定文件夹中所有文件名并保存到Excel文件
    
    Args:
        folder_path: 文件夹路径（字符串或Path对象）
        output_file: 输出Excel文件名
    """
    try:
        # 确保folder_path是Path对象
        folder = Path(folder_path)
        
        # 检查文件夹是否存在
        if not folder.exists():
            print(f"错误：文件夹 {folder} 不存在！")
            return
        if not folder.is_dir():
            print(f"错误：{folder} 不是一个文件夹！")
            return
        
        print("开始提取文件名...")
        print(f"目标文件夹：{folder}")
        print("-" * 50)
        
        # 获取所有文件的详细信息
        file_info = []
        for file in folder.iterdir():
            if file.is_file():
                # 获取文件统计信息
                stat = file.stat()
                file_info.append({
                    '文件名': file.name,
                    '文件名（无扩展名）': file.stem,
                    '扩展名': file.suffix,
                    '文件大小(字节)': stat.st_size,
                    '修改时间': pd.Timestamp.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
                })
        
        # 创建DataFrame
        df = pd.DataFrame(file_info)
        
        # 确保输出路径在当前项目文件夹中
        current_dir = Path(__file__).parent
        output_path = current_dir / output_file
        
        # 保存到Excel（不带格式）
        df.to_excel(output_path, index=False, engine='openpyxl')
        
        # 打开工作簿进行格式优化
        from openpyxl import load_workbook
        wb = load_workbook(output_path)
        ws = wb.active
        
        # 定义样式
        header_font = Font(name='微软雅黑', size=11, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        # 定义边框样式
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # 设置表头样式
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_alignment
            cell.border = thin_border
        
        # 设置数据区域样式
        for row in range(2, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                cell.alignment = Alignment(horizontal='left', vertical='center')
                cell.border = thin_border
        
        # 自动调整列宽
        for col in range(1, ws.max_column + 1):
            column_letter = get_column_letter(col)
            # 获取该列最长内容的长度
            max_length = 0
            for row in range(1, ws.max_row + 1):
                cell_value = str(ws.cell(row=row, column=col).value)
                if len(cell_value) > max_length:
                    max_length = len(cell_value)
            
            # 设置列宽（根据内容长度计算，中文字符宽度需要调整）
            adjusted_width = max_length + 4
            ws.column_dimensions[column_letter].width = adjusted_width
        
        # 冻结首行
        ws.freeze_panes = "A2"
        
        # 保存格式化后的Excel
        wb.save(output_path)
        
        print(f"完成：共提取 {len(file_info)} 个文件的详细信息")
        print(f"已保存到：{output_path.absolute()}")
        print("-" * 50)
        
        return file_info
            
    except Exception as e:
        print(f"出错：{e}")
        return []

# 使用示例
if __name__ == "__main__":
    # 获取用户输入的文件夹路径
    folder_input = input("请输入要提取文件名的文件夹路径（直接回车使用默认路径）：").strip()
    
    if folder_input:
        folder_path = Path(folder_input)
    else:
        # 默认路径
        default_path = Path(r"E:\liu\Documents\WPSDrive\201050461\WPS云盘\工作项目\11.电力装备\04上线试用\预算导出\2024预算导出")
        folder_path = default_path
        print(f"使用默认路径：{folder_path}")
    
    # 获取用户输入的输出文件名
    output_input = input("请输入输出文件名(直接回车使用默认'文件名列表.xlsx')：").strip()
    output_file = output_input if output_input else "文件名列表.xlsx"
    
    # 提取文件名到Excel
    extract_filenames_to_excel(folder_path, output_file) 
