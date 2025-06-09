import pandas as pd
import openpyxl
import xlrd
import os
import re
import time
from pathlib import Path
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# 全局变量用于统计
stats = {
    "total_files": 0,
    "processed_files": 0,
    "matched_budgets": 0,
    "unmatched_budgets": 0,
    "extracted_from_filename": 0,
    "missing_data": {
        "事业部预算编号": 0,
        "合同号": 0,
        "部门（显示值）": 0,
        "单据编号": 0,
        "备注": 0
    }
}

def extract_filenames_to_excel(folder_path, output_file="文件名列表.xlsx", extract_content=False):
    """
    提取指定文件夹中所有文件名并保存到Excel文件
    
    Args:
        folder_path: 文件夹路径（字符串或Path对象）
        output_file: 输出Excel文件名
        extract_content: 是否提取Excel文件内容
    """
    try:
        start_time = time.time()
        
        # 重置统计数据
        global stats
        stats = {
            "total_files": 0,
            "processed_files": 0,
            "matched_budgets": 0,
            "unmatched_budgets": 0,
            "extracted_from_filename": 0,
            "missing_data": {
                "事业部预算编号": 0,
                "合同号": 0,
                "部门（显示值）": 0,
                "单据编号": 0,
                "备注": 0
            }
        }
        
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
        excel_files = [f for f in folder.iterdir() if f.is_file() and f.suffix.lower() in ['.xls', '.xlsx']]
        stats["total_files"] = len(excel_files)
        
        # 进度显示
        total_files = len(excel_files)
        processed = 0
        
        for file in excel_files:
            # 更新进度
            processed += 1
            if processed % 10 == 0 or processed == total_files:
                print(f"处理进度: {processed}/{total_files} ({processed/total_files*100:.1f}%)")
            
            try:
                # 获取文件统计信息
                stat = file.stat()
                
                file_data = {
                    '文件名': file.name,
                    '文件名（无扩展名）': file.stem,
                    '扩展名': file.suffix,
                    '文件大小(字节)': stat.st_size,
                    '修改时间': pd.Timestamp.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
                }
                
                # 如果需要提取文件内容
                if extract_content:
                    try:
                        # 尝试读取Excel文件内容
                        content_data = extract_excel_content(file)
                        # 合并字典
                        file_data.update(content_data)
                    except Exception as e:
                        print(f"警告：无法从文件 {file.name} 提取内容: {e}")
                        # 创建空数据
                        file_data.update({
                            '事业部预算编号': '',
                            '合同号': '',
                            '部门（显示值）': '',
                            '单据编号': '',
                            '备注': ''
                        })
                
                file_info.append(file_data)
                stats["processed_files"] += 1
                
            except Exception as e:
                print(f"处理文件 {file.name} 时出错: {e}")
        
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
                cell_value = str(ws.cell(row=row, column=col).value or '')
                if len(cell_value) > max_length:
                    max_length = len(cell_value)
            
            # 设置列宽（根据内容长度计算，中文字符宽度需要调整）
            adjusted_width = max_length * 1.2 + 4  # 中文字符宽度调整系数
            ws.column_dimensions[column_letter].width = adjusted_width
        
        # 冻结首行
        ws.freeze_panes = "A2"
        
        # 添加统计信息到新工作表
        ws_stats = wb.create_sheet(title="统计信息")
        
        # 添加标题
        ws_stats['A1'] = "提取统计信息"
        ws_stats.merge_cells('A1:B1')
        ws_stats['A1'].font = Font(name='微软雅黑', size=14, bold=True)
        ws_stats['A1'].alignment = Alignment(horizontal='center')
        
        # 添加统计数据
        ws_stats['A3'] = "总文件数"
        ws_stats['B3'] = stats["total_files"]
        ws_stats['A4'] = "成功处理文件数"
        ws_stats['B4'] = stats["processed_files"]
        ws_stats['A5'] = "预算编号匹配文件数"
        ws_stats['B5'] = stats["matched_budgets"]
        ws_stats['A6'] = "预算编号不匹配文件数"
        ws_stats['B6'] = stats["unmatched_budgets"]
        ws_stats['A7'] = "从文件名提取预算编号数"
        ws_stats['B7'] = stats["extracted_from_filename"]
        
        # 缺失数据统计
        ws_stats['A9'] = "缺失数据统计"
        ws_stats.merge_cells('A9:B9')
        ws_stats['A9'].font = Font(bold=True)
        
        row = 10
        for field, count in stats["missing_data"].items():
            ws_stats[f'A{row}'] = f"缺失{field}的文件数"
            ws_stats[f'B{row}'] = count
            row += 1
        
        # 设置统计表格的列宽
        ws_stats.column_dimensions['A'].width = 25
        ws_stats.column_dimensions['B'].width = 15
        
        # 保存格式化后的Excel
        wb.save(output_path)
        
        end_time = time.time()
        elapsed_time = end_time - start_time
        
        print(f"完成：共提取 {len(file_info)} 个文件的详细信息")
        print(f"处理时间：{elapsed_time:.2f}秒")
        print(f"已保存到：{output_path.absolute()}")
        print("-" * 50)
        print("统计信息:")
        print(f"  总文件数: {stats['total_files']}")
        print(f"  成功处理文件数: {stats['processed_files']}")
        print(f"  预算编号匹配文件数: {stats['matched_budgets']}")
        print(f"  预算编号不匹配文件数: {stats['unmatched_budgets']}")
        print(f"  从文件名提取预算编号数: {stats['extracted_from_filename']}")
        print("  缺失数据统计:")
        for field, count in stats["missing_data"].items():
            print(f"    缺失{field}的文件数: {count}")
        print("-" * 50)
        
        return file_info
            
    except Exception as e:
        print(f"出错：{e}")
        import traceback
        traceback.print_exc()
        return []

def normalize_budget_id(text):
    """
    标准化预算编号格式，去除多余空格和特殊字符
    """
    if not text:
        return ""
        
    # 去除多余空格和非法字符
    text = re.sub(r'\s+', '', text)
    
    # 尝试匹配标准格式 WZ-FJ-YYYYMM-NNN
    pattern = r'([A-Z]{1,2})[-_]?([A-Z]{1,2})[-_]?(\d{6})[-_]?(\d{3})'
    match = re.search(pattern, text)
    if match:
        # 重新格式化为标准格式
        return f"{match.group(1)}-{match.group(2)}-{match.group(3)}-{match.group(4)}"
    
    return text

def extract_excel_content(file_path):
    """
    从Excel文件中提取特定内容
    
    Args:
        file_path: Excel文件路径
    
    Returns:
        包含提取内容的字典
    """
    try:
        file_ext = Path(file_path).suffix.lower()
        result = {
            '事业部预算编号': '',
            '合同号': '',
            '部门（显示值）': '',
            '单据编号': '',
            '备注': ''
        }
        
        if file_ext == '.xlsx':
            # 使用openpyxl读取.xlsx文件
            return extract_with_openpyxl(file_path, result)
        elif file_ext == '.xls':
            # 使用xlrd读取.xls文件
            return extract_with_xlrd(file_path, result)
        else:
            raise ValueError(f"不支持的文件格式: {file_ext}")
    
    except Exception as e:
        print(f"提取文件 {file_path} 内容时出错: {e}")
        return {
            '事业部预算编号': '',
            '合同号': '',
            '部门（显示值）': '',
            '单据编号': '',
            '备注': ''
        }

def find_value_by_keyword(worksheet, keywords, max_rows=50, max_cols=20):
    """
    在工作表中查找关键字并返回相应的值
    
    Args:
        worksheet: 工作表对象
        keywords: 关键字列表
        max_rows: 搜索的最大行数
        max_cols: 搜索的最大列数
        
    Returns:
        找到的值或空字符串
    """
    # 标准化关键字
    normalized_keywords = [k.strip().lower() for k in keywords if k]
    
    # 对openpyxl工作表的处理
    if hasattr(worksheet, 'iter_rows'):
        # 首先，尝试查找精确匹配的单元格
        for row_idx, row in enumerate(worksheet.iter_rows(min_row=1, max_row=max_rows, max_col=max_cols), 1):
            for col_idx, cell in enumerate(row, 1):
                if cell.value:
                    cell_text = str(cell.value).strip().lower()
                    for keyword in normalized_keywords:
                        # 精确匹配关键字
                        if cell_text == keyword or cell_text == keyword + "：" or cell_text == keyword + ":":
                            # 检查右侧单元格
                            if col_idx < max_cols:
                                right_cell = worksheet.cell(row=row_idx, column=col_idx+1)
                                if right_cell.value:
                                    return str(right_cell.value).strip()
                            
                            # 检查下方单元格
                            if row_idx < max_rows:
                                below_cell = worksheet.cell(row=row_idx+1, column=col_idx)
                                if below_cell.value:
                                    return str(below_cell.value).strip()

        # 然后，尝试查找包含关键字的单元格
        for row_idx, row in enumerate(worksheet.iter_rows(min_row=1, max_row=max_rows, max_col=max_cols), 1):
            for col_idx, cell in enumerate(row, 1):
                if cell.value:
                    cell_text = str(cell.value).strip().lower()
                    for keyword in normalized_keywords:
                        if keyword in cell_text:
                            # 检查右侧单元格
                            if col_idx < max_cols:
                                right_cell = worksheet.cell(row=row_idx, column=col_idx+1)
                                if right_cell.value:
                                    return str(right_cell.value).strip()
                            
                            # 检查下方单元格
                            if row_idx < max_rows:
                                below_cell = worksheet.cell(row=row_idx+1, column=col_idx)
                                if below_cell.value:
                                    return str(below_cell.value).strip()
                                    
                            # 尝试检查合并单元格的情况
                            try:
                                for merged_range in worksheet.merged_cells.ranges:
                                    min_col, min_row, max_col, max_row = merged_range.bounds
                                    if min_row <= row_idx <= max_row and min_col <= col_idx <= max_col:
                                        # 如果当前单元格在一个合并区域内
                                        # 检查右侧的第一个非合并单元格
                                        next_col = max_col + 1
                                        if next_col <= worksheet.max_column:
                                            right_cell = worksheet.cell(row=row_idx, column=next_col)
                                            if right_cell.value:
                                                return str(right_cell.value).strip()
                            except:
                                pass
                            
                            # 查找同行中的其他单元格
                            for check_col in range(1, min(max_cols, worksheet.max_column) + 1):
                                if check_col != col_idx:
                                    check_cell = worksheet.cell(row=row_idx, column=check_col)
                                    if check_cell.value and isinstance(check_cell.value, str) and len(check_cell.value.strip()) > 0:
                                        return str(check_cell.value).strip()
    
    # 对xlrd工作表的处理
    else:
        # 首先，尝试查找精确匹配的单元格
        for row_idx in range(min(max_rows, worksheet.nrows)):
            for col_idx in range(min(max_cols, worksheet.ncols)):
                cell_value = worksheet.cell_value(row_idx, col_idx)
                if cell_value:
                    cell_text = str(cell_value).strip().lower()
                    for keyword in normalized_keywords:
                        # 精确匹配关键字
                        if cell_text == keyword or cell_text == keyword + "：" or cell_text == keyword + ":":
                            # 检查右侧单元格
                            if col_idx + 1 < worksheet.ncols:
                                right_value = worksheet.cell_value(row_idx, col_idx + 1)
                                if right_value:
                                    return str(right_value).strip()
                            
                            # 检查下方单元格
                            if row_idx + 1 < worksheet.nrows:
                                below_value = worksheet.cell_value(row_idx + 1, col_idx)
                                if below_value:
                                    return str(below_value).strip()

        # 然后，尝试查找包含关键字的单元格
        for row_idx in range(min(max_rows, worksheet.nrows)):
            for col_idx in range(min(max_cols, worksheet.ncols)):
                cell_value = worksheet.cell_value(row_idx, col_idx)
                if cell_value:
                    cell_text = str(cell_value).strip().lower()
                    for keyword in normalized_keywords:
                        if keyword in cell_text:
                            # 检查右侧单元格
                            if col_idx + 1 < worksheet.ncols:
                                right_value = worksheet.cell_value(row_idx, col_idx + 1)
                                if right_value:
                                    return str(right_value).strip()
                            
                            # 检查下方单元格
                            if row_idx + 1 < worksheet.nrows:
                                below_value = worksheet.cell_value(row_idx + 1, col_idx)
                                if below_value:
                                    return str(below_value).strip()
                                    
                            # 查找同行中的其他单元格
                            for check_col in range(min(max_cols, worksheet.ncols)):
                                if check_col != col_idx:
                                    check_value = worksheet.cell_value(row_idx, check_col)
                                    if check_value and isinstance(check_value, str) and len(str(check_value).strip()) > 0:
                                        return str(check_value).strip()
    
    return ""

def extract_with_openpyxl(file_path, result):
    """使用openpyxl提取.xlsx文件内容"""
    try:
        # 使用openpyxl读取Excel文件
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb.active
        
        # 使用增强的关键字搜索提取内容
        result['事业部预算编号'] = find_value_by_keyword(ws, ['事业部预算编号', '预算编号', '预算号', '预算单号', '项目编号', '项目号', '事业部编号'])
        result['合同号'] = find_value_by_keyword(ws, ['合同号', '合同编号', '合同名称', '合同内容', '合同标的', '项目名称', '项目内容'])
        result['部门（显示值）'] = find_value_by_keyword(ws, ['部门（显示值）', '部门(显示值)', '部门', '使用部门', '申请部门', '所属部门', '责任部门'])
        result['单据编号'] = find_value_by_keyword(ws, ['单据编号', '单据号', '凭证号', '凭证编号', '发票号', '发票编号', '申请单号'])
        result['备注'] = find_value_by_keyword(ws, ['备注', '备注说明', '说明', '项目说明', '其他说明', '补充说明', '附注'])
        
        # 记录缺失数据统计
        for field, value in result.items():
            if not value:
                stats["missing_data"][field] += 1
        
        # 如果没有找到事业部预算编号，尝试从文件名提取
        if not result['事业部预算编号']:
            file_stem = Path(file_path).stem
            # 尝试从文件名中提取预算编号格式
            budget_id_match = re.search(r'([A-Z]{1,2})[-_]?([A-Z]{1,2})[-_]?(\d{6})[-_]?(\d{3})', file_stem)
            if budget_id_match:
                extracted_id = f"{budget_id_match.group(1)}-{budget_id_match.group(2)}-{budget_id_match.group(3)}-{budget_id_match.group(4)}"
                result['事业部预算编号'] = extracted_id
                stats["extracted_from_filename"] += 1
                print(f"✓ 从文件名 {file_stem} 提取预算编号: {extracted_id}")
        
        # 标准化预算编号
        result['事业部预算编号'] = normalize_budget_id(result['事业部预算编号'])
        
        # 验证事业部预算编号与文件名的关系（更宽松的匹配）
        file_stem = Path(file_path).stem
        if result['事业部预算编号']:
            # 标准化文件名中的预算编号格式
            normalized_stem = normalize_budget_id(file_stem)
            
            # 提取纯数字部分进行比较
            file_numbers = re.sub(r'[^0-9]', '', normalized_stem)
            budget_numbers = re.sub(r'[^0-9]', '', result['事业部预算编号'])
            
            # 如果数字部分包含关系，也算匹配
            if (budget_numbers in file_numbers) or (file_numbers in budget_numbers):
                print(f"√ 文件 {file_stem} 的事业部预算编号验证通过")
                stats["matched_budgets"] += 1
            else:
                print(f"! 警告：文件 {file_stem} 的事业部预算编号与文件名不匹配")
                stats["unmatched_budgets"] += 1
                # 使用文件名作为预算编号
                if not result['事业部预算编号'] and normalized_stem:
                    result['事业部预算编号'] = normalized_stem
                    print(f"  > 已使用文件名 {normalized_stem} 作为预算编号")
        
        return result
    
    except Exception as e:
        print(f"使用openpyxl提取 {file_path} 时出错: {e}")
        for field in result:
            stats["missing_data"][field] += 1
        return result

def extract_with_xlrd(file_path, result):
    """使用xlrd提取.xls文件内容"""
    try:
        # 使用xlrd读取Excel文件
        wb = xlrd.open_workbook(file_path)
        ws = wb.sheet_by_index(0)  # 获取第一个工作表
        
        # 使用增强的关键字搜索提取内容
        result['事业部预算编号'] = find_value_by_keyword(ws, ['事业部预算编号', '预算编号', '预算号', '预算单号', '项目编号', '项目号', '事业部编号'])
        result['合同号'] = find_value_by_keyword(ws, ['合同号', '合同编号', '合同名称', '合同内容', '合同标的', '项目名称', '项目内容'])
        result['部门（显示值）'] = find_value_by_keyword(ws, ['部门（显示值）', '部门(显示值)', '部门', '使用部门', '申请部门', '所属部门', '责任部门'])
        result['单据编号'] = find_value_by_keyword(ws, ['单据编号', '单据号', '凭证号', '凭证编号', '发票号', '发票编号', '申请单号'])
        result['备注'] = find_value_by_keyword(ws, ['备注', '备注说明', '说明', '项目说明', '其他说明', '补充说明', '附注'])
        
        # 记录缺失数据统计
        for field, value in result.items():
            if not value:
                stats["missing_data"][field] += 1
        
        # 如果没有找到事业部预算编号，尝试从文件名提取
        if not result['事业部预算编号']:
            file_stem = Path(file_path).stem
            # 尝试从文件名中提取预算编号格式
            budget_id_match = re.search(r'([A-Z]{1,2})[-_]?([A-Z]{1,2})[-_]?(\d{6})[-_]?(\d{3})', file_stem)
            if budget_id_match:
                extracted_id = f"{budget_id_match.group(1)}-{budget_id_match.group(2)}-{budget_id_match.group(3)}-{budget_id_match.group(4)}"
                result['事业部预算编号'] = extracted_id
                stats["extracted_from_filename"] += 1
                print(f"✓ 从文件名 {file_stem} 提取预算编号: {extracted_id}")
        
        # 标准化预算编号
        result['事业部预算编号'] = normalize_budget_id(result['事业部预算编号'])
        
        # 验证事业部预算编号与文件名的关系（更宽松的匹配）
        file_stem = Path(file_path).stem
        if result['事业部预算编号']:
            # 标准化文件名中的预算编号格式
            normalized_stem = normalize_budget_id(file_stem)
            
            # 提取纯数字部分进行比较
            file_numbers = re.sub(r'[^0-9]', '', normalized_stem)
            budget_numbers = re.sub(r'[^0-9]', '', result['事业部预算编号'])
            
            # 如果数字部分包含关系，也算匹配
            if (budget_numbers in file_numbers) or (file_numbers in budget_numbers):
                print(f"√ 文件 {file_stem} 的事业部预算编号验证通过")
                stats["matched_budgets"] += 1
            else:
                print(f"! 警告：文件 {file_stem} 的事业部预算编号与文件名不匹配")
                stats["unmatched_budgets"] += 1
                # 使用文件名作为预算编号
                if not result['事业部预算编号'] and normalized_stem:
                    result['事业部预算编号'] = normalized_stem
                    print(f"  > 已使用文件名 {normalized_stem} 作为预算编号")
        
        return result
    
    except Exception as e:
        print(f"使用xlrd提取 {file_path} 时出错: {e}")
        for field in result:
            stats["missing_data"][field] += 1
        return result

def create_test_files():
    """创建测试文件"""
    # 创建测试目录
    test_dir = Path(__file__).parent / "test_files"
    test_dir.mkdir(exist_ok=True)
    
    # 创建一个简单的.xlsx文件
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    
    # 添加标题行
    ws['A1'] = "预算单"
    ws.merge_cells('A1:E1')
    ws['A1'].alignment = Alignment(horizontal='center')
    
    # 添加内容
    ws['A3'] = "事业部预算编号"
    ws['B3'] = "WZ-FJ-202406-032"
    ws['A4'] = "合同号"
    ws['B4'] = "太平洋1#锅炉形滤网"
    ws['A5'] = "部门（显示值）"
    ws['B5'] = "辅机事业部"
    ws['A6'] = "单据编号"
    ws['B6'] = "WZBD20240197"
    ws['A7'] = "备注"
    ws['B7'] = "王志中"
    
    # 保存文件
    test_file_path = test_dir / "WZ-FJ-202406-032.xlsx"
    wb.save(test_file_path)
    
    print(f"已创建测试文件: {test_file_path}")
    return test_dir

def test():
    """测试函数"""
    print("开始测试...")
    # 创建测试文件
    test_dir = create_test_files()
    
    # 测试提取功能
    print("\n测试文件内容提取:")
    output_file = "test_output.xlsx"
    extract_filenames_to_excel(test_dir, output_file, extract_content=True)
    
    # 测试完成后删除测试文件
    output_path = Path(__file__).parent / output_file
    if output_path.exists():
        print(f"\n测试输出文件位置: {output_path.absolute()}")
        print("测试完成！")

# 使用示例
if __name__ == "__main__":
    # 运行测试
    test_mode = input("是否运行测试模式？(y/n，直接回车默认为n)：").strip().lower()
    if test_mode == 'y':
        test()
    else:
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
        
        # 是否提取Excel内容
        extract_content_input = input("是否提取Excel文件内容？(y/n，直接回车默认为y)：").strip().lower()
        extract_content = extract_content_input != 'n'
        
        # 提取文件名到Excel
        extract_filenames_to_excel(folder_path, output_file, extract_content) 
