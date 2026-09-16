'''
file: IYATT_BatchDeleteRowsOrColsByContent.py
description: 在单行的整行或单列的整列中检索关键字，匹配则删除对应的列或行
author: IYATT-yx
copyright: Copyright (c) 2026 IYATT-yx.
            Licensed under the MIT License. See LICENSE file in the project root for full license information.
'''
import tkinter as tk
from tkinter import messagebox, simpledialog
from pytableenginesdk import utils

pluginInfo = {
    'name': 'BatchDeleteRowsOrColsByContent',
    'author': 'IYATT',
    'description': '在单行的整行或单列的整列中检索关键字，匹配则删除对应的列或行',
    'version': '0.0.1',
}

def checkMatch(cellValue, targetKeywords):
    """辅助函数：检查单元格文本内容是否包含任意目标关键字"""
    if cellValue is None:
        return False
        
    val = str(cellValue).strip()
    if not val:
        return False
    
    return any(kw in val for kw in targetKeywords)

def getUserInputKeywords():
    """弹窗获取用户想要删除的内容关键字"""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)

    promptMsg = "请输入要检索的内容关键字：\n（选单行将删除匹配单元格所在的【列】；选单列将删除匹配单元格所在的【行】）\n多个关键字请用逗号分隔"
    userInput = simpledialog.askstring("删除条件设置", promptMsg, parent=root)
    root.destroy()

    if not userInput:
        return None
    
    rawKeywords = userInput.replace("，", ",").split(",")
    keywords = [kw.strip() for kw in rawKeywords if kw.strip()]
    return set(keywords)

def checkSingleFullRowOrCol(targets) -> bool | str:
    """
    注入式校验：检查用户是否选择了【单行的整行】或【单列的整列】
    """
    # 1. 校验提取列表中是否有且仅有 1 个选区项（替代 singleMode 的作用）
    if len(targets) != 1:
        return f"选区错误：当前提取了 {len(targets)} 个选区！\n请先在界面上移除多余的项，仅保留 1 个选区。"

    sheetName, address, mainRng = targets[0]

    # 2. 校验是否包含了非连续的多块区域（如按 Ctrl 多选）
    areasCount = getattr(mainRng.Areas, 'Count', 1)
    if areasCount > 1:
        return f"选区错误：选区 [{address}] 包含 {areasCount} 个离散块！\n本功能仅允许选择连续的一行或一列。"

    # 3. 获取 Excel 实例的极限行列数
    app = mainRng.Application
    maxRows = app.Rows.Count
    maxCols = app.Columns.Count

    # 4. 校验精确尺寸：必须是“整行且仅 1 行”或“整列且仅 1 列”
    isSingleFullRow = (mainRng.Columns.Count == maxCols) and (mainRng.Rows.Count == 1)
    isSingleFullCol = (mainRng.Rows.Count == maxRows) and (mainRng.Columns.Count == 1)

    if not isSingleFullRow and not isSingleFullCol:
        return f"选区尺寸错误：当前选区 [{address}] 不符合要求！\n必须且只能选择【单行的整行】（如点击行号 1）或【单列的整列】（如点击列标 A）。"

    return True

def processDelete(appComHandle, targets, logger):
    """核心业务逻辑函数"""
    
    # 既然代码能走到这里，说明 validator 已经 100% 确保了 targets 是单选区，且合法。
    # 我们直接取第 0 个元素即可
    sheetName, address, mainRng = targets[0]
    worksheet = mainRng.Worksheet
    totalDeleted = 0

    # 判断是行模式还是列模式（校验器已保证只会是其一）
    deleteMode = 'ROW' if mainRng.Rows.Count == 1 else 'COL'

    # 弹窗获取关键词
    targetKeywords = getUserInputKeywords()
    if not targetKeywords:
        logger.info("用户未输入有效的内容关键字或取消了输入，插件退出。")
        return

    logger.info(f"正在分析工作表 [{sheetName}] 的选区 [{address}]...")

    # 裁剪选区至 UsedRange，避免无用遍历（整行整列高达上百万个单元格，必须裁切）
    effectiveRng = appComHandle.Intersect(mainRng, worksheet.UsedRange)
    if effectiveRng is None:
        logger.info(f"工作表 [{sheetName}] 的选区内无有效数据，插件退出。")
        return

    if deleteMode == 'ROW':
        # ========== 选了单行整行 -> 检查选区内的单元格 -> 删除对应的【列】 ==========
        colsToDelete = set()
        for cell in effectiveRng.Cells:
            if checkMatch(cell.Value2, targetKeywords):
                colsToDelete.add(cell.Column)

        if not colsToDelete:
            logger.info(f"工作表 [{sheetName}] 的选中行内未找到匹配内容。")
            return

        # 从右往左（倒序）删除列，防止列索引偏移
        sortedCols = sorted(colsToDelete, reverse=True)
        for cIdx in sortedCols:
            worksheet.Columns(cIdx).Delete()
            totalDeleted += 1

        logger.info(f"处理完成！在选中行 [{address}] 中找到匹配内容，已成功删除 {totalDeleted} 列。")

    elif deleteMode == 'COL':
        # ========== 选了单列整列 -> 检查选区内的单元格 -> 删除对应的【行】 ==========
        rowsToDelete = set()
        for cell in effectiveRng.Cells:
            if checkMatch(cell.Value2, targetKeywords):
                rowsToDelete.add(cell.Row)

        if not rowsToDelete:
            logger.info(f"工作表 [{sheetName}] 的选中列内未找到匹配内容。")
            return

        # 从下往上（倒序）删除行，防止行索引偏移
        sortedRows = sorted(rowsToDelete, reverse=True)
        for rIdx in sortedRows:
            worksheet.Rows(rIdx).Delete()
            totalDeleted += 1

        logger.info(f"处理完成！在选中列 [{address}] 中找到匹配内容，已成功删除 {totalDeleted} 行。")

def run(appComHandle, logQueue):
    utils.runWithSelection(
        appComHandle, 
        logQueue, 
        pluginInfo, 
        pickerTitle="选单行删列 / 选单列删行（仅限选择单行整行或单列整列）", 
        actionFunc=processDelete,
        disableEvents=True,
        manualCalc=True,
        # 保持默认的 singleMode=False，避免提示语变成“提取单元格”产生歧义
        validator=checkSingleFullRowOrCol # <--- 全靠这里的 Area、Len、Count 进行降维打击
    )