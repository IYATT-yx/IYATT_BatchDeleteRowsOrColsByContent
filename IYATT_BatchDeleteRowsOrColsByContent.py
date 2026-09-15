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

def showWarningMessage(msg):
    """弹出警告提示框"""
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    messagebox.showwarning("选区不符合要求", msg, parent=root)
    root.destroy()

def validateSingleRowOrCol(appComHandle, targets, logger):
    """
    输入前的前置强校验：只允许【单行整行】或【单列整列】
    返回: (isValid, deleteMode, targetItem)
    deleteMode: 'ROW' 或 'COL'
    """
    # 1. 校验提取列表中是否有且仅有 1 个选区项
    if len(targets) != 1:
        msg = f"选区错误：当前提取了 {len(targets)} 个选区！\n本功能仅允许选择【单行的整行】或【单列的整列】。"
        logger.info(f"选区校验失败：提取列表中包含了 {len(targets)} 个选区，已中断。")
        showWarningMessage(msg)
        return False, None, None

    sheetName, address, mainRng = targets[0]

    # 2. 校验是否包含了非连续的多块区域（如按 Ctrl 多选）
    areasCount = getattr(mainRng.Areas, 'Count', 1)
    if areasCount > 1:
        msg = f"选区错误：选区 [{address}] 包含多个离散块！\n本功能仅允许选择连续的【单行的整行】或【单列的整列】。"
        logger.info(f"选区校验失败：选区包含 {areasCount} 个离散块，已中断。")
        showWarningMessage(msg)
        return False, None, None

    maxRows = appComHandle.Rows.Count
    maxCols = appComHandle.Columns.Count

    # 3. 校验精确尺寸：必须是“整行且仅 1 行”或“整列且仅 1 列”
    isSingleFullRow = (mainRng.Columns.Count == maxCols) and (mainRng.Rows.Count == 1)
    isSingleFullCol = (mainRng.Rows.Count == maxRows) and (mainRng.Columns.Count == 1)

    if not isSingleFullRow and not isSingleFullCol:
        msg = f"选区错误：当前选区 [{address}] 不符合要求！\n必须且只能选择【单行的整行】（如第 1 行）或【单列的整列】（如 A 列）。"
        logger.info(f"选区校验失败：选区 [{address}] 行数为 {mainRng.Rows.Count}，列数为 {mainRng.Columns.Count}，不符合单行整行或单列整列条件。")
        showWarningMessage(msg)
        return False, None, None

    deleteMode = 'ROW' if isSingleFullRow else 'COL'
    return True, deleteMode, targets[0]

def processDelete(appComHandle, targets, logger):
    """核心业务逻辑函数"""
    # ================= 1. 强校验：是否为“单行整行”或“单列整列” =================
    isValid, deleteMode, targetItem = validateSingleRowOrCol(appComHandle, targets, logger)
    if not isValid:
        return

    # ================= 2. 校验彻底通过后，才弹窗获取关键词 =================
    targetKeywords = getUserInputKeywords()
    if not targetKeywords:
        logger.info("用户未输入有效的内容关键字或取消了输入，插件退出。")
        return

    sheetName, address, mainRng = targetItem
    worksheet = mainRng.Worksheet
    totalDeleted = 0

    logger.info(f"正在分析工作表 [{sheetName}] 的选区 [{address}]...")

    # 裁剪选区至 UsedRange，避免无用遍历
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
        manualCalc=True
    )