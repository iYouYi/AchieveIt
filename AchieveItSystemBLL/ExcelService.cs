using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace AchieveItSystemBLL
{
    public class ExcelService : IExcelService
    {
        /// <summary>
        /// 读取Excel数据到DataTable
        /// </summary>
        /// <param name="excelPath">Excel文件路径</param>
        /// <returns>包含Excel数据的DataTable</returns>
        public DataTable GetExcelData(string excelPath)
        {
            // 校验文件是否存在
            if (!File.Exists(excelPath))
            {
                throw new FileNotFoundException("指定的Excel文件不存在", excelPath);
            }

            DataTable dataTable = new DataTable();

            using (var fileStream = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fileStream))
            {
                // 校验工作表是否存在
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    throw new Exception("Excel文件中未找到任何工作表，请检查文件有效性");
                }

                // 处理空工作表（仅创建标准表头）
                if (worksheet.Dimension == null)
                {
                    AddStandardColumns(dataTable);
                    return dataTable;
                }

                // 读取表头（处理重复列名）
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    string colName = worksheet.Cells[1, col].Text.Trim();
                    // 空列名默认命名
                    if (string.IsNullOrEmpty(colName))
                    {
                        colName = $"列{col}";
                    }
                    // 处理重复列名（避免DataTable列名冲突）
                    int duplicateCount = 1;
                    string originalColName = colName;
                    while (dataTable.Columns.Contains(colName))
                    {
                        colName = $"{originalColName}_{duplicateCount}";
                        duplicateCount++;
                    }
                    dataTable.Columns.Add(colName);
                }

                // 读取数据行（跳过表头行）
                if (worksheet.Dimension.End.Row < 2)
                {
                    // 仅含表头无数据，返回空数据DataTable
                    return dataTable;
                }

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Value;
                        // 统一日期格式为yyyy-MM-dd
                        if (cellValue is DateTime dateValue)
                        {
                            dataRow[col - 1] = dateValue.ToString("yyyy-MM-dd");
                        }
                        else
                        {
                            dataRow[col - 1] = cellValue ?? string.Empty;
                        }
                    }
                    dataTable.Rows.Add(dataRow);
                }
            }

            return dataTable;
        }

        /// <summary>
        /// 保存数据到Excel（支持单条更新和批量更新）
        /// </summary>
        /// <param name="excelPath">Excel文件路径</param>
        /// <param name="data">待保存的DataTable数据</param>
        /// <param name="selectedRowIndex">选中行索引（-1表示批量更新）</param>
        /// <param name="completedColumnIndex">“是否完成”列索引</param>
        /// <param name="completionDateColumnIndex">“完成日期”列索引</param>
        /// <param name="overdueColumnIndex">“是否逾期”列索引</param>
        public void SaveDataToExcel(string excelPath, DataTable data, int selectedRowIndex, int completedColumnIndex, int completionDateColumnIndex, int overdueColumnIndex)
        {
            // 1. 基础校验
            if (!File.Exists(excelPath))
            {
                throw new FileNotFoundException("指定的Excel文件不存在", excelPath);
            }
            if (data == null || data.Rows.Count == 0)
            {
                throw new ArgumentException("待保存的DataTable为空，无数据可保存", nameof(data));
            }

            // 2. 检测文件是否被占用（如Excel已打开）
            CheckFileOccupied(excelPath);

            FileInfo fileInfo = new FileInfo(excelPath);
            using (var excelPackage = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    throw new Exception("Excel文件中未找到任何工作表，请检查文件有效性");
                }

                // 3. 区分批量更新和单条更新
                if (selectedRowIndex == -1)
                {
                    // 批量更新：遍历所有行，仅更新状态列
                    BatchUpdateStatusColumns(worksheet, data, completedColumnIndex, completionDateColumnIndex, overdueColumnIndex);
                }
                else
                {
                    // 单条更新：仅更新指定行的状态列
                    SingleRowUpdateStatusColumns(worksheet, data, selectedRowIndex, completedColumnIndex, completionDateColumnIndex, overdueColumnIndex);
                }

                // 4. 保存修改
                excelPackage.Save();
            }
        }

        /// <summary>
        /// 新增任务到Excel
        /// </summary>
        /// <param name="excelPath">Excel文件路径</param>
        /// <param name="date">任务日期</param>
        /// <param name="theme">核心主题</param>
        /// <param name="task">具体任务内容</param>
        /// <param name="isCompleted">是否完成</param>
        public void AddNewTaskToExcel(string excelPath, string date, string theme, string task, bool isCompleted)
        {
            // 1. 基础校验
            if (!File.Exists(excelPath))
            {
                throw new FileNotFoundException("指定的Excel文件不存在", excelPath);
            }
            if (string.IsNullOrEmpty(date) || string.IsNullOrEmpty(theme) || string.IsNullOrEmpty(task))
            {
                throw new ArgumentException("任务日期、核心主题、具体任务内容不可为空", nameof(task));
            }

            // 2. 检测文件是否被占用
            CheckFileOccupied(excelPath);

            FileInfo fileInfo = new FileInfo(excelPath);
            using (var excelPackage = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null)
                {
                    throw new Exception("Excel文件中未找到任何工作表，请检查文件有效性");
                }

                // 3. 按列名获取列索引（避免硬编码，适配表格结构）
                Dictionary<string, int> columnMap = GetColumnIndexMap(worksheet);
                // 校验必需列是否存在
                ValidateRequiredColumns(columnMap);

                // 4. 计算新增行信息
                int newRow = GetNewRowNumber(worksheet); // 新增行号
                string dayNumber = CalculateDayNumber(worksheet, columnMap, date); // 计算“第几天”
                string completedValue = isCompleted ? "是" : "否";
                string completionDateValue = isCompleted ? DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") : string.Empty;
                string overdueValue = CalculateOverdueStatus(date, isCompleted); // 计算“是否逾期”

                // 5. 填充新增任务数据（按列名定位，无错位）
                worksheet.Cells[newRow, columnMap["日期"]].Value = date;
                worksheet.Cells[newRow, columnMap["第几天"]].Value = dayNumber;
                worksheet.Cells[newRow, columnMap["核心主题"]].Value = theme;
                worksheet.Cells[newRow, columnMap["具体任务内容"]].Value = task;
                worksheet.Cells[newRow, columnMap["是否完成"]].Value = completedValue;
                worksheet.Cells[newRow, columnMap["完成日期"]].Value = completionDateValue;
                worksheet.Cells[newRow, columnMap["是否逾期"]].Value = overdueValue;

                // 6. 优化列宽（保持显示一致性）
                SetColumnWidths(worksheet, columnMap);

                // 7. 保存修改
                excelPackage.Save();
            }
        }

        /// <summary>
        /// 创建任务管理Excel模板
        /// </summary>
        /// <param name="filePath">模板保存路径</param>
        public void CreateExcelTemplate(string filePath)
        {
            // 校验路径是否合法
            if (string.IsNullOrEmpty(filePath))
            {
                throw new ArgumentException("模板保存路径不可为空", nameof(filePath));
            }

            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("任务管理");

                // 1. 设置表头
                string[] headers = { "日期", "第几天", "核心主题", "具体任务内容", "是否完成", "完成日期", "是否逾期" };
                for (int col = 0; col < headers.Length; col++)
                {
                    worksheet.Cells[1, col + 1].Value = headers[col];
                }

                // 2. 设置表头样式（加粗、浅蓝色背景、居中）
                using (var headerRange = worksheet.Cells[1, 1, 1, headers.Length])
                {
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                    headerRange.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                }

                // 3. 设置列宽
                SetColumnWidths(worksheet, new Dictionary<string, int>
                {
                    { "日期", 15 },
                    { "第几天", 10 },
                    { "核心主题", 20 },
                    { "具体任务内容", 35 },
                    { "是否完成", 12 },
                    { "完成日期", 22 },
                    { "是否逾期", 12 }
                });

                // 4. 添加示例数据
                AddTemplateSampleData(worksheet);

                // 5. 保存模板
                FileInfo fileInfo = new FileInfo(filePath);
                excelPackage.SaveAs(fileInfo);
            }
        }


        #region 私有辅助方法（封装重复逻辑，提高可维护性）
        /// <summary>
        /// 为DataTable添加标准列（空工作表时使用）
        /// </summary>
        private void AddStandardColumns(DataTable dataTable)
        {
            dataTable.Columns.AddRange(new DataColumn[]
            {
                new DataColumn("日期"),
                new DataColumn("第几天"),
                new DataColumn("核心主题"),
                new DataColumn("具体任务内容"),
                new DataColumn("是否完成"),
                new DataColumn("完成日期"),
                new DataColumn("是否逾期")
            });
        }

        /// <summary>
        /// 检测文件是否被占用
        /// </summary>
        private void CheckFileOccupied(string filePath)
        {
            try
            {
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Write, FileShare.None))
                {
                    // 仅用于检测文件占用，无实际读写
                }
            }
            catch (IOException)
            {
                throw new Exception("Excel文件已被其他程序（如Microsoft Excel）打开，请关闭后重试");
            }
        }

        /// <summary>
        /// 批量更新状态列（是否完成、完成日期、是否逾期）
        /// </summary>
        private void BatchUpdateStatusColumns(ExcelWorksheet worksheet, DataTable data, int completedColIdx, int completionDateColIdx, int overdueColIdx)
        {
            for (int i = 0; i < data.Rows.Count; i++)
            {
                int excelRow = i + 2; // Excel行号从2开始（跳过表头）
                // 安全校验：避免行号超出工作表范围
                if (worksheet.Dimension != null && excelRow > worksheet.Dimension.End.Row)
                {
                    continue;
                }

                // 更新“是否完成”列
                UpdateSingleColumn(worksheet, excelRow, completedColIdx, data.Rows[i][completedColIdx], true);
                // 更新“完成日期”列
                UpdateSingleColumn(worksheet, excelRow, completionDateColIdx, data.Rows[i][completionDateColIdx], false);
                // 更新“是否逾期”列
                UpdateSingleColumn(worksheet, excelRow, overdueColIdx, data.Rows[i][overdueColIdx], false);
            }
        }

        /// <summary>
        /// 单条更新状态列
        /// </summary>
        private void SingleRowUpdateStatusColumns(ExcelWorksheet worksheet, DataTable data, int rowIdx, int completedColIdx, int completionDateColIdx, int overdueColIdx)
        {
            // 行索引校验
            if (rowIdx < 0 || rowIdx >= data.Rows.Count)
            {
                throw new ArgumentOutOfRangeException(nameof(rowIdx), "选中的行索引超出数据范围，无法更新");
            }

            int excelRow = rowIdx + 2;
            // 行号校验
            if (excelRow < 2 || (worksheet.Dimension != null && excelRow > worksheet.Dimension.End.Row))
            {
                throw new Exception($"选中行对应Excel行号{excelRow}无效，请检查数据");
            }

            // 更新状态列
            UpdateSingleColumn(worksheet, excelRow, completedColIdx, data.Rows[rowIdx][completedColIdx], true);
            UpdateSingleColumn(worksheet, excelRow, completionDateColIdx, data.Rows[rowIdx][completionDateColIdx], false);
            UpdateSingleColumn(worksheet, excelRow, overdueColIdx, data.Rows[rowIdx][overdueColIdx], false);
        }

        /// <summary>
        /// 更新单个列的数据
        /// </summary>
        private void UpdateSingleColumn(ExcelWorksheet worksheet, int excelRow, int colIdx, object value, bool isCenterAlign)
        {
            if (colIdx == -1 || colIdx >= worksheet.Dimension?.End.Column)
            {
                return; // 列索引无效，跳过
            }

            int excelCol = colIdx + 1; // DataTable列从0开始，Excel列从1开始
            worksheet.Cells[excelRow, excelCol].Value = value ?? string.Empty;
            // 仅“是否完成”列需要居中对齐
            if (isCenterAlign)
            {
                worksheet.Cells[excelRow, excelCol].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
            }
        }

        /// <summary>
        /// 获取列名-列索引映射（按表头行）
        /// </summary>
        private Dictionary<string, int> GetColumnIndexMap(ExcelWorksheet worksheet)
        {
            Dictionary<string, int> columnMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            if (worksheet.Dimension == null)
            {
                return columnMap;
            }

            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                string colName = worksheet.Cells[1, col].Text.Trim();
                if (!string.IsNullOrEmpty(colName) && !columnMap.ContainsKey(colName))
                {
                    columnMap[colName] = col;
                }
            }
            return columnMap;
        }

        /// <summary>
        /// 校验必需列是否存在
        /// </summary>
        private void ValidateRequiredColumns(Dictionary<string, int> columnMap)
        {
            List<string> requiredCols = new List<string> { "日期", "第几天", "核心主题", "具体任务内容", "是否完成", "完成日期", "是否逾期" };
            var missingCols = requiredCols.Where(col => !columnMap.ContainsKey(col)).ToList();

            if (missingCols.Count > 0)
            {
                throw new Exception($"Excel表格缺少必需列：{string.Join("、", missingCols)}，请使用标准模板创建表格");
            }
        }

        /// <summary>
        /// 获取新增任务的行号（最后一行+1）
        /// </summary>
        private int GetNewRowNumber(ExcelWorksheet worksheet)
        {
            return worksheet.Dimension == null ? 2 : worksheet.Dimension.End.Row + 1;
        }

        /// <summary>
        /// 计算“第几天”（基于最早任务日期）
        /// </summary>
        private string CalculateDayNumber(ExcelWorksheet worksheet, Dictionary<string, int> columnMap, string taskDateStr)
        {
            // 无法解析日期时默认“第1天”
            if (!DateTime.TryParse(taskDateStr, out DateTime taskDate))
            {
                return "第1天";
            }

            // 工作表无数据时默认“第1天”
            if (worksheet.Dimension == null || worksheet.Dimension.End.Row < 2)
            {
                return "第1天";
            }

            // 查找最早任务日期
            DateTime earliestDate = DateTime.MaxValue;
            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
            {
                string dateText = worksheet.Cells[row, columnMap["日期"]].Text.Trim();
                if (DateTime.TryParse(dateText, out DateTime currentDate) && currentDate < earliestDate)
                {
                    earliestDate = currentDate;
                }
            }

            // 计算天数差（+1是因为“第1天”开始计数）
            int dayDiff = (taskDate - earliestDate).Days + 1;
            return $"第{dayDiff}天";
        }

        /// <summary>
        /// 计算“是否逾期”状态
        /// </summary>
        private string CalculateOverdueStatus(string taskDateStr, bool isCompleted)
        {
            if (!DateTime.TryParse(taskDateStr, out DateTime taskDate))
            {
                return "否";
            }

            DateTime today = DateTime.Today;
            // 已完成：完成日期（当前时间）> 任务日期 → 逾期
            if (isCompleted)
            {
                return DateTime.Now.Date > taskDate.Date ? "是" : "否";
            }
            // 未完成：今天 > 任务日期 → 逾期
            else
            {
                return today > taskDate ? "是" : "否";
            }
        }

        /// <summary>
        /// 设置列宽（保持显示一致性）
        /// </summary>
        private void SetColumnWidths(ExcelWorksheet worksheet, Dictionary<string, int> columnWidthMap)
        {
            foreach (var item in columnWidthMap)
            {
                // 查找列索引（忽略大小写）
                int colIndex = worksheet.Cells[1, 1, 1, worksheet.Dimension?.End.Column ?? 7]
                    .FirstOrDefault(cell => string.Equals(cell.Text.Trim(), item.Key, StringComparison.OrdinalIgnoreCase))
                    ?.Start
                    .Column ?? 0;

                if (colIndex > 0)
                {
                    worksheet.Column(colIndex).Width = item.Value;
                }
            }
        }

        /// <summary>
        /// 为模板添加示例数据
        /// </summary>
        private void AddTemplateSampleData(ExcelWorksheet worksheet)
        {
            DateTime today = DateTime.Today;
            // 示例数据行（3行）
            var sampleData = new List<object[]>
            {
                new object[] { today.ToString("yyyy-MM-dd"), "第1天", "基础夯实期", "背诵四级核心词汇50个（含复习前1天词汇）", "否", "", "否" },
                new object[] { today.AddDays(1).ToString("yyyy-MM-dd"), "第2天", "基础夯实期", "学习语法专题：名词性从句", "否", "", "否" },
                new object[] { today.AddDays(2).ToString("yyyy-MM-dd"), "第3天", "基础夯实期", "听力入门训练15分钟（短对话10题）", "否", "", "否" }
            };

            // 填充示例数据
            for (int i = 0; i < sampleData.Count; i++)
            {
                int row = i + 2;
                for (int col = 0; col < sampleData[i].Length; col++)
                {
                    worksheet.Cells[row, col + 1].Value = sampleData[i][col];
                    // “是否完成”“是否逾期”列居中
                    if (col == 4 || col == 6)
                    {
                        worksheet.Cells[row, col + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                    }
                }
            }
        }
        #endregion
    }
}