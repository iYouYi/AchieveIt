using OfficeOpenXml;
using System;
using System.Data;
using System.IO;
using System.Linq;

namespace AchieveItSystemBLL
{
    public class ExcelService : IExcelService
    {
        public DataTable GetExcelData(string excelPath)
        {
            DataTable dataTable = new DataTable();

            using (var fileStream = new FileStream(excelPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var excelPackage = new ExcelPackage(fileStream))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0];
                if (worksheet.Dimension == null)
                {
                    return dataTable;
                }

                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                {
                    string colName = worksheet.Cells[1, col].Text.Trim();
                    if (string.IsNullOrEmpty(colName))
                    {
                        colName = $"列{col}";
                    }
                    dataTable.Columns.Add(colName);
                }

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Value;
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

        public void SaveDataToExcel(string excelPath, DataTable data, int selectedRowIndex, int completedColumnIndex, int completionDateColumnIndex, int overdueColumnIndex)
        {
            FileInfo fileInfo = new FileInfo(excelPath);
            using (var excelPackage = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0];
                int excelRow = selectedRowIndex + 2;

                int excelCol = completedColumnIndex + 1;
                worksheet.Cells[excelRow, excelCol].Value = data.Rows[selectedRowIndex][completedColumnIndex];
                worksheet.Cells[excelRow, excelCol].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                if (completionDateColumnIndex != -1)
                {
                    excelCol = completionDateColumnIndex + 1;
                    worksheet.Cells[excelRow, excelCol].Value = data.Rows[selectedRowIndex][completionDateColumnIndex];
                }

                if (overdueColumnIndex != -1)
                {
                    excelCol = overdueColumnIndex + 1;
                    worksheet.Cells[excelRow, excelCol].Value = data.Rows[selectedRowIndex][overdueColumnIndex];
                }

                excelPackage.Save();
            }
        }

        public void AddNewTaskToExcel(string excelPath, string date, string theme, string task, bool isCompleted)
        {
            FileInfo fileInfo = new FileInfo(excelPath);
            using (var excelPackage = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0];
                int newRow = worksheet.Dimension.End.Row + 1;

                string completedValue = isCompleted ? "是" : "否";
                string completionDateValue = isCompleted ? DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") : string.Empty;

                worksheet.Cells[newRow, 1].Value = date;
                worksheet.Cells[newRow, 2].Value = theme;
                worksheet.Cells[newRow, 3].Value = task;
                worksheet.Cells[newRow, 4].Value = completedValue;
                worksheet.Cells[newRow, 5].Value = completionDateValue;

                if (DateTime.TryParse(date, out DateTime taskDate))
                {
                    bool isOverdue = isCompleted ?
                        (DateTime.Now.Date > taskDate.Date) :
                        (DateTime.Today > taskDate);
                    worksheet.Cells[newRow, 6].Value = isOverdue ? "是" : "否";
                }
                else
                {
                    worksheet.Cells[newRow, 6].Value = "否";
                }

                excelPackage.Save();
            }
        }

        public void CreateExcelTemplate(string filePath)
        {
            using (var excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("任务管理");

                worksheet.Cells[1, 1].Value = "日期";
                worksheet.Cells[1, 2].Value = "核心主题";
                worksheet.Cells[1, 3].Value = "具体任务内容";
                worksheet.Cells[1, 4].Value = "是否完成";
                worksheet.Cells[1, 5].Value = "完成日期";
                worksheet.Cells[1, 6].Value = "是否逾期";

                using (var range = worksheet.Cells[1, 1, 1, 6])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                    range.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                }

                worksheet.Column(1).Width = 15;
                worksheet.Column(2).Width = 20;
                worksheet.Column(3).Width = 30;
                worksheet.Column(4).Width = 12;
                worksheet.Column(5).Width = 20;
                worksheet.Column(6).Width = 12;

                worksheet.Cells[2, 1].Value = DateTime.Now.ToString("yyyy-MM-dd");
                worksheet.Cells[2, 2].Value = "单词学习";
                worksheet.Cells[2, 3].Value = "背诵50个新单词";
                worksheet.Cells[2, 4].Value = "否";

                worksheet.Cells[3, 1].Value = DateTime.Now.AddDays(1).ToString("yyyy-MM-dd");
                worksheet.Cells[3, 2].Value = "语法练习";
                worksheet.Cells[3, 3].Value = "完成10个语法练习题";
                worksheet.Cells[3, 4].Value = "否";

                worksheet.Cells[4, 1].Value = DateTime.Now.AddDays(2).ToString("yyyy-MM-dd");
                worksheet.Cells[4, 2].Value = "阅读理解";
                worksheet.Cells[4, 3].Value = "阅读一篇英文文章并总结";
                worksheet.Cells[4, 4].Value = "否";

                FileInfo fileInfo = new FileInfo(filePath);
                excelPackage.SaveAs(fileInfo);
            }
        }
    }
}