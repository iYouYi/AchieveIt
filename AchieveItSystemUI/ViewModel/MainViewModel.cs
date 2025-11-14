using AchieveItSystemBLL;
using AchieveItSystemModels;
using AchieveItSystemUI.View;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Media;

namespace AchieveItSystemUI
{
    public class MainViewModel : NotifyPropertyBase
    {
        private readonly IExcelService _excelService;

        private string _excelPath = string.Empty;
        private DataTable _originalData;
        private int _selectedRowIndex = -1;
        private int _dateColumnIndex = -1;
        private int _themeColumnIndex = -1;
        private int _taskColumnIndex = -1;
        private int _completedColumnIndex = -1;
        private int _completionDateColumnIndex = -1;
        private int _overdueColumnIndex = -1;
        private int _dayNumberColumnIndex = -1;
        private const double ProgressBarMaxWidth = 400;
        private bool _isAddingNew = false;

        private ObservableCollection<TaskItem> _tasks;
        private TaskDetail _selectedTaskDetail;
        private double _progressPercent;
        private string _progressDescription;
        private string _filePathText;
        private bool _isDetailVisible;
        private TaskItem _selectedTask;

        public ObservableCollection<TaskItem> Tasks
        {
            get => _tasks;
            set => Set(ref _tasks, value);
        }

        public TaskDetail SelectedTaskDetail
        {
            get => _selectedTaskDetail;
            set => Set(ref _selectedTaskDetail, value);
        }

        public TaskItem SelectedTask
        {
            get => _selectedTask;
            set
            {
                if (Set(ref _selectedTask, value))
                {
                    SelectTask(value);
                }
            }
        }

        public double ProgressPercent
        {
            get => _progressPercent;
            set => Set(ref _progressPercent, value);
        }

        public string ProgressDescription
        {
            get => _progressDescription;
            set => Set(ref _progressDescription, value);
        }

        public string FilePathText
        {
            get => _filePathText;
            set => Set(ref _filePathText, value);
        }

        public bool IsDetailVisible
        {
            get => _isDetailVisible;
            set => Set(ref _isDetailVisible, value);
        }

        public CommandBase SelectFileCommand { get; }
        public CommandBase DownloadTemplateCommand { get; }
        public CommandBase OpenExcelCommand { get; }
        public CommandBase AddNewCommand { get; }
        public CommandBase RefreshCommand { get; }
        public CommandBase SubmitCommand { get; }
        public CommandBase BackPortalCommand { get; set; }

        public MainViewModel(IExcelService excelService)
        {
            _excelService = excelService;

            ExcelPackage.License.SetNonCommercialPersonal("YYY");

            Tasks = new ObservableCollection<TaskItem>();
            SelectedTaskDetail = new TaskDetail();
            FilePathText = "未选择文件";
            ProgressDescription = "未加载数据";

            SelectFileCommand = new CommandBase(SelectFile);
            DownloadTemplateCommand = new CommandBase(DownloadTemplate);
            OpenExcelCommand = new CommandBase(OpenExcel);
            AddNewCommand = new CommandBase(AddNew);
            RefreshCommand = new CommandBase(Refresh);
            SubmitCommand = new CommandBase(Submit);
            BackPortalCommand = new CommandBase(BackPortalFunc);

            UpdateProgressBar(0);
        }

        public void BackPortalFunc(object o)
        {
            PortalWindow portal = new PortalWindow();
            (o as Window).Close();
            portal.Show();
        }
        private void SelectFile(object parameter)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "Excel文件 (*.xlsx)|*.xlsx",
                Title = "选择Excel文件",
                Multiselect = false,
                RestoreDirectory = true
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _excelPath = openFileDialog.FileName;
                FilePathText = _excelPath;
                Properties.Settings.Default.LastExcelPath = _excelPath;
                Properties.Settings.Default.Save();
                ReadExcelAndShow();
                ResetDetailArea();
            }
        }

        private void DownloadTemplate(object parameter)
        {
            try
            {
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel文件 (*.xlsx)|*.xlsx",
                    Title = "保存Excel模板",
                    FileName = "任务管理模板.xlsx",
                    RestoreDirectory = true
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    string filePath = saveFileDialog.FileName;
                    _excelService.CreateExcelTemplate(filePath);
                    MessageBox.Show($"模板已成功保存到：\n{filePath}", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"下载模板失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void OpenExcel(object parameter)
        {
            if (string.IsNullOrEmpty(_excelPath) || !File.Exists(_excelPath))
            {
                MessageBox.Show("请先选择有效的Excel文件！", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                var result = MessageBox.Show(
                    "即将打开Excel文件，请在Excel中完成编辑后保存并关闭文件，然后回到本程序点击\"刷新\"按钮更新数据。\n\n注意：在Excel打开期间请不要在本程序中进行修改操作！",
                    "提示",
                    MessageBoxButton.OKCancel,
                    MessageBoxImage.Information);

                if (result == MessageBoxResult.OK)
                {
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = _excelPath,
                        UseShellExecute = true
                    });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开Excel文件失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void AddNew(object parameter)
        {
            if (string.IsNullOrEmpty(_excelPath) || !File.Exists(_excelPath))
            {
                MessageBox.Show("请先选择有效的Excel文件！", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            _isAddingNew = true;
            _selectedRowIndex = -1;

            SelectedTaskDetail = new TaskDetail
            {
                日期 = DateTime.Now.ToString("yyyy-MM-dd"),
                核心主题 = string.Empty,
                具体任务 = string.Empty,
                是否完成 = false,
                完成日期 = string.Empty,
                是否逾期 = "否"
            };

            IsDetailVisible = true;
            SelectedTask = null;
        }

        private void Refresh(object parameter)
        {
            if (string.IsNullOrEmpty(_excelPath) || !File.Exists(_excelPath))
            {
                MessageBox.Show("请先选择有效的Excel文件！", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                ResetDetailArea();
                ReadExcelAndShow();
                MessageBox.Show("数据已刷新！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"刷新数据失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void Submit(object parameter)
        {
            if (_isAddingNew)
            {
                SaveNewTask();
            }
            else
            {
                UpdateExistingTask();
            }
        }

        public void SelectTask(TaskItem selectedTask)
        {
            if (_isAddingNew)
            {
                _isAddingNew = false;
            }

            if (selectedTask == null)
            {
                ResetDetailArea();
                return;
            }

            if (_originalData == null || _dateColumnIndex == -1 || _themeColumnIndex == -1 ||
                _taskColumnIndex == -1 || _completedColumnIndex == -1)
            {
                ResetDetailArea();
                return;
            }

            try
            {
                DataRow selectedRow = selectedTask.OriginalRow;
                _selectedRowIndex = _originalData.Rows.IndexOf(selectedRow);

                SelectedTaskDetail = new TaskDetail
                {
                    日期 = selectedRow[_dateColumnIndex]?.ToString() ?? string.Empty,
                    核心主题 = selectedRow[_themeColumnIndex]?.ToString() ?? string.Empty,
                    具体任务 = selectedRow[_taskColumnIndex]?.ToString() ?? string.Empty,
                    是否完成 = string.Equals(selectedRow[_completedColumnIndex]?.ToString()?.Trim(), "是", StringComparison.OrdinalIgnoreCase),
                    完成日期 = _completionDateColumnIndex != -1 ? selectedRow[_completionDateColumnIndex]?.ToString() ?? string.Empty : string.Empty,
                    是否逾期 = _overdueColumnIndex != -1 ? selectedRow[_overdueColumnIndex]?.ToString() ?? string.Empty : string.Empty
                };

                IsDetailVisible = true;
            }
            catch (Exception ex)
            {
                ResetDetailArea();
                MessageBox.Show($"显示详细信息时出错：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ReadExcelAndShow()
        {
            try
            {
                if (!File.Exists(_excelPath))
                {
                    MessageBox.Show("所选Excel文件不存在！", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                _originalData = _excelService.GetExcelData(_excelPath);
                FindColumnIndexes();

                FillEmptyStatusColumns();

                CalculateOverdueStatus();

                if (_originalData.Rows.Count > 0)
                {
                    CreateDisplayData();
                    CalculateAndUpdateProgress();
                    CheckAndShowOverdueTasks();
                }
                else
                {
                    Tasks.Clear();
                    UpdateProgressBar(0);
                    ProgressDescription = "Excel文件为空";
                }
            }
            catch (Exception ex)
            {
                Tasks.Clear();
                UpdateProgressBar(0);
                ProgressDescription = "读取失败";
                MessageBox.Show($"读取失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void FillEmptyStatusColumns()
        {
            if (_originalData == null) return;

            bool hasChanges = false;

            if (_completedColumnIndex != -1)
            {
                foreach (DataRow row in _originalData.Rows)
                {
                    string completedValue = row[_completedColumnIndex]?.ToString()?.Trim() ?? "";
                    if (string.IsNullOrEmpty(completedValue))
                    {
                        string completionDate = "";
                        if (_completionDateColumnIndex != -1)
                        {
                            completionDate = row[_completionDateColumnIndex]?.ToString()?.Trim() ?? "";
                        }

                        if (!string.IsNullOrEmpty(completionDate))
                        {
                            row[_completedColumnIndex] = "是";
                        }
                        else
                        {
                            row[_completedColumnIndex] = "否";
                        }
                        hasChanges = true;
                    }
                }
            }

            if (_overdueColumnIndex != -1)
            {
                foreach (DataRow row in _originalData.Rows)
                {
                    string overdueValue = row[_overdueColumnIndex]?.ToString()?.Trim() ?? "";
                    if (string.IsNullOrEmpty(overdueValue))
                    {
                        string taskDateValue = row[_dateColumnIndex]?.ToString()?.Trim() ?? "";
                        string completedValue = row[_completedColumnIndex]?.ToString()?.Trim() ?? "";

                        if (DateTime.TryParse(taskDateValue, out DateTime taskDate))
                        {
                            if (completedValue == "是")
                            {
                                string completionDateValue = "";
                                if (_completionDateColumnIndex != -1)
                                {
                                    completionDateValue = row[_completionDateColumnIndex]?.ToString()?.Trim() ?? "";
                                }

                                if (!string.IsNullOrEmpty(completionDateValue) &&
                                    DateTime.TryParse(completionDateValue, out DateTime completionDate))
                                {
                                    row[_overdueColumnIndex] = completionDate.Date > taskDate.Date ? "是" : "否";
                                }
                                else
                                {
                                    row[_overdueColumnIndex] = "否";
                                }
                            }
                            else
                            {
                                row[_overdueColumnIndex] = DateTime.Today > taskDate ? "是" : "否";
                            }
                        }
                        else
                        {
                            row[_overdueColumnIndex] = "否";
                        }
                        hasChanges = true;
                    }
                }
            }

            if (_dayNumberColumnIndex != -1)
            {
                var sortedRows = _originalData.AsEnumerable()
                    .Where(row => !string.IsNullOrEmpty(row[_dateColumnIndex]?.ToString()))
                    .OrderBy(row =>
                    {
                        string dateStr = row[_dateColumnIndex]?.ToString() ?? "";
                        if (DateTime.TryParse(dateStr, out DateTime date))
                            return date;
                        return DateTime.MaxValue;
                    })
                    .ToList();

                if (sortedRows.Count > 0)
                {
                    DateTime? firstDate = null;
                    foreach (DataRow row in sortedRows)
                    {
                        string dateStr = row[_dateColumnIndex]?.ToString() ?? "";
                        if (DateTime.TryParse(dateStr, out DateTime currentDate))
                        {
                            if (firstDate == null || currentDate < firstDate.Value)
                                firstDate = currentDate;
                        }
                    }

                    if (firstDate.HasValue)
                    {
                        foreach (DataRow row in sortedRows)
                        {
                            string dayNumberValue = row[_dayNumberColumnIndex]?.ToString()?.Trim() ?? "";
                            if (string.IsNullOrEmpty(dayNumberValue))
                            {
                                string dateStr = row[_dateColumnIndex]?.ToString() ?? "";
                                if (DateTime.TryParse(dateStr, out DateTime currentDate))
                                {
                                    int dayNumber = (currentDate - firstDate.Value).Days + 1;
                                    row[_dayNumberColumnIndex] = $"第{dayNumber}天";
                                    hasChanges = true;
                                }
                            }
                        }
                    }
                }
            }

            if (hasChanges)
            {
                SaveDataTableToExcel();
            }
        }

        private void SaveDataTableToExcel()
        {
            try
            {
                _excelService.SaveDataToExcel(_excelPath, _originalData, -1, _completedColumnIndex, _completionDateColumnIndex, _overdueColumnIndex);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存自动填充的状态列时出错：{ex.Message}", "警告", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void CreateDisplayData()
        {
            Tasks.Clear();

            if (_originalData == null || _originalData.Rows.Count == 0)
                return;

            var sortedRows = _originalData.AsEnumerable()
                .OrderBy(row =>
                {
                    string dateStr = row[_dateColumnIndex]?.ToString() ?? "";
                    if (DateTime.TryParse(dateStr, out DateTime date))
                        return date;
                    return DateTime.MaxValue;
                })
                .ToList();

            DateTime? firstDate = null;
            for (int i = 0; i < sortedRows.Count; i++)
            {
                DataRow row = sortedRows[i];
                string dateStr = row[_dateColumnIndex]?.ToString() ?? "";

                if (DateTime.TryParse(dateStr, out DateTime currentDate))
                {
                    if (firstDate == null)
                        firstDate = currentDate;

                    int dayNumber = (currentDate - firstDate.Value).Days + 1;

                    string completedValue = row[_completedColumnIndex]?.ToString()?.Trim() ?? "";
                    string overdueValue = row[_overdueColumnIndex]?.ToString()?.Trim() ?? "";
                    string dayNumberValue = "";

                    if (_dayNumberColumnIndex != -1)
                    {
                        dayNumberValue = row[_dayNumberColumnIndex]?.ToString()?.Trim() ?? "";
                    }

                    if (string.IsNullOrEmpty(dayNumberValue))
                    {
                        dayNumberValue = $"第{dayNumber}天";
                    }

                    Tasks.Add(new TaskItem
                    {
                        日期 = dateStr,
                        第几天 = dayNumberValue,
                        核心主题 = row[_themeColumnIndex]?.ToString() ?? "",
                        具体任务内容 = row[_taskColumnIndex]?.ToString() ?? "",
                        是否完成 = completedValue,
                        是否逾期 = overdueValue,
                        状态颜色 = completedValue == "是" ? "Green" : "Red", // 使用颜色名称
                        逾期颜色 = overdueValue == "是" ? "Red" : "Green", // 使用颜色名称
                        OriginalRow = row
                    });
                }
            }
        }

        private void CheckAndShowOverdueTasks()
        {
            if (_originalData == null || _overdueColumnIndex == -1 || _dateColumnIndex == -1 || _themeColumnIndex == -1 || _completedColumnIndex == -1)
                return;

            List<string> overdueTasks = new List<string>();

            foreach (DataRow row in _originalData.Rows)
            {
                string overdueValue = row[_overdueColumnIndex]?.ToString()?.Trim() ?? string.Empty;
                string completedValue = row[_completedColumnIndex]?.ToString()?.Trim() ?? string.Empty;

                if (string.Equals(overdueValue, "是", StringComparison.OrdinalIgnoreCase) &&
                    !string.Equals(completedValue, "是", StringComparison.OrdinalIgnoreCase))
                {
                    string taskDate = row[_dateColumnIndex]?.ToString() ?? string.Empty;
                    string theme = row[_themeColumnIndex]?.ToString() ?? string.Empty;
                    overdueTasks.Add($"完成日期：{taskDate}，核心主题：{theme}");
                }
            }

            if (overdueTasks.Count > 0)
            {
                StringBuilder message = new StringBuilder();
                message.AppendLine($"发现 {overdueTasks.Count} 个逾期任务：");
                message.AppendLine();

                for (int i = 0; i < overdueTasks.Count; i++)
                {
                    message.AppendLine($"{i + 1}. {overdueTasks[i]}");
                }

                MessageBox.Show(message.ToString(), "逾期任务提醒", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void FindColumnIndexes()
        {
            _dateColumnIndex = -1;
            _themeColumnIndex = -1;
            _taskColumnIndex = -1;
            _completedColumnIndex = -1;
            _completionDateColumnIndex = -1;
            _overdueColumnIndex = -1;
            _dayNumberColumnIndex = -1;

            for (int i = 0; i < _originalData.Columns.Count; i++)
            {
                string colName = _originalData.Columns[i].ColumnName.Trim();
                if (colName == "日期") _dateColumnIndex = i;
                else if (colName == "核心主题") _themeColumnIndex = i;
                else if (colName == "具体任务内容") _taskColumnIndex = i;
                else if (colName == "是否完成") _completedColumnIndex = i;
                else if (colName == "完成日期") _completionDateColumnIndex = i;
                else if (colName == "是否逾期") _overdueColumnIndex = i;
                else if (colName == "第几天") _dayNumberColumnIndex = i;
            }
        }

        private void CalculateOverdueStatus()
        {
            if (_dateColumnIndex == -1 || _completedColumnIndex == -1 || _overdueColumnIndex == -1)
                return;

            DateTime today = DateTime.Today;

            foreach (DataRow row in _originalData.Rows)
            {
                string completedValue = row[_completedColumnIndex]?.ToString()?.Trim() ?? string.Empty;
                string completionDateValue = _completionDateColumnIndex != -1 ? row[_completionDateColumnIndex]?.ToString()?.Trim() ?? string.Empty : string.Empty;
                string taskDateValue = row[_dateColumnIndex]?.ToString()?.Trim() ?? string.Empty;

                if (string.Equals(completedValue, "是", StringComparison.OrdinalIgnoreCase))
                {
                    if (!string.IsNullOrEmpty(completionDateValue) &&
                        DateTime.TryParse(completionDateValue, out DateTime completionDate) &&
                        DateTime.TryParse(taskDateValue, out DateTime taskDate))
                    {
                        row[_overdueColumnIndex] = completionDate.Date > taskDate.Date ? "是" : "否";
                    }
                    else
                    {
                        row[_overdueColumnIndex] = "否";
                    }
                }
                else
                {
                    if (DateTime.TryParse(taskDateValue, out DateTime taskDate))
                    {
                        row[_overdueColumnIndex] = today > taskDate ? "是" : "否";
                    }
                    else
                    {
                        row[_overdueColumnIndex] = "否";
                    }
                }
            }
        }

        private void CalculateAndUpdateProgress()
        {
            if (_originalData == null || _originalData.Rows.Count == 0 || _completedColumnIndex == -1)
            {
                UpdateProgressBar(0);
                ProgressDescription = "无有效数据";
                return;
            }

            int totalCount = _originalData.Rows.Count;
            int completedCount = 0;

            foreach (DataRow row in _originalData.Rows)
            {
                string completedValue = row[_completedColumnIndex]?.ToString()?.Trim() ?? string.Empty;
                if (string.Equals(completedValue, "是", StringComparison.OrdinalIgnoreCase))
                {
                    completedCount++;
                }
            }

            double progressPercent = (double)completedCount / totalCount * 100;
            UpdateProgressBar(progressPercent);
            ProgressDescription = $"已完成{completedCount}/{totalCount}";
        }

        private void UpdateProgressBar(double progressPercent)
        {
            progressPercent = Math.Clamp(progressPercent, 0, 100);
            ProgressPercent = progressPercent;
        }

        private void ResetDetailArea()
        {
            SelectedTaskDetail = new TaskDetail();
            _selectedRowIndex = -1;
            _isAddingNew = false;
            IsDetailVisible = false;
            SelectedTask = null;
        }

        private void SaveNewTask()
        {
            if (string.IsNullOrEmpty(SelectedTaskDetail.日期) || string.IsNullOrEmpty(SelectedTaskDetail.核心主题) || string.IsNullOrEmpty(SelectedTaskDetail.具体任务))
            {
                MessageBox.Show("请填写任务日期、核心主题和具体任务内容！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (string.IsNullOrEmpty(_excelPath) || !File.Exists(_excelPath))
            {
                MessageBox.Show("Excel文件路径无效！", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                _excelService.AddNewTaskToExcel(_excelPath, SelectedTaskDetail.日期, SelectedTaskDetail.核心主题, SelectedTaskDetail.具体任务, SelectedTaskDetail.是否完成);

                ReadExcelAndShow();
                ResetDetailArea();

                MessageBox.Show("新目标已成功添加到Excel！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"保存新目标失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void UpdateExistingTask()
        {
            if (_selectedRowIndex < 0 || _selectedRowIndex >= _originalData.Rows.Count)
            {
                MessageBox.Show("请先选择一行数据！", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }
            if (_completedColumnIndex == -1)
            {
                MessageBox.Show("Excel中未找到\"是否完成\"列，无法提交修改！", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            if (string.IsNullOrEmpty(_excelPath) || !File.Exists(_excelPath))
            {
                MessageBox.Show("Excel文件路径无效！", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            try
            {
                bool isCompleted = SelectedTaskDetail.是否完成;
                string completedValue = isCompleted ? "是" : "否";
                _originalData.Rows[_selectedRowIndex][_completedColumnIndex] = completedValue;

                if (_completionDateColumnIndex != -1)
                {
                    if (isCompleted)
                    {
                        _originalData.Rows[_selectedRowIndex][_completionDateColumnIndex] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                    }
                    else
                    {
                        _originalData.Rows[_selectedRowIndex][_completionDateColumnIndex] = string.Empty;
                    }
                }

                if (_overdueColumnIndex != -1)
                {
                    if (isCompleted)
                    {
                        string completionDateValue = _originalData.Rows[_selectedRowIndex][_completionDateColumnIndex]?.ToString()?.Trim() ?? string.Empty;
                        string taskDateValue = _originalData.Rows[_selectedRowIndex][_dateColumnIndex]?.ToString()?.Trim() ?? string.Empty;

                        if (!string.IsNullOrEmpty(completionDateValue) &&
                            DateTime.TryParse(completionDateValue, out DateTime completionDate) &&
                            DateTime.TryParse(taskDateValue, out DateTime taskDate))
                        {
                            _originalData.Rows[_selectedRowIndex][_overdueColumnIndex] = completionDate.Date > taskDate.Date ? "是" : "否";
                        }
                        else
                        {
                            _originalData.Rows[_selectedRowIndex][_overdueColumnIndex] = "否";
                        }
                    }
                    else
                    {
                        string taskDateValue = _originalData.Rows[_selectedRowIndex][_dateColumnIndex]?.ToString()?.Trim() ?? string.Empty;
                        if (DateTime.TryParse(taskDateValue, out DateTime taskDate))
                        {
                            _originalData.Rows[_selectedRowIndex][_overdueColumnIndex] = DateTime.Today > taskDate ? "是" : "否";
                        }
                        else
                        {
                            _originalData.Rows[_selectedRowIndex][_overdueColumnIndex] = "否";
                        }
                    }
                }

                _excelService.SaveDataToExcel(_excelPath, _originalData, _selectedRowIndex, _completedColumnIndex, _completionDateColumnIndex, _overdueColumnIndex);

                CreateDisplayData();
                CalculateAndUpdateProgress();

                MessageBox.Show("修改已成功提交到Excel！", "成功", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"提交失败：{ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}