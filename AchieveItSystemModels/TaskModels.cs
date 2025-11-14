using System;
using System.Data;

namespace AchieveItSystemModels
{
    public class TaskItem : NotifyPropertyBase
    {
        private string _date;
        private string _dayNumber;
        private string _theme;
        private string _taskContent;
        private string _isCompleted;
        private string _isOverdue;
        private string _statusColor; 
        private string _overdueColor;
        private DataRow _originalRow;

        public string 日期
        {
            get => _date;
            set => Set(ref _date, value);
        }

        public string 第几天
        {
            get => _dayNumber;
            set => Set(ref _dayNumber, value);
        }

        public string 核心主题
        {
            get => _theme;
            set => Set(ref _theme, value);
        }

        public string 具体任务内容
        {
            get => _taskContent;
            set => Set(ref _taskContent, value);
        }

        public string 是否完成
        {
            get => _isCompleted;
            set => Set(ref _isCompleted, value);
        }

        public string 是否逾期
        {
            get => _isOverdue;
            set => Set(ref _isOverdue, value);
        }

        public string 状态颜色
        {
            get => _statusColor;
            set => Set(ref _statusColor, value);
        }

        public string 逾期颜色
        {
            get => _overdueColor;
            set => Set(ref _overdueColor, value);
        }

        public DataRow OriginalRow
        {
            get => _originalRow;
            set => Set(ref _originalRow, value);
        }
    }

    public class TaskDetail : NotifyPropertyBase
    {
        private string _date;
        private string _theme;
        private string _task;
        private bool _isCompleted;
        private string _completionDate;
        private string _overdue;

        public string 日期
        {
            get => _date;
            set => Set(ref _date, value);
        }

        public string 核心主题
        {
            get => _theme;
            set => Set(ref _theme, value);
        }

        public string 具体任务
        {
            get => _task;
            set => Set(ref _task, value);
        }

        public bool 是否完成
        {
            get => _isCompleted;
            set => Set(ref _isCompleted, value);
        }

        public string 完成日期
        {
            get => _completionDate;
            set => Set(ref _completionDate, value);
        }

        public string 是否逾期
        {
            get => _overdue;
            set => Set(ref _overdue, value);
        }
    }
}