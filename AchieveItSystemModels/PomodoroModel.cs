using System;
using System.Collections.Generic;

namespace AchieveItSystemModels
{
    public class PomodoroModel : NotifyPropertyBase
    {
        private string _title = "专注工作";
        private int _workDuration = 25;
        private int _breakDuration = 5;
        private int _remainingSeconds;
        private PomodoroStatus _status = PomodoroStatus.Stopped;
        private DateTime _startTime;

        public string Title
        {
            get => _title;
            set => Set(ref _title, value);
        }

        public int WorkDuration
        {
            get => _workDuration;
            set
            {
                if (Set(ref _workDuration, value) && Status == PomodoroStatus.Stopped)
                {
                    RemainingSeconds = value * 60;
                }
            }
        }

        public int BreakDuration
        {
            get => _breakDuration;
            set => Set(ref _breakDuration, value);
        }

        public int RemainingSeconds
        {
            get => _remainingSeconds;
            set
            {
                if (Set(ref _remainingSeconds, value))
                {
                    RaisePropertyChanged(nameof(DisplayTime));
                }
            }
        }

        public string DisplayTime
        {
            get
            {
                var minutes = RemainingSeconds / 60;
                var seconds = RemainingSeconds % 60;
                return $"{minutes:D2}:{seconds:D2}";
            }
        }

        public PomodoroStatus Status
        {
            get => _status;
            set
            {
                if (Set(ref _status, value))
                {
                    RaisePropertyChanged(nameof(IsRunning));
                    RaisePropertyChanged(nameof(IsStopped));
                    RaisePropertyChanged(nameof(IsPaused));
                }
            }
        }

        public DateTime StartTime
        {
            get => _startTime;
            set => Set(ref _startTime, value);
        }

        public bool IsRunning => Status == PomodoroStatus.Running;
        public bool IsStopped => Status == PomodoroStatus.Stopped;
        public bool IsPaused => Status == PomodoroStatus.Paused;
    }

    public enum PomodoroStatus
    {
        Stopped,
        Running,
        Paused,
        Break
    }

    public class QuickTimeOption
    {
        public string Name { get; set; }
        public int WorkMinutes { get; set; }
        public int BreakMinutes { get; set; }
    }
}