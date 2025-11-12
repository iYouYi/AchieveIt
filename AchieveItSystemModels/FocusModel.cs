using System;

namespace AchieveItSystemModels
{
    public class FocusModel : NotifyPropertyBase
    {
        private TimeSpan _elapsedTime;
        private bool _isRunning;
        private DateTime _startTime;
        private bool _isFullScreen;

        public TimeSpan ElapsedTime
        {
            get => _elapsedTime;
            set
            {
                if (Set(ref _elapsedTime, value))
                {
                    RaisePropertyChanged(nameof(DisplayTime));
                }
            }
        }

        public bool IsRunning
        {
            get => _isRunning;
            set => Set(ref _isRunning, value);
        }

        public DateTime StartTime
        {
            get => _startTime;
            set => Set(ref _startTime, value);
        }

        public bool IsFullScreen
        {
            get => _isFullScreen;
            set => Set(ref _isFullScreen, value);
        }

        public string DisplayTime => ElapsedTime.ToString(@"hh\:mm\:ss");

        public FocusModel()
        {
            ElapsedTime = TimeSpan.Zero;
            IsRunning = false;
            IsFullScreen = false;
        }
    }
}