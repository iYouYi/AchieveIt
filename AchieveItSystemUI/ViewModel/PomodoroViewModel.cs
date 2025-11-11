using AchieveItSystemModels;
using AchieveItSystemUI.View;
using System;
using System.Collections.ObjectModel;
using System.Timers;
using System.Windows;
using Timer = System.Timers.Timer;

namespace AchieveItSystemUI
{
    public class PomodoroWindowViewModel : NotifyPropertyBase
    {
        private readonly Timer _timer;
        private PomodoroModel _pomodoro;

        public PomodoroModel Pomodoro
        {
            get => _pomodoro;
            set => Set(ref _pomodoro, value);
        }

        public ObservableCollection<QuickTimeOption> QuickTimeOptions { get; set; }

        public CommandBase StartCommand { get; }
        public CommandBase PauseCommand { get; }
        public CommandBase StopCommand { get; }
        public CommandBase ResetCommand { get; }
        public CommandBase SelectQuickTimeCommand { get; }
        public CommandBase TomatoClockCommand { get; set; }
        public CommandBase BackPortalCommand { get; set; }

        public PomodoroWindowViewModel()
        {
            Pomodoro = new PomodoroModel
            {
                Title = "专注工作",
                WorkDuration = 25,
                BreakDuration = 5,
                RemainingSeconds = 25 * 60,
                Status = PomodoroStatus.Stopped
            };

            InitializeQuickTimeOptions();

            _timer = new Timer(1000);
            _timer.Elapsed += Timer_Elapsed;

            StartCommand = new CommandBase(StartTimer, CanStartTimer);
            PauseCommand = new CommandBase(PauseTimer, CanPauseTimer);
            StopCommand = new CommandBase(StopTimer, CanStopTimer);
            ResetCommand = new CommandBase(ResetTimer);
            SelectQuickTimeCommand = new CommandBase(SelectQuickTime);
            BackPortalCommand = new CommandBase(BackPortalFunc);

            // 监听状态变化，更新命令可用性
            Pomodoro.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(Pomodoro.Status))
                {
                    StartCommand.RaiseCanExecuteChanged();
                    PauseCommand.RaiseCanExecuteChanged();
                    StopCommand.RaiseCanExecuteChanged();
                }
            };

            TomatoClockCommand = new CommandBase((o) =>
            {
                TomatoClockFunc(o);
            });
        }

        private void InitializeQuickTimeOptions()
        {
            QuickTimeOptions = new ObservableCollection<QuickTimeOption>
            {
                new QuickTimeOption { Name = "标准番茄", WorkMinutes = 25, BreakMinutes = 5 },
                new QuickTimeOption { Name = "短时专注", WorkMinutes = 15, BreakMinutes = 3 },
                new QuickTimeOption { Name = "深度工作", WorkMinutes = 50, BreakMinutes = 10 },
                new QuickTimeOption { Name = "高效冲刺", WorkMinutes = 45, BreakMinutes = 15 }
            };
        }

        private void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                if (Pomodoro.RemainingSeconds > 0)
                {
                    Pomodoro.RemainingSeconds--;
                }
                else
                {
                    _timer.Stop();
                    // 时间到，切换到休息或工作状态
                    if (Pomodoro.Status == PomodoroStatus.Running)
                    {
                        Pomodoro.Status = PomodoroStatus.Break;
                        Pomodoro.RemainingSeconds = Pomodoro.BreakDuration * 60;
                        MessageBox.Show("工作时间结束！开始休息。", "番茄钟", MessageBoxButton.OK, MessageBoxImage.Information);
                        _timer.Start();
                    }
                    else if (Pomodoro.Status == PomodoroStatus.Break)
                    {
                        Pomodoro.Status = PomodoroStatus.Stopped;
                        Pomodoro.RemainingSeconds = Pomodoro.WorkDuration * 60;
                        MessageBox.Show("休息时间结束！", "番茄钟", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                }
            });
        }

        private void StartTimer(object parameter)
        {
            if (Pomodoro.Status == PomodoroStatus.Stopped || Pomodoro.Status == PomodoroStatus.Break)
            {
                Pomodoro.StartTime = DateTime.Now;
                Pomodoro.RemainingSeconds = Pomodoro.WorkDuration * 60;
                Pomodoro.Status = PomodoroStatus.Running;
            }
            else if (Pomodoro.Status == PomodoroStatus.Paused)
            {
                Pomodoro.Status = PomodoroStatus.Running;
            }

            _timer.Start();

            // 更新命令状态
            StartCommand.RaiseCanExecuteChanged();
            PauseCommand.RaiseCanExecuteChanged();
            StopCommand.RaiseCanExecuteChanged();
        }

        private void PauseTimer(object parameter)
        {
            _timer.Stop();
            Pomodoro.Status = PomodoroStatus.Paused;

            // 更新命令状态
            StartCommand.RaiseCanExecuteChanged();
            PauseCommand.RaiseCanExecuteChanged();
            StopCommand.RaiseCanExecuteChanged();
        }

        private void StopTimer(object parameter)
        {
            _timer.Stop();
            Pomodoro.Status = PomodoroStatus.Stopped;
            Pomodoro.RemainingSeconds = Pomodoro.WorkDuration * 60;

            // 更新命令状态
            StartCommand.RaiseCanExecuteChanged();
            PauseCommand.RaiseCanExecuteChanged();
            StopCommand.RaiseCanExecuteChanged();
        }

        private void ResetTimer(object parameter)
        {
            _timer.Stop();
            Pomodoro.Status = PomodoroStatus.Stopped;
            Pomodoro.RemainingSeconds = Pomodoro.WorkDuration * 60;

            // 更新命令状态
            StartCommand.RaiseCanExecuteChanged();
            PauseCommand.RaiseCanExecuteChanged();
            StopCommand.RaiseCanExecuteChanged();
        }

        private void SelectQuickTime(object parameter)
        {
            if (parameter is QuickTimeOption option)
            {
                Pomodoro.WorkDuration = option.WorkMinutes;
                Pomodoro.BreakDuration = option.BreakMinutes;
                Pomodoro.RemainingSeconds = option.WorkMinutes * 60;

                if (Pomodoro.Status == PomodoroStatus.Stopped)
                {
                    ResetTimer(null);
                }
            }
        }

        private bool CanStartTimer(object parameter)
        {
            return Pomodoro.Status != PomodoroStatus.Running;
        }

        private bool CanPauseTimer(object parameter)
        {
            return Pomodoro.Status == PomodoroStatus.Running;
        }

        private bool CanStopTimer(object parameter)
        {
            return Pomodoro.Status != PomodoroStatus.Stopped;
        }
        public void TomatoClockFunc(object o)
        {
            PomodoroWindow pomodoroWindow = new PomodoroWindow();
            (o as Window).Close();
            pomodoroWindow.Show();
        }
        public void BackPortalFunc(object o)
        {
            PortalWindow portal = new PortalWindow();
            (o as Window).Close();
            portal.Show();
        }
    }
}