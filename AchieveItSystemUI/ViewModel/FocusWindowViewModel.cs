using AchieveItSystemModels;
using AchieveItSystemUI.View;
using System;
using System.Timers;
using System.Windows;
using Timer = System.Timers.Timer;

namespace AchieveItSystemUI
{
    public class FocusWindowViewModel : NotifyPropertyBase
    {
        private FocusModel _focus;
        private Timer _timer;
        private string _windowTitle = "专注模式";

        public FocusModel Focus
        {
            get => _focus;
            set => Set(ref _focus, value);
        }

        public string WindowTitle
        {
            get => _windowTitle;
            set => Set(ref _windowTitle, value);
        }

        public CommandBase StartFocusCommand { get; }
        public CommandBase PauseResumeCommand { get; }
        public CommandBase ExitFocusCommand { get; }
        public CommandBase BackPortalCommand { get; }

        public FocusWindowViewModel()
        {
            Focus = new FocusModel();

            _timer = new Timer(1000);
            _timer.Elapsed += Timer_Elapsed;

            StartFocusCommand = new CommandBase(StartFocus);
            PauseResumeCommand = new CommandBase(PauseResumeFocus);
            ExitFocusCommand = new CommandBase(ExitFocus);
            BackPortalCommand = new CommandBase(BackPortalFunc);
        }

        private void StartFocus(object parameter)
        {
            if (!Focus.IsRunning)
            {
                Focus.IsRunning = true;
                Focus.StartTime = DateTime.Now - Focus.ElapsedTime;
                _timer.Start();

                Focus.IsFullScreen = true;

                if (parameter is Window window)
                {
                    EnterFullScreen(window);
                }
            }
        }

        private void PauseResumeFocus(object parameter)
        {
            if (Focus.IsRunning)
            {
                // 暂停
                Focus.IsRunning = false;
                _timer.Stop();
            }
            else
            {
                // 恢复
                Focus.IsRunning = true;
                Focus.StartTime = DateTime.Now - Focus.ElapsedTime;
                _timer.Start();
            }

            // 更新命令状态
            PauseResumeCommand.RaiseCanExecuteChanged();
        }

        private void ExitFocus(object parameter)
        {
            _timer.Stop();
            Focus.IsRunning = false;
            Focus.ElapsedTime = TimeSpan.Zero;
            Focus.IsFullScreen = false;

            if (parameter is Window window)
            {
                ExitFullScreen(window);
            }
        }

        private void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                Focus.ElapsedTime = DateTime.Now - Focus.StartTime;
            });
        }

        private void EnterFullScreen(Window window)
        {
            window.WindowState = WindowState.Maximized;
            window.WindowStyle = WindowStyle.None;
            window.Topmost = true;
        }

        private void ExitFullScreen(Window window)
        {
            window.WindowState = WindowState.Normal;
            window.WindowStyle = WindowStyle.SingleBorderWindow;
            window.Topmost = false;
        }

        public void BackPortalFunc(object o)
        {
            _timer?.Stop();
            PortalWindow portal = new PortalWindow();
            (o as Window).Close();
            portal.Show();
        }
    }
}