using System.Windows;

namespace AchieveItSystemUI
{
    public partial class PomodoroWindow : Window
    {
        PomodoroWindowViewModel _pomodoroWindowViewModel;
        public PomodoroWindow()
        {
            _pomodoroWindowViewModel = new PomodoroWindowViewModel();
            this.DataContext = _pomodoroWindowViewModel;
            InitializeComponent();
        }
    }
}