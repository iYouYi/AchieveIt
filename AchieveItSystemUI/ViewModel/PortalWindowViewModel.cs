using AchieveItSystemModels;
using AchieveItSystemUI.View;
using System.Diagnostics;
using System.Security.Policy;
using System.Windows;
using System.Windows.Threading;

namespace AchieveItSystemUI
{
    public class PortalWindowViewModel:NotifyPropertyBase
    {
        string testString = "今天也在慢慢发光呀，别急，时间会给你答案。@允许自己慢慢来， 每一步坚持都算数。@疲惫时就歇一歇，重新出发的你依然很棒。@Do your best—that's enough.";
        string _everydaySentence;
        public string EverydaySentence
        {
            get {
                ADailyWord aDailyWord = new ADailyWord(testString);
                Random random = new Random();
                int rndnum = random.Next(0, aDailyWord.ADailyWordLists.Length -1);
                return aDailyWord.ADailyWordLists[rndnum]; }
        }
        private DateTime _currentDateTime;
        public DateTime CurrentDateTime
        {
            get { return _currentDateTime; }
            set { _currentDateTime = value; this.RaisePropertyChanged(); }
        }
        public CommandBase TomatoClockCommand { get; set; }
        public CommandBase TossCoinCommand { get; set; }
        public CommandBase FocusModeCommand { get;set; }
        public CommandBase TargetCommand { get;set; }
        
        public PortalWindowViewModel()
        {
            //时间显示
            CurrentDateTime = DateTime.Now;
            var timer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(1)
            };
            timer.Tick += (o, e) =>
            {
                CurrentDateTime = DateTime.Now;
            };
            timer.Start();
            TomatoClockCommand = new CommandBase((o) =>
            {
                TomatoClockFunc(o);
            });
            TossCoinCommand = new CommandBase(TossCoinFunc);
            FocusModeCommand = new CommandBase(FocusModeFunc);
            TargetCommand = new CommandBase(TargetFunc);

        }

        private void TargetFunc(object o)
        {
            TargetWindow targetWindow = new TargetWindow();
            (o as Window).Close();
            targetWindow.Show();
        }

        public void TomatoClockFunc(object o)
        {
            PomodoroWindow pomodoroWindow = new PomodoroWindow();
            (o as Window).Close();
            pomodoroWindow.Show();
        }
        public void TossCoinFunc(object o)
        {
            CoinWindow coinWindow = new CoinWindow();
            (o as Window).Close();
            coinWindow.Show();
        }
        public void FocusModeFunc(object o)
        {
            FocusWindow focusWindow = new FocusWindow();
            (o as Window).Close();
            focusWindow.Show();
        }
    }
}