using AchieveItSystemModels;
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
        }
    }
}