using AchieveItSystemModels;
using AchieveItSystemUI.View;
using System.Security.Policy;
using System.Windows;
namespace AchieveItSystemUI
{
    public class LoginWindowViewModel :NotifyPropertyBase
    {
        UserInfo _userInfo;
        public UserInfo CurrentUserInfo
        {
            get { return _userInfo; }
            set { _userInfo = value; this.RaisePropertyChanged(); }
        }
        public CommandBase LoginCommand { get; set; }
        public CommandBase CancelCommand { get; set; }
        public LoginWindowViewModel()
        {
            CurrentUserInfo = new UserInfo() { UserCode = "YY", UserPassword = "123" };
            LoginCommand = new CommandBase((o =>
            {
                DoLogin(o);
            }));
            CancelCommand = new CommandBase((o) =>
            {
                DoCancel(o);
            });
        }

        public void DoLogin(object o)
        {
            string userCode = "YY";
            string userPassWord = "123";
            if (CurrentUserInfo.UserCode == userCode && CurrentUserInfo.UserPassword == userPassWord)
            {
                MessageBox.Show("登录成功！");
            }
            PortalWindow portal = new PortalWindow();
            (o as Window).Close();
            portal.Show();
        }

        public void DoCancel(object o)
        {
            (o as Window).Close();
        }

    }
}