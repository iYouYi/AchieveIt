using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AchieveItSystemModels
{
    public class UserInfo : NotifyPropertyBase
    {
        private int _userId;
        private string _userCode;
        private string _userPassword;
        public int UserId
        {
            get { return _userId; }
            set { _userId = value; this.RaisePropertyChanged(); }
        }
        public string UserCode
        {
            get {return _userCode; }
            set { _userCode = value;this.RaisePropertyChanged() ; }
        }
        public string UserPassword
        {
            get {return _userPassword; }
            set { _userPassword = value;this.RaisePropertyChanged() ; }
        }
    }
}
