using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Collections.Generic;

namespace AchieveItSystemModels
{
    public class NotifyPropertyBase : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        // 某个属性发生变化自动通知UI该属性发生了变化
        public void RaisePropertyChanged([CallerMemberName] string propName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));
        }

        // 修改：返回bool值表示是否设置了新值
        public bool Set<T>(ref T field, T value, [CallerMemberName] string propName = "")
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
                return false;
            field = value;
            RaisePropertyChanged(propName);
            return true;
        }
    }
}