using System.ComponentModel;
using System.Runtime.CompilerServices;

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
        // 封装属性赋值的通用逻辑，自动比较新旧值并触发通知（避免手动写重复代码）
        public void Set<T>(ref T field, T value, [CallerMemberName] string propName = "")
        {
            if (EqualityComparer<T>.Default.Equals(field, value))
                return;
            field = value;
            RaisePropertyChanged(propName);
        }
    }
}