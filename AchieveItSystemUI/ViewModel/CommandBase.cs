using System.Windows.Input;

namespace AchieveItSystemUI
{
    public class CommandBase : ICommand
    {
        // 当命令的可执行状态发生变化时触发的事件。
        public event EventHandler CanExecuteChanged;
        // 判断命令当前是否可以执行。
        public bool CanExecute(object parameter)
        {
            return DoCanExecute == null ? true : DoCanExecute.Invoke(null);
        }
        // 执行命令的核心逻辑
        public void Execute(object parameter)
        {
            DoExecute?.Invoke(parameter);
        }
        // 实际执行命令操作的委托。
        // 当命令被调用时（如按钮点击），执行此委托。
        public Action<object> DoExecute { get; set; }
        // 判断命令是否可执行的委托。
        public Func<object, bool> DoCanExecute { get; set; }

        public CommandBase() { }
        public CommandBase(Action<object> action)
        {
            DoExecute = action;
        }
        public CommandBase(Action<object> action, Func<object, bool> func) : this(action)
        {
            DoCanExecute = func;
        }
        public void RaiseCanExecuteChanged()
        {
            CanExecuteChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}