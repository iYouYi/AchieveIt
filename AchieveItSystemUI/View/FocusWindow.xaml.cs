using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace AchieveItSystemUI
{
    /// <summary>
    /// FocusWindow.xaml 的交互逻辑
    /// </summary>
    public partial class FocusWindow : Window
    {
        FocusWindowViewModel _focusWindowViewModel;
        public FocusWindow()
        {
            _focusWindowViewModel = new FocusWindowViewModel();
            this.DataContext = _focusWindowViewModel;
            InitializeComponent();
            this.KeyDown += FocusWindow_KeyDown;
        }
        private void FocusWindow_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                var viewModel = DataContext as FocusWindowViewModel;
                viewModel?.ExitFocusCommand.Execute(this);
            }
        }
    }

}
