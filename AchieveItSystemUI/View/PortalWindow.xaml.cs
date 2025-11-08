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

namespace AchieveItSystemUI.View
{
    /// <summary>
    /// PortalWindow.xaml 的交互逻辑
    /// </summary>
    public partial class PortalWindow : Window
    {
        PortalWindowViewModel _portalWindowViewModel;
        public PortalWindow()
        {
            _portalWindowViewModel = new PortalWindowViewModel();
            this.DataContext = _portalWindowViewModel;
            InitializeComponent();
        }
    }
}
