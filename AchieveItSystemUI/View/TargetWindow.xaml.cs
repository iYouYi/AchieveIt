using AchieveItSystemBLL;
using AchieveItSystemUI;
using System.Windows;

namespace AchieveItSystemUI
{
    public partial class TargetWindow : Window
    {
        public TargetWindow()
        {
            InitializeComponent();

            var excelService = new ExcelService();
            DataContext = new MainViewModel(excelService);
        }
    }
}