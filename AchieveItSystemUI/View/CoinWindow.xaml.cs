using System.Windows;

namespace AchieveItSystemUI
{
    public partial class CoinWindow : Window
    {
        CoinWindowViewModel _coinWindowViewModel;
        public CoinWindow()
        {
            _coinWindowViewModel = new CoinWindowViewModel();
            this.DataContext = _coinWindowViewModel;
            InitializeComponent();
        }
    }
}