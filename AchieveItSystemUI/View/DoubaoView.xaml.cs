using AchieveItSystemUI.ViewModels;
using Microsoft.Web.WebView2.Core;
using System.Windows;

namespace AchieveItSystemUI.View
{
    public partial class DoubaoView : Window
    {
        public DoubaoViewModel ViewModel => (DoubaoViewModel)DataContext;

        public DoubaoView()
        {
            InitializeComponent();
            Loaded += DoubaoView_Loaded;
        }

        private async void DoubaoView_Loaded(object sender, RoutedEventArgs e)
        {
            await DoubaoWebView.EnsureCoreWebView2Async();
        }

        private void DoubaoWebView_CoreWebView2InitializationCompleted(object sender, CoreWebView2InitializationCompletedEventArgs e)
        {
            if (e.IsSuccess && DoubaoWebView.CoreWebView2 != null)
            {
                ViewModel.CoreWebView = DoubaoWebView.CoreWebView2;
                // 修复事件处理程序匹配问题
                DoubaoWebView.CoreWebView2.DownloadStarting += (s, args) => ViewModel.HandleDownloadStarting(args);
            }
            else
            {
                MessageBox.Show("WebView2 初始化失败: " + e.InitializationException?.Message, "错误", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void DoubaoWebView_DragEnter(object sender, System.Windows.DragEventArgs e)
        {
            ViewModel.DragEnterCommand.Execute(e);
        }

        private void DoubaoWebView_Drop(object sender, System.Windows.DragEventArgs e)
        {
            ViewModel.DropCommand.Execute(e);
        }
    }
}