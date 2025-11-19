using AchieveItSystemModels;
using AchieveItSystemUI.View;
using Microsoft.Web.WebView2.Core;
using Microsoft.Win32;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace AchieveItSystemUI.ViewModels
{
    public class DoubaoViewModel : NotifyPropertyBase
    {
        private string _webViewSource = "https://www.doubao.com/";
        private double _progressValue;
        private Visibility _progressVisibility = Visibility.Collapsed;
        private CoreWebView2 _coreWebView;

        public CommandBase BackPortalCommand { get; set; }
        public string WebViewSource
        {
            get => _webViewSource;
            set => Set(ref _webViewSource, value);
        }

        public double ProgressValue
        {
            get => _progressValue;
            set => Set(ref _progressValue, value);
        }

        public Visibility ProgressVisibility
        {
            get => _progressVisibility;
            set => Set(ref _progressVisibility, value);
        }

        public CoreWebView2 CoreWebView
        {
            get => _coreWebView;
            set => Set(ref _coreWebView, value);
        }

        public ICommand DragEnterCommand { get; }
        public ICommand DropCommand { get; }

        public DoubaoViewModel()
        {
            DragEnterCommand = new CommandBase(ExecuteDragEnter);
            DropCommand = new CommandBase(ExecuteDrop);
            BackPortalCommand = new CommandBase(BackPortalFunc);
        }

        public void BackPortalFunc(object o)
        {
            PortalWindow portal = new PortalWindow();
            (o as Window).Close();
            portal.Show();
        }

        private void ExecuteDragEnter(object parameter)
        {
            if (parameter is DragEventArgs e)
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                    e.Effects = DragDropEffects.Copy;
            }
        }

        private void ExecuteDrop(object parameter)
        {
            if (parameter is DragEventArgs e)
            {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
                {
                    var files = e.Data.GetData(DataFormats.FileDrop) as string[];
                    if (files != null)
                    {
                        foreach (var file in files)
                            UploadFileByDragDrop(file);
                    }
                }
            }
        }

        private void UploadFileByDragDrop(string filePath)
        {
            if (!File.Exists(filePath)) return;

            var fileName = Path.GetFileName(filePath);
            var script = $@"
                (function() {{
                    const uploadInputs = document.querySelectorAll('input[type=""file""]');
                    if (uploadInputs.length > 0) {{
                        const input = uploadInputs[uploadInputs.length - 1];
                        const event = new Event('change', {{ bubbles: true }});
                        Object.defineProperty(event, 'target', {{ value: {{ files: [new File([''], '{fileName}')] }} }});
                        input.dispatchEvent(event);
                    }}
                }})();
            ";

            _ = CoreWebView?.ExecuteScriptAsync(script);
        }

        // 修复方法签名以匹配事件委托
        public void HandleDownloadStarting(CoreWebView2DownloadStartingEventArgs e)
        {
            try
            {
                // 1. 生成默认文件名
                string defaultFileName = $"豆包文件_{DateTime.Now:yyyyMMddHHmmss}";
                string originalFileName = Path.GetFileName(e.DownloadOperation.Uri);
                string fileName = string.IsNullOrWhiteSpace(originalFileName) ? defaultFileName : originalFileName;

                // 2. 提醒选择后缀
                if (string.IsNullOrEmpty(Path.GetExtension(fileName)))
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        MessageBox.Show(
                            "请在保存时选择合适的文件后缀名（如 .txt、.pdf、.xlsx 等）",
                            "提示：选择文件后缀",
                            MessageBoxButton.OK,
                            MessageBoxImage.Information
                        );
                    });
                }

                // 3. 显示保存对话框，判断用户操作
                bool? dialogResult = null;
                string finalName = null;

                Application.Current.Dispatcher.Invoke(() =>
                {
                    try
                    {
                        var saveDialog = new SaveFileDialog
                        {
                            FileName = fileName,
                            Title = "保存豆包处理后的文件",
                            // 修复 Filter 字符串格式
                            Filter = "文本文件 (*.txt)|*.txt|PDF文件 (*.pdf)|*.pdf|Word文件 (*.docx)|*.docx|Excel文件 (*.xlsx;*.xls)|*.xlsx;*.xls|所有文件 (*.*)|*.*",
                            DefaultExt = ".txt",
                            AddExtension = true
                        };

                        // 移除错误的 Owner 设置，SaveFileDialog 没有 Owner 属性
                        dialogResult = saveDialog.ShowDialog();
                        finalName = saveDialog.FileName;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"保存文件时出错: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                        dialogResult = false;
                    }
                });

                // 关键：如果用户取消保存（点击取消或关闭窗口），则终止下载
                if (dialogResult != true || string.IsNullOrEmpty(finalName))
                {
                    e.Cancel = true; // 取消下载操作
                    return; // 直接返回，不执行后续逻辑
                }

                // 4. 处理用户选择的文件名
                if (string.IsNullOrEmpty(Path.GetExtension(finalName)))
                {
                    finalName += ".txt";
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        MessageBox.Show(
                            $"未检测到文件后缀名，已自动添加 .txt 后缀\n文件路径：{finalName}",
                            "提示：已添加默认后缀",
                            MessageBoxButton.OK,
                            MessageBoxImage.Warning
                        );
                    });
                }
                e.ResultFilePath = finalName;

                // 5. 进度跟踪逻辑
                e.DownloadOperation.StateChanged += (s, args) =>
                {
                    var download = (CoreWebView2DownloadOperation)s;
                    if (download.TotalBytesToReceive.HasValue && download.TotalBytesToReceive > 0)
                    {
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            ProgressVisibility = Visibility.Visible;
                            ProgressValue = (double)download.BytesReceived / (double)download.TotalBytesToReceive * 100;
                        });
                    }

                    if (download.State == CoreWebView2DownloadState.Completed)
                    {
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            ProgressVisibility = Visibility.Collapsed;
                            MessageBox.Show($"下载成功！\n文件路径：{download.ResultFilePath}", "成功");
                        });
                    }
                    else if (download.State == CoreWebView2DownloadState.Interrupted)
                    {
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            ProgressVisibility = Visibility.Collapsed;
                            MessageBox.Show("下载失败，请重试", "错误");
                        });
                    }
                };
            }
            catch (Exception ex)
            {
                // 处理整个下载过程中的异常
                Application.Current.Dispatcher.Invoke(() =>
                {
                    MessageBox.Show($"下载过程中出现错误: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
                });
                e.Cancel = true;
            }
        }
    }
}