using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Path = System.IO.Path;

namespace wordFileMerge
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Combine_Click(object sender, RoutedEventArgs e)
        {
            progressLabel.Content = "";
            if (string.IsNullOrWhiteSpace(srcFolder.Text))
            {
                progressLabel.Content = "请输入源目录";
                return;
            }

            if (string.IsNullOrWhiteSpace(TargetDocFile.Text))
                TargetDocFile.Text = "合并.docx";


            Dispatcher.Invoke(() =>
            {
                progressBar.Visibility = Visibility.Visible;
            });

            try
            {
                string filePaths = srcFolder.Text.Trim();

                string[] allWordDocuments = Directory.GetFiles(filePaths, "*.doc*", SearchOption.AllDirectories);
                //Or if you want only SearchOptions.TopDirectoryOnly
                if (!TargetDocFile.Text.EndsWith(".docx"))
                {
                    TargetDocFile.Text += ".docx";
                }
                string outputFileName = Path.Combine(filePaths, TargetDocFile.Text);
                progressLabel.Content = WordDocHelper.Merge(allWordDocuments, outputFileName, true, UpdateProgress);
            }
            finally
            {
                progressBar.Visibility = Visibility.Collapsed;
            }
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void UpdateProgress(string fileName, string progress)
        {
            Dispatcher.Invoke(() =>
            {
                progressLabel.Content = $"{progress.PadRight(10)}, 正在处理：{fileName}";
            });
        }
    }
}
