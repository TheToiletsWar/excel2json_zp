
using Microsoft.Win32;
using Microsoft.WindowsAPICodePack.Dialogs;
using System;
using System.Collections.Generic;
using System.ComponentModel;
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

namespace excel2json_zpmc
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        // Excel导入数据管理
        private DataManager mDataMgr;
        private string mCurrentXlsx;
        // 打开的excel文件名，不包含后缀xlsx。。。
        private String FileName;

        //private BackgroundWorker _backgroundworker;
        public MainWindow()
        {
            InitializeComponent();
            mDataMgr = new DataManager();
         
        }
        
        private void btn_import_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.RestoreDirectory = true;
            dlg.Filter = "Excel File(*.xlsx)|*.xlsx";
            if ((bool)dlg.ShowDialog())
            {
                this.import_filepath.Text = dlg.FileName;
            }
        }

        private void btn_out_Click(object sender, RoutedEventArgs e)
        {
            CommonOpenFileDialog dlg = new CommonOpenFileDialog();
            dlg.RestoreDirectory = true;
            dlg.IsFolderPicker = true;
            if (dlg.ShowDialog()==CommonFileDialogResult.Ok)
            {
                this.out_filepath.Text = dlg.FileName;
            }
        }
        /// <summary>
        /// 保存导出文件
        /// </summary>
        private void saveJsonToFile(string filter)
        {
            try
            {
                SaveFileDialog dlg = new SaveFileDialog();
                dlg.RestoreDirectory = true;
                dlg.Filter = filter;
                dlg.FileName = FileName;
                if ((bool)dlg.ShowDialog())
                {
                    lock (mDataMgr)
                    {
                                mDataMgr.saveJson(dlg.FileName);
                    }
                }// end of if
            }
            catch (Exception ex)
            {
            }
        }

        private void btn_converter_Click(object sender, RoutedEventArgs e)
        {
            loadExcel(this.import_filepath.Text.Trim());
        }

        private void loadExcel(string path)
        {
            mCurrentXlsx = path;
            FileName = System.IO.Path.GetFileNameWithoutExtension(path);

            //使用主线程加载excel文件,可更改为后台线程执行
            lock (this.mDataMgr)
            {
                this.mDataMgr.loadExcel(path, "utf8-nobom");
                
            }
          
        }
    }
}
