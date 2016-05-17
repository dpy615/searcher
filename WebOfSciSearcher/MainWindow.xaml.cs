using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
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
using WebSearcher;

namespace WebOfSciSearcher {
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window {
        
        public MainWindow() {
            InitializeComponent();
        }
        Southampton south = new Southampton();
        DataTable dtsour = new DataTable();
        System.Timers.Timer t = new System.Timers.Timer();

        OpenFileDialog openFile = new OpenFileDialog();

        private void Button_Click(object sender, RoutedEventArgs e) {
            openFile.ShowDialog();
            openFileTextBox.Text = openFile.FileName;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e) {
            if (string.IsNullOrEmpty(openFileTextBox.Text) || !File.Exists(openFileTextBox.Text)) {
                MessageBox.Show("检查文件是否正确");
                return;
            }
            getSData get = new getSData(south.GetData);
            get.BeginInvoke(6, openFileTextBox.Text, null, null);
            t.Interval = 1000;
            t.Elapsed += new System.Timers.ElapsedEventHandler(updateProc);
            t.Start();
            btn.IsEnabled = false;


            //south.GetData(1, openFileTextBox.Text);
        }

        private void updateProc(object sender, System.Timers.ElapsedEventArgs e) {
            try {
                Dispatcher.Invoke((Action)delegate {
                    this.proc.Content = south.over + "/" + south.all;
                });
            } catch (Exception) {

            }
        }



        private void SetDataSources(IAsyncResult ar) {
            //Dispatcher.Invoke((Action)delegate {
            //    dt.ItemsSource = south.dt.DefaultView;
            //}); 
        }


        delegate void getSData(int i, string t);

        private void Button_Click_2(object sender, RoutedEventArgs e) {
            //south.isRun = false;
            //t.Stop();
            //btn.IsEnabled = true;
            dtsour = south.dt.Copy();
            try {
                Dispatcher.Invoke((Action)delegate {
                    dt.ItemsSource = dtsour.DefaultView;
                });
            } catch (Exception) {

            }
        }
    }
}
