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
using System.Windows.Navigation;
using System.Windows.Shapes;
using MySearcher;
using System.IO;
using System.Threading;

namespace WebOfSciSearcher {
    /// <summary>
    /// Interaction logic for wos.xaml
    /// </summary>
    public partial class wos : UserControl {
        public wos() {
            InitializeComponent();
            //timer.Interval = 500;
            //timer.Elapsed += new System.Timers.ElapsedEventHandler(onTimeDo);
            //timer.Start();
        }

        System.Timers.Timer timer = new System.Timers.Timer();
        HttpClient hc;
        Thread t;
        int step;
        int count;
        System.Windows.Forms.FolderBrowserDialog folder = new System.Windows.Forms.FolderBrowserDialog();
        private void Button_Click(object sender, RoutedEventArgs e) {
            folder.ShowDialog();
            folder_box.Text = folder.SelectedPath;
        }

        private void BeginDown(object sender, RoutedEventArgs e) {
            hc = new HttpClient(uni.Text, year.Text);
            btn_begin.IsEnabled = false;
            t = new Thread(DownMethod);
            t.Start();

        }

        private void onTimeDo(object sender, System.Timers.ElapsedEventArgs e) {
            Dispatcher.Invoke((Action)delegate {

            });
        }

        private void DownMethod() {

            try {
                string folder = "";
                string type = "";
                Dispatcher.Invoke(new Action(() => { 
                    folder = folder_box.Text;
                    type = type_box.Text;
                    pro.Content = "已经启动，耐心等待哈~~~";
                }));
                if (!hc.OpenWeb()) {
                    Dispatcher.Invoke(new Action(() => { btn_begin.IsEnabled = true; }));
                    return;
                }
                hc.Search(type);
                hc.GetResultHtm();

                count = hc.ChangeNum50();
                step = 1;
                //Dispatcher.Invoke(new Action(() => { pro.Content = "进度：" + step + "/" + count; }));
                //File.WriteAllText(folder + "\\result1.html", hc.GetResultHtm());
                for (int i = 1; i < count+1; i++) {
                    step = i;
                    hc.ChangePage(i);
                    File.WriteAllText(folder + "\\result" + i + ".html", hc.GetResultHtm());
                    Dispatcher.Invoke(new Action(() => { pro.Content = "进度：" + step + "/" + count; }));
                }
                Dispatcher.Invoke(new Action(() => { btn_begin.IsEnabled = true;
                pro.Content = "完成啦，赶紧去看下文件夹吧";
                }));
            } catch (Exception) {
                Dispatcher.Invoke(new Action(() => { btn_begin.IsEnabled = true; }));
                MessageBox.Show("燕子呀，出错啦，重点一次吧！");
            }
        }
        public void Close() {
            t.Abort();
            timer.Stop();
        }






    }
}
