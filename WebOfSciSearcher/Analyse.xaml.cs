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
using System.IO;
using MySearcher;
using System.Data;

namespace WebOfSciSearcher {
    /// <summary>
    /// Interaction logic for Analyse.xaml
    /// </summary>
    public partial class Analyse : UserControl {

        private DataTable dt = new DataTable();
        HtmAnalyse han;

        public Analyse() {
            InitializeComponent();
        }
        private string desc = "";
        System.Windows.Forms.FolderBrowserDialog dialog = new System.Windows.Forms.FolderBrowserDialog();
        System.Timers.Timer timer = new System.Timers.Timer();

        private void Button_Click_1(object sender, RoutedEventArgs e) {
            dialog.ShowDialog();
            textBox_src.Text = dialog.SelectedPath;
        }

        private void Button_selDesc(object sender, RoutedEventArgs e) {
            dialog.ShowDialog();
            if (textBox_src.Text != "") {
                textBox_des.Text = dialog.SelectedPath + textBox_src.Text.Substring(textBox_src.Text.LastIndexOf("\\")) + ".xls";
            } else {
                textBox_des.Text = dialog.SelectedPath + "\\results";
            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e) {
            if (textBox_src.Text != "") {
                DirectoryInfo di = new DirectoryInfo(textBox_src.Text);
                int count = di.GetFiles().Length;
                han = new HtmAnalyse(textBox_src.Text, count);
                desc = textBox_des.Text;
                han.Analyse();

                timer.Interval = 500;
                timer.Elapsed += new System.Timers.ElapsedEventHandler(onTimeEvt);
                timer.Start();
            }
        }

        private void onTimeEvt(object sender, System.Timers.ElapsedEventArgs e) {
            if (han != null) {

                Dispatcher.Invoke(new Action(() => {
                    proc.Content = "进度：" + han.step + "/" + han.count;
                    if (han.dt.Columns.Count != 0) {
                        DataTable dd = han.dt.Copy();
                        Dg_han.ItemsSource = dd.DefaultView;
                        Dg_han.Columns[0].Width = 150;
                        Dg_han.Columns[1].Width = 150;
                        Dg_han.Columns[2].Width = 150;
                    }
                }));

                if (han.step == han.count) {
                    timer.Stop();
                    try {
                        han.SaveToExcel(han.dt, desc);
                        MessageBox.Show("完成，已保存！");
                        Dispatcher.Invoke(new Action(() => {
                            if (han.dt.Columns.Count != 0) {
                                DataTable dd = han.dt.Copy();
                                Dg_han.ItemsSource = dd.DefaultView;
                                Dg_han.Columns[0].Width = 150;
                                Dg_han.Columns[1].Width = 150;
                                Dg_han.Columns[2].Width = 150;
                            }
                        }));
                    } catch (Exception ex) {
                        MessageBox.Show(ex.Message);
                    }
                }
            }
        }
    }
}
