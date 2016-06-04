using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using WebSearcher;

namespace WebOfSciSearcher {
    /// <summary>
    /// Interaction logic for SchoolSearch.xaml
    /// </summary>
    public partial class SchoolSearch : UserControl, IDisposable {
        public void Dispose() {
            if (dtsour != null) {
                dtsour.Dispose();
                dtsour = null;
            }
            if (t != null) {
                t.Dispose();
                t = null;
            }
        }
        public SchoolSearch() {
            InitializeComponent();
            configs = Utils.GetConfig();
            foreach (var schoolName in configs.Keys) {
                select_school.Items.Add(schoolName);
            }
        }
        Searcher searcher;
        Dictionary<string,Config> configs = new Dictionary<string,Config>();
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
            if (!initSearcher(select_school.Text)) {
                return;
            }
            getSData get = new getSData(searcher.GetData);
            get.BeginInvoke(int.Parse(thCount.Text), openFileTextBox.Text, null, null);
            t.Interval = 1000;
            t.Elapsed += new System.Timers.ElapsedEventHandler(updateProc);
            t.Start();
            //btn.IsEnabled = false;
        }

        private bool initSearcher(string text) {
            searcher = new WebSearcher.SchoolSearchCommon();
            try {
                searcher.config = configs[text];
            } catch (Exception) {
                return false;
            }
            //if (text == "NIH") {
            //    searcher = new NIH();
            //} else if (text == "University of Liege") {
            //    searcher = new UniversityOfLiege();
            //} else if (text == "University of Reading") {
            //    searcher = new UniversityOfReading();
            //} else if (text == "University of Edinburgh") {
            //    searcher = new UniversityOfEdinburgh();
            //} else if (text == "University of Bern") {
            //    searcher = new UniversityOfBern();
            //} else if (text == "University of Bristol") {
            //    searcher = new UniversityOfBristol();
            //} else if (text == "University of Cambridge") {
            //    searcher = new UniversityOfCambridge();
            //} else if (text == "University of Oxford") {
            //    searcher = new UniversityOfOxford();
            //} else if (text == "Southampton") {
            //    searcher = new Southampton();
            //} else if (text == "Imperial College London") {
            //    searcher = new ImperialCollegeLondon();
            //} else if (text == "University of St Andrews") {
            //    searcher = new UniversityOfStAndrews();
            //} else if (text == "University of Exeter") {
            //    searcher = new UniversityOfExeter();
            //} else if (text == "Loughborough University") {
            //    searcher = new LoughboroughUniversity();
            //} else if (text == "University of Leeds") {
            //    searcher = new UniversityOfLeeds();
            //} else if (text == "University of Sheffield") {
            //    searcher = new UniversityOfLeeds();
            //} else if (text == "University of Sussex") {
            //    searcher = new UniversityOfSussex();
            //} else if (text == "University of Kent") {
            //    searcher = new UniversityOfKent();
            //} else if (text == "University of Leicester") {
            //    searcher = new UniversityOfLeicester();
            //} else if (text == "University of Dundee") {
            //    searcher = new UniversityOfDundee();
            //} else if (text == "Swansea University") {
            //    searcher = new SwanseaUniversity();
            //} else if (text == "University of Stirling") {
            //    searcher = new UniversityOfStirling();
            //} else if (text == "University of Birmingham") {
            //    searcher = new UniversityOfBirmingham();
            //} else if (text == "University College London") {
            //    searcher = new UniversityCollegeLondon();
            //} else {
            //    return false;
            //}
            return true;
        }

        private void updateProc(object sender, System.Timers.ElapsedEventArgs e) {
            try {
                if (searcher != null) {
                    Dispatcher.Invoke((Action)delegate {
                        this.proc.Content = searcher.over + "/" + searcher.all + "\r\nYES :" + searcher.yes + "  NO :" + searcher.no;
                    });
                }
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
            if (searcher.all > 0) {
                dtsour = searcher.dt.Copy();
                try {
                    Dispatcher.Invoke((Action)delegate {
                        dt.ItemsSource = dtsour.DefaultView;
                        dt.Columns[0].Width = 100;
                        dt.Columns[1].Width = 100;
                        dt.Columns[2].Width = 100;
                    });
                } catch (Exception) {

                }
            }
        }
    }
}
