using System;
using System.Collections.Generic;
using System.Diagnostics;
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
using System.Windows.Shapes;

namespace WebOfSciSearcher {
    /// <summary>
    /// Interaction logic for WOSSearch.xaml
    /// </summary>
    public partial class WOSSearch : Window {
        public WOSSearch() {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e) {
            wosUI.Visibility = System.Windows.Visibility.Visible;
            schoolSearch.Visibility = System.Windows.Visibility.Hidden;
            ana.Visibility = System.Windows.Visibility.Hidden;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e) {
            wosUI.Visibility = System.Windows.Visibility.Hidden;
            schoolSearch.Visibility = System.Windows.Visibility.Visible;
            ana.Visibility = System.Windows.Visibility.Hidden;
        }

        private void analyseClick(object sender, RoutedEventArgs e) {
            wosUI.Visibility = System.Windows.Visibility.Hidden;
            schoolSearch.Visibility = System.Windows.Visibility.Hidden;
            ana.Visibility = System.Windows.Visibility.Visible;
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e) {
            Process p = Process.GetCurrentProcess();
            p.Kill();
        }
    }
}
