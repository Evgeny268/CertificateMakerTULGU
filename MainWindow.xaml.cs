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

namespace CertificateMaker
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();            
        }

        private void Add_Button_Click(object sender, RoutedEventArgs e)
        {
            var data = new User { TagName = "1", Value = "2" };

            dataGrid.Items.Add(data);
        }
    }
    public class User
    {
        public string TagName { get; set; }

        public ComboBox DataType { get; set; }

        public string Value { get; set; }
    }
}
