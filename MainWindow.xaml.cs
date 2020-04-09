using CertificateMaker.core.presets;
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
        core.presets.Preset preset = new core.presets.Preset();
        public MainWindow()
        {
            InitializeComponent();            
        }               

        private void btnWordLoad_Click(object sender, RoutedEventArgs e)
        {
            // Configure open file dialog box
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.FileName = "Document"; // Default file name
            dlg.DefaultExt = ".DOC .DOCX .DOCM"; // Default file extension
            dlg.Filter = "MS Word Документ|*.DOC;*.DOCX;*.DOCM"; // Filter files by extension

            // Show open file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                // Open document
                WordFileName.Content = dlg.FileName;
                preset.templatePath = dlg.FileName;
            }
        }

        private void btnExcelLoad_Click(object sender, RoutedEventArgs e)
        {
            // Configure open file dialog box
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.FileName = "Document"; // Default file name
            dlg.DefaultExt = ".xls"; // Default file extension
            dlg.Filter = "Excel таблица|*.xls;*.xlsx"; // Filter files by extension

            // Show open file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                // Open document
                ExcelFileName.Content = dlg.FileName;
                preset.excelPath = dlg.FileName;
            }
        }

        private void OnPreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            e.Handled = !e.Text.All(IsGood);
        }
        private void OnPasting(object sender, DataObjectPastingEventArgs e)
        {
            var stringData = (string)e.DataObject.GetData(typeof(string));
            if (stringData == null || !stringData.All(IsGood))
                e.CancelCommand();
        }

        bool IsGood(char c)
        {
            if (c >= '0' && c <= '9')
                return true;
            return false;
        }

        private void subLoad_Click(object sender, RoutedEventArgs e)
        {
            // Configure open file dialog box
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.FileName = "Preset"; // Default file name
            dlg.DefaultExt = ".cm"; // Default file extension
            dlg.Filter = "Пресет|*.cm"; // Filter files by extension

            // Show open file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                //Load preset
                preset = core.presets.PresetLoader.LoadPreset(dlg.FileName);
                UpdateFromPreset();
            }
        }

        private void subSave_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = "Preset"; // Default file name
            dlg.DefaultExt = ".cm"; // Default file extension
            dlg.Filter = "Пресет|*.cm"; // Filter files by extension  

            if (toRow.Text != null && !toRow.Text.Equals(""))
            {
                preset.endRowImport = int.Parse(toRow.Text);
            }
            if (fromRow.Text != null && !fromRow.Text.Equals(""))
            {
                preset.startRowImport = int.Parse(fromRow.Text);
            }

            if (dlg.ShowDialog() == true)
                core.presets.PresetLoader.SavePreset(dlg.FileName, preset);
        }

        private void UpdateFromPreset()
        {
            fromRow.Text = "";
            toRow.Text = "";
            ExcelFileName.Content = "Выберите файл Excel";
            WordFileName.Content = "Выберите шаблон в Word";
            templateItems.Items.Clear();
            if (preset.startRowImport != null)
                fromRow.Text = preset.startRowImport.GetValueOrDefault().ToString();
            if (preset.endRowImport != null)
                toRow.Text = preset.endRowImport.GetValueOrDefault().ToString();
            if (preset.excelPath != null)
                ExcelFileName.Content = preset.excelPath;
            if (preset.templatePath != null)
                WordFileName.Content = preset.templatePath;
            for (int i = 0; i < preset.rows.Count(); i++)
            {
                templateItems.Items.Add(preset.rows[i]);
            }
        }
        private void Add_Button_Click(object sender, RoutedEventArgs e)
        {
            string templateField = textBoxTemplateName.Text;
            if (templateField.Equals(""))
            {
                //TODO заругать, что не введне тэг
                return;
            }
            TemplateType type;
            int indexType = comboBoxType.SelectedIndex;
            if (indexType == 0)
            {
                type = TemplateType.excel;
            }
            else
            {
                type = TemplateType.generate;
            }
            string value = textBoxValue.Text;
            if (value.Equals(""))
            {
                //TODO заругать, что не введено значение
                return;
            }
            core.presets.Table oldTable = preset.GetTableByName(templateField);
            if (oldTable != null)
            {
                //TODO заругать, что такой тэг уже есть
                return;
            }
            textBoxTemplateName.Text = "";
            comboBoxType.SelectedIndex = 0;
            textBoxValue.Text = "";
            core.presets.Table addTable = new core.presets.Table(templateField, type, int.Parse(value));
            preset.rows.Add(addTable);
            UpdateFromPreset();
        }

        private void add(object sender, RoutedEventArgs e)
        {

        }
    }
    public class User
    {
        public string TagName { get; set; }

        public ComboBox DataType { get; set; }

        public Nullable<int> Value { get; set; }
    }
}
