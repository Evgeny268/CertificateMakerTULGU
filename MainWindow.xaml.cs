using CertificateMaker.core.appWorker;
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
using System.Threading;

namespace CertificateMaker
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, core.appWorker.IWorkerListener
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
                Progress_Lbl.Content = "";
                Progress_Lbl.Background = Brushes.DarkGray;
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
                Progress_Lbl.Content = "";
                Progress_Lbl.Background = Brushes.DarkGray;
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
            if (preset.rows != null)
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
                Progress_Lbl.Content = "Введите название поля!";
                Progress_Lbl.Background = Brushes.Red;
                textBoxTemplateName.BorderBrush = Brushes.Red;
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
                Progress_Lbl.Content = "Введите значение!";
                Progress_Lbl.Background = Brushes.Red;
                textBoxValue.BorderBrush = Brushes.Red;
                return;
            }
            core.presets.Table oldTable = preset.GetTableByName(templateField);
            if (oldTable != null)
            {
                Progress_Lbl.Content = "Такое название поля уже существует!";
                Progress_Lbl.Background = Brushes.Red;
                textBoxTemplateName.BorderBrush = Brushes.Red;
                return;
            }
            textBoxTemplateName.Text = "";
            comboBoxType.SelectedIndex = 0;
            textBoxValue.Text = "";
            core.presets.Table addTable = new core.presets.Table(templateField, type, int.Parse(value));
            preset.rows.Add(addTable);
            UpdateFromPreset();
            Progress_Lbl.Content = "";
            Progress_Lbl.Background = Brushes.DarkGray;
            textBoxTemplateName.BorderBrush = Brushes.Gray;
            if (textBoxTemplateName.Text.Equals("") || textBoxTemplateName.Text.Equals("Название поля"))
            {
                AddBtn.IsEnabled = false;
                textBoxTemplateName.Text = "Название поля";
                textBoxTemplateName.Foreground = Brushes.DarkGray;
            }
            if (textBoxValue.Text.Equals("") || textBoxValue.Text.Equals("Номер столбца") || textBoxValue.Text.Equals("Начальное значение"))
            {
                AddBtn.IsEnabled = false;
                if (comboBoxType.SelectedIndex == 0)
                    textBoxValue.Text = "Номер столбца";
                else
                    textBoxValue.Text = "Начальное значение";
                textBoxValue.Foreground = Brushes.DarkGray;
            }
        }

        private void ClickDeleteField(object sender, RoutedEventArgs e)
        {            
            int currentRowIndex = templateItems.SelectedIndex;
            if (currentRowIndex != -1)
            {
                preset.rows.RemoveAt(currentRowIndex);
                UpdateFromPreset();
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (!core.appWorker.AppWorker.CheckPreset(preset))
            {
                Progress_Lbl.Content = "Не все данные заполнены";
                Progress_Lbl.Background = Brushes.Red;
                return;
            }

            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = "Document"; // Default file name
            dlg.DefaultExt = ".DOCX"; // Default file extension
            dlg.Filter = "MS Word Документ|*.DOCX"; // Filter files by extension

            if (dlg.ShowDialog() == true)
            {
                core.appWorker.AppWorker appWorker = new core.appWorker.AppWorker();
                try
                {
                    appWorker.workerListener = this;
                    Thread thread = new Thread(() => appWorker.MakeDocs(dlg.FileName, preset));
                    thread.Start();
                }
                catch
                {
                    Progress_Lbl.Content = "Произошла ошибка! Проверьте пути до файлов и начилие MS Office";
                    Progress_Lbl.Background = Brushes.Red;
                    progressStatus.Value = 0;
                    return;
                }
            }         
        }

        public void WorkStatus(WorkStage stage, int current, int all)
        {
            Application.Current.Dispatcher.Invoke(new Action(() =>
                {
                switch (stage)
                {
                    case WorkStage.READ_FROM_EXCEL:
                        progressStatus.Value = 0;
                        Progress_Lbl.Content = "Чтение из Excel";
                        Progress_Lbl.Background = Brushes.DarkGray;
                        break;
                    case WorkStage.CREATE_DOC:
                        progressStatus.Value = TransferToProgress(current, 0, all, 0, 100);
                        Progress_Lbl.Content = "Создание документа " + current + " из " + all;
                        Progress_Lbl.Background = Brushes.DarkGray;
                        break;
                    case WorkStage.MERGE_DOC:
                        progressStatus.Value = 100;
                        Progress_Lbl.Content = "Объединение документов";
                        Progress_Lbl.Background = Brushes.DarkGray;
                        break;
                    case WorkStage.DELETE_TEMP_FILES:
                        progressStatus.Value = 100;
                        Progress_Lbl.Content = "Удаление временной директории";
                        Progress_Lbl.Background = Brushes.DarkGray;
                        break;
                    case WorkStage.DONE:
                        progressStatus.Value = 0;
                        Progress_Lbl.Content = "Готово";
                        Progress_Lbl.Background = Brushes.DarkGray;
                        System.Media.SystemSounds.Exclamation.Play();
                        break;
                }
            }));
        }

        private int TransferToProgress(int value, int in_min, int in_max, int out_min, int out_max)
        {
            return (value - in_min) * (out_max - out_min) / (in_max - in_min) + out_min;
        }

        private void toRow_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!fromRow.Text.Equals(""))
                if ((int.Parse(toRow.Text) <= int.Parse(fromRow.Text)) )
                {
                    Progress_Lbl.Content = "Значение ДО не может быть меньше значения ОТ!";
                    Progress_Lbl.Background = Brushes.Red;
                    fromRow.BorderBrush = Brushes.Red;
                    toRow.BorderBrush = Brushes.Red;
                    btnSave.IsEnabled = false;
                    return;
                }            
            if (toRow.Text.Equals(""))
            {
                Progress_Lbl.Content = "Значение ДО не может пустым!";
                Progress_Lbl.Background = Brushes.Red;
                toRow.BorderBrush = Brushes.Red;
                btnSave.IsEnabled = false;
                return;
            }
            preset.endRowImport = int.Parse(toRow.Text);           
            Progress_Lbl.Content = "";
            Progress_Lbl.Background = Brushes.DarkGray;
            fromRow.BorderBrush = Brushes.Gray;
            toRow.BorderBrush = Brushes.Gray;
            btnSave.IsEnabled = true;
        }

        private void fromRow_LostFocus(object sender, RoutedEventArgs e)
        {            
            if (fromRow.Text.Equals("0") || fromRow.Text.Equals(""))
            {
                Progress_Lbl.Content = "Значение полея ОТ не может быть пустым или равно 0";
                Progress_Lbl.Background = Brushes.Red;
                fromRow.BorderBrush = Brushes.Red;
                btnSave.IsEnabled = false;
                return;
            }
            if (!toRow.Text.Equals(""))
                if (int.Parse(toRow.Text) <= int.Parse(fromRow.Text))
                {
                    Progress_Lbl.Content = "Значение ОТ не может быть больше значения ДО!";
                    Progress_Lbl.Background = Brushes.Red;
                    fromRow.BorderBrush = Brushes.Red;
                    toRow.BorderBrush = Brushes.Red;
                    btnSave.IsEnabled = false;
                    return;
                }
            preset.startRowImport = int.Parse(fromRow.Text);
            Progress_Lbl.Content = "";
            Progress_Lbl.Background = Brushes.DarkGray;
            fromRow.BorderBrush = Brushes.Gray;
            toRow.BorderBrush = Brushes.Gray;
            btnSave.IsEnabled = true;
        }

        private void textBoxValue_LostFocus(object sender, RoutedEventArgs e)
        {
            if (textBoxValue.Text.Equals("") || textBoxValue.Text.Equals("Номер столбца") || textBoxValue.Text.Equals("Начальное значение"))
            {
                AddBtn.IsEnabled = false;
                if (comboBoxType.SelectedIndex == 0)
                    textBoxValue.Text = "Номер столбца";
                else
                    textBoxValue.Text = "Начальное значение";
                textBoxValue.Foreground = Brushes.DarkGray;
            }
            if (textBoxValue.Text.Equals("0"))
            {
                Progress_Lbl.Content = "Значение не может быть равно 0";
                Progress_Lbl.Background = Brushes.Red;
                textBoxValue.BorderBrush = Brushes.Red;
                btnSave.IsEnabled = false;
                return;
            }
            Progress_Lbl.Content = "";
            Progress_Lbl.Background = Brushes.DarkGray;
            textBoxValue.BorderBrush = Brushes.Gray;            
            btnSave.IsEnabled = true;      
        }

        private void textBoxTemplateName_GotFocus(object sender, RoutedEventArgs e)
        {
            if (textBoxTemplateName.Text.Equals("Название поля"))
                textBoxTemplateName.Text = "";
            textBoxTemplateName.Foreground = Brushes.Black;
        }

        private void textBoxTemplateName_LostFocus(object sender, RoutedEventArgs e)
        {
            if (textBoxTemplateName.Text.Equals("") || textBoxTemplateName.Text.Equals("Название поля"))
            {
                AddBtn.IsEnabled = false;
                textBoxTemplateName.Text = "Название поля";
                textBoxTemplateName.Foreground = Brushes.DarkGray;
            }
            if (!textBoxTemplateName.Text.Equals("") && !textBoxTemplateName.Text.Equals("Название поля") && !textBoxValue.Text.Equals("") && !textBoxValue.Text.Equals("Номер столбца") && !textBoxValue.Text.Equals("Начальное значение"))
                AddBtn.IsEnabled = true;
        }

        private void textBoxValue_GotFocus(object sender, RoutedEventArgs e)
        {
            if (textBoxValue.Text.Equals("Номер столбца") || textBoxValue.Text.Equals("Начальное значение"))
                textBoxValue.Text = "";
            textBoxValue.Foreground = Brushes.Black;
        }

        private void comboBoxType_MouseLeave(object sender, MouseEventArgs e)
        {
            if (textBoxValue.Text.Equals("") || textBoxValue.Text.Equals("Номер столбца") || textBoxValue.Text.Equals("Начальное значение"))
            {
                if (comboBoxType.SelectedIndex == 0)
                    textBoxValue.Text = "Номер столбца";
                else
                    textBoxValue.Text = "Начальное значение";
                textBoxValue.Foreground = Brushes.DarkGray;
            }
        }

        private void textBoxValue_MouseLeave(object sender, MouseEventArgs e)
        {
            if (!textBoxTemplateName.Text.Equals("") && !textBoxTemplateName.Text.Equals("Название поля") && !textBoxValue.Text.Equals("") && !textBoxValue.Text.Equals("Номер столбца") && !textBoxValue.Text.Equals("Начальное значение"))
                AddBtn.IsEnabled = true;
            else
                AddBtn.IsEnabled = false;
        }

        private void textBoxTemplateName_MouseLeave(object sender, MouseEventArgs e)
        {
            if (!textBoxTemplateName.Text.Equals("") && !textBoxTemplateName.Text.Equals("Название поля") && !textBoxValue.Text.Equals("") && !textBoxValue.Text.Equals("Номер столбца") && !textBoxValue.Text.Equals("Начальное значение"))
                AddBtn.IsEnabled = true;
            else
                AddBtn.IsEnabled = false;
        }
    }
}
