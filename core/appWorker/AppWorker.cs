using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace CertificateMaker.core.appWorker
{
    class AppWorker
    {
        private presets.Preset preset;

        private static readonly string TEMP_FOLDER_NAME = "app_temp";
        private static readonly string TEMP_DOC_NAME = "DOC";

        public IWorkerListener workerListener { private get; set; } = null;

        private string filePathOut = null;
        private List<string[]> excelData = null;
        private int[] cellsNeeedReadFromExcel = null;
        private Dictionary<int, int> excelCellsIndex = null;

        public void MakeDocs(string filePathOut, presets.Preset preset)
        {
            if (filePathOut == null)
            {
                throw new NullReferenceException("filePathOut is null!");
            }

            if (preset == null)
            {
                throw new NullReferenceException("preset is null!");
            }

            if (!CheckPreset(preset))
            {
                throw new ArgumentException("some field in preset not set!");
            }
            this.filePathOut = filePathOut;
            this.preset = preset;
            FindCellsNumForRead();
            if (workerListener != null)
            {
                workerListener.WorkStatus(WorkStage.READ_FROM_EXCEL, 0, 0);
            }
            ReadDataFromExcel();
            CreateTempFolder();
            for (int i = 0; i < excelData.Count(); i++)
            {
                if (workerListener != null)
                {
                    workerListener.WorkStatus(WorkStage.CREATE_DOC, i, excelData.Count());
                }
                CreateDoc(i);
            }
            if (workerListener != null)
            {
                workerListener.WorkStatus(WorkStage.MERGE_DOC, 0, 0);
            }
            MergeDocs();
            if (workerListener != null)
            {
                workerListener.WorkStatus(WorkStage.DELETE_TEMP_FILES, 0, 0);
            }
            DeleteTempFolder();
            if (workerListener != null)
            {
                workerListener.WorkStatus(WorkStage.DONE, 0, 0);
            }
        }

        public bool CheckPreset(presets.Preset preset)
        {
            if (preset.templatePath == null || preset.excelPath == null || preset.startRowImport == null || preset.endRowImport == null || preset.rows == null)
            {
                return false;
            }
            else
            {
                if (preset.rows.Count() == 0) return false;
                for (int i = 0; i < preset.rows.Count(); i++)
                {
                    if (preset.rows[i].templateField == null || preset.rows[i].value == null) return false;
                }
                return true;
            }
        }

        private void FindCellsNumForRead()
        {
            List<int> needRead = new List<int>();
            for (int i = 0; i < preset.rows.Count(); i++)
            {
                if (preset.rows[i].type == presets.TemplateType.excel)
                {
                    if (preset.rows[i].value != null)
                    {
                        int val = preset.rows[i].value.GetValueOrDefault();
                        if (!needRead.Contains(val))
                        {
                            needRead.Add(val);
                        }
                    }
                }
            }
            needRead.Sort();
            cellsNeeedReadFromExcel = needRead.ToArray();
            excelCellsIndex = new Dictionary<int, int>();
            for (int i = 0; i < cellsNeeedReadFromExcel.Length; i++)
            {
                excelCellsIndex.Add(cellsNeeedReadFromExcel[i], i);
            }
        }

        private void ReadDataFromExcel()
        {
            excelData = office.ExcelWorker.ReadCells(preset.excelPath, preset.startRowImport.GetValueOrDefault(), 
                preset.endRowImport.GetValueOrDefault(), cellsNeeedReadFromExcel);
        }

        private void CreateTempFolder()
        {
            if (Directory.Exists(Directory.GetCurrentDirectory()+"/"+TEMP_FOLDER_NAME))
            {
                Directory.Delete(Directory.GetCurrentDirectory()+"/"+TEMP_FOLDER_NAME, true);
            }
            Directory.CreateDirectory(Directory.GetCurrentDirectory()+"/"+TEMP_FOLDER_NAME);
        }

        private void CreateDoc(int index)
        {
            File.Copy(preset.templatePath, Directory.GetCurrentDirectory()+"/"+TEMP_FOLDER_NAME + "/" + TEMP_DOC_NAME + index);
            Dictionary<string, string> wordsForReplace = new Dictionary<string, string>();
            string[] excelRow = excelData[index];
            for (int i = 0; i < preset.rows.Count(); i++)
            {
                string key = preset.rows[i].templateField;
                string value = "NULL";
                if (preset.rows[i].type == presets.TemplateType.generate)
                {
                    value = (index + preset.rows[i].value.GetValueOrDefault()).ToString();
                }
                else
                {
                    int column = excelCellsIndex[preset.rows[i].value.GetValueOrDefault()];
                    value = excelRow[column];
                }
                wordsForReplace.Add(key, value);
            }
            office.WordWorker.ReplaceText(Directory.GetCurrentDirectory()+"/"+TEMP_FOLDER_NAME + "/" + TEMP_DOC_NAME + index, wordsForReplace);
        }

        private void MergeDocs()
        {
            string[] fnames = new string[excelData.Count()];
            for (int i = 0; i < excelData.Count(); i++)
            {
                fnames[i] = Directory.GetCurrentDirectory()+"/"+TEMP_FOLDER_NAME + "/" + TEMP_DOC_NAME + i;
            }
            office.WordWorker.MergerDocs(filePathOut, fnames);
        }

        public void DeleteTempFolder()
        {
            if (Directory.Exists(Directory.GetCurrentDirectory()+"/"+TEMP_FOLDER_NAME))
            {
                Directory.Delete(Directory.GetCurrentDirectory()+"/"+TEMP_FOLDER_NAME, true);
            }
        }
    }
}
