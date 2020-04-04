using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

/// <summary>
/// Класс для работы с документами Microsoft Word
/// </summary>
namespace CertificateMaker.core.office
{
    class WordWorker
    {
        /// <summary>
        /// Заменяет выбранный текст в документе Word на указанный
        /// </summary>
        /// <param name="filepath">путь до документа на диске</param>
        /// <param name="wordsForReplace">словарь замены слов, ключ - слово, которое следует заменить, значение - слово, на которое нужно заменить</param>
        public static void ReplaceText(string filepath, Dictionary<string, string> wordsForReplace)
        {
            Word.Application app = null;
            try
            {
                app = new Word.Application();
                Word.Document doc = app.Documents.Open(filepath);
                doc.Activate();
                foreach (KeyValuePair<string, string> keyValue in wordsForReplace)
                {
                    FindAndReplace(app, keyValue.Key, keyValue.Value);
                }
                doc.Save();
                doc.Close();
            }
            finally
            {
                try
                {
                    app.Quit();
                }
                catch (Exception) { }
            }
        }

        /// <summary>
        /// Объединение нескольких документов word в один
        /// </summary>
        /// <param name="filePathOutput">Имя выходного файла</param>
        /// <param name="mergeDocs">Массив имен файлов, из которых будет скопировано содержимое в выходной файл</param>
        public static void MergerDocs(string filePathOutput, string[] mergeDocs)
        {
            object sectionBreak = Word.WdBreakType.wdSectionBreakNextPage;
            Word.Application app = null;
            try
            {
                app = new Word.Application();
                Word.Document doc = app.Documents.Add();
                Word.Selection selection = app.Selection;
                for (int i = 0; i < mergeDocs.Length; i++)
                {
                    selection.InsertFile(mergeDocs[i]);
                    if (i != mergeDocs.Length - 1)
                    {
                        selection.InsertBreak(sectionBreak);
                    }
                }
                doc.SaveAs2(filePathOutput);
                doc.Close();
            }
            finally
            {
                try
                {
                    app.Quit();
                }
                catch (Exception) { }
            }
        }

        protected static void FindAndReplace(Microsoft.Office.Interop.Word.Application doc, object findText, object replaceWithText)
        {
            //options
            object matchCase = false;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;
            //execute find and replace
            doc.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
                ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWithText, ref replace,
                ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }
    }
}
