using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CertificateMaker.core.presets
{
    /// <summary>
    /// Пресет для программы
    /// </summary>
    [Serializable]
    class Preset
    {
        public string templatePath { get; set; } = null;
        public string excelPath { get; set; } = null;
        public Nullable<int> startRowImport { get; set; } = null;
        public Nullable<int> endRowImport { get; set; } = null;
        public List<Table> rows { get; set; } = null;

        public Preset() {
            rows = new List<Table>();
        }

        /// <summary>
        /// Конструктор для пресета
        /// </summary>
        /// <param name="templatePath">Путь до файла с шаблоном</param>
        /// <param name="excelPath">Путь до файла excel</param>
        /// <param name="startRowImport">Номер строки, с которой начать импорт</param>
        /// <param name="endRowImport">Номер строки, на которой завершить импорт</param>
        /// <param name="rows">Лист с настройками для каждого тега из шаблона</param>
        public Preset(string templatePath, string excelPath, int? startRowImport, int? endRowImport, List<Table> rows)
        {
            this.templatePath = templatePath;
            this.excelPath = excelPath;
            this.startRowImport = startRowImport;
            this.endRowImport = endRowImport;
            this.rows = rows;
        }

        public Table GetTableByName(string templateField)
        {
            if (rows == null) return null;
            for (int i = 0; i < rows.Count(); i++)
            {
                if (rows[i].templateField.Equals(templateField))
                {
                    return rows[i];
                }
            }
            return null;
        }
    }
}
