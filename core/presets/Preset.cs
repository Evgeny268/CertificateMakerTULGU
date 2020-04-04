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
    class Preset
    {
        private string templatePath { get; set; } = null;
        private string excelPath { get; set; } = null;
        private Nullable<int> startRowImport { get; set; } = null;
        private Nullable<int> endRowImport { get; set; } = null;
        List<Table> rows { get; set; } = null;

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
    }
}
