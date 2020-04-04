using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CertificateMaker.core.presets
{
    /// <summary>
    /// Пресеты для каждого тега из шаблона
    /// </summary>
    class Table
    {
        private string templateField { get; set; } = null;
        private TemplateType type { get; set; } = TemplateType.excel;
        private Nullable<int> value { get; set; } = null;

        /// <summary>
        /// Конструктор для тега из шаблона
        /// </summary>
        /// <param name="templateField">название тега</param>
        /// <param name="type">тип тега</param>
        /// <param name="value">значение</param>
        public Table(string templateField, TemplateType type, int? value)
        {
            this.templateField = templateField;
            this.type = type;
            this.value = value;
        }
    }

    public enum TemplateType
    {
        excel,
        generate
    }
}
