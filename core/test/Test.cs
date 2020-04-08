using CertificateMaker.core.appWorker;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CertificateMaker.core.test
{
    class Test:appWorker.IWorkerListener
    {
        public void Tester() {
            core.presets.Table table1 = new core.presets.Table("<test1>", core.presets.TemplateType.excel, 4);
            core.presets.Table table2 = new core.presets.Table("<test2>", core.presets.TemplateType.excel, 2);
            core.presets.Table table3 = new core.presets.Table("<test3>", core.presets.TemplateType.excel, 2);
            core.presets.Table table4 = new core.presets.Table("<test4>", core.presets.TemplateType.excel, 1);
            core.presets.Table table5 = new core.presets.Table("<test5>", core.presets.TemplateType.generate, 10);
            List<core.presets.Table> rows = new List<core.presets.Table>();
            rows.Add(table1);
            rows.Add(table2);
            rows.Add(table3);
            rows.Add(table4);
            rows.Add(table5);
            core.presets.Preset preset = new core.presets.Preset("D:/test.docx", "D:/test.xlsx", 1, 4, rows);
            core.appWorker.AppWorker appWorker = new core.appWorker.AppWorker();
            appWorker.workerListener = this;
            appWorker.MakeDocs("D:/result.docx", preset);
        }

        public void WorkStatus(WorkStage stage, int current, int all)
        {
            Console.WriteLine(current);
        }
    }
}
