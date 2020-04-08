using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CertificateMaker.core.appWorker
{
    interface IWorkerListener
    {
        void WorkStatus(WorkStage stage, int current, int all);
    }

    public enum WorkStage
    {
        READ_FROM_EXCEL,
        CREATE_DOC,
        MERGE_DOC,
        DELETE_TEMP_FILES
    }
}
