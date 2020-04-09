using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CertificateMaker.core.appWorker
{
    /// <summary>
    /// Слушатель для отслеживания прогресса
    /// </summary>
    interface IWorkerListener
    {
        /// <summary>
        /// Вызывается в момент выполнения очередного шага
        /// </summary>
        /// <param name="stage">Текущий этап</param>
        /// <param name="current">Текущее значение</param>
        /// <param name="all">Всего необходимо выполнить</param>
        void WorkStatus(WorkStage stage, int current, int all);
    }

    public enum WorkStage
    {
        READ_FROM_EXCEL, //Чтение из excel
        CREATE_DOC, //Создание документа word, тут применяется отслеживание прогресса (int current, int all)
        MERGE_DOC, //Объединение документов
        DELETE_TEMP_FILES, //Удаление временной директории
        DONE //Все задачи выполнены
    }
}
