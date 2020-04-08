using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;

namespace CertificateMaker.core.presets
{
    /// <summary>
    /// Класс для созранения и загрузки пресета на диск
    /// </summary>
    class PresetLoader
    {
        /// <summary>
        /// Сохранение пресета на диск
        /// </summary>
        /// <param name="filepath">Путь до файла на диске</param>
        /// <param name="preset">Сохраняемый пресет</param>
        public static void SavePreset(string filepath, Preset preset)
        {
            IFormatter formatter = new BinaryFormatter();
            Stream stream = new FileStream(filepath, FileMode.Create, FileAccess.Write);
            formatter.Serialize(stream, preset);
            stream.Close();
        }

        /// <summary>
        /// Загрузка пресета с диска
        /// </summary>
        /// <param name="filepath">Путь до файла на диске</param>
        /// <returns>Загруженный с диска пресет</returns>
        public static Preset LoadPreset(string filepath)
        {
            IFormatter formatter = new BinaryFormatter();
            Stream stream = new FileStream(filepath, FileMode.Open, FileAccess.Read);
            Preset preset = (Preset)formatter.Deserialize(stream);
            return preset;
        }
    }
}
