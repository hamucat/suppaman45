using Newtonsoft.Json;
using Parzan.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Parzan
{
    class UserSettings
    {
        public string JsonPath { get; set; }
    }

    /// <summary>
    /// suppaman45の場所を記録するやつ
    /// </summary>
    public class JsonPathManager
    {
        string saveFilePath;
        public JsonPathManager()
        {
            saveFilePath = "suppaman_path.json";
        }

        /// <summary>
        /// suppaman45.exeのパスを返す
        /// </summary>
        /// <returns></returns>
        /// <exception cref="IOException"></exception>
        public string LoadPath()
        {
            try
            {
                if (!File.Exists(saveFilePath))
                {
                    return "../suppaman45/settings.json";
                }

                string json = File.ReadAllText(saveFilePath);
                var userSettings = JsonConvert.DeserializeObject<UserSettings>(json);
                return userSettings.JsonPath;
            }
            catch (JsonException)
            {
                File.Delete(saveFilePath);
                return "../suppaman45/settings.json";
            }
        }

        public void SavePath(string path)
        {
            var settings = new UserSettings { JsonPath = path };
            string json = JsonConvert.SerializeObject(settings, Formatting.Indented);
            File.WriteAllText(saveFilePath, json);
        }
    }
}
