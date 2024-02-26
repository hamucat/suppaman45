using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Newtonsoft.Json;

namespace suppaman45
{
    /// <summary>
    /// 設定のIO
    /// </summary>
    public class SettingManager
    {
        private readonly string filePath;
        NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        /// <summary>
        /// 設定のIOをするクラスを初期化
        /// </summary>
        /// <param name="filePath">設定を保存するパス</param>
        public SettingManager(string filePath)
        {
            this.filePath = filePath;
        }

        /// <summary>
        /// 設定の読み込み。ファイルが存在しないまたは壊れている場合には既定値で作成。
        /// </summary>
        /// <returns>UserSetting</returns>
        public UserSettings LoadSettings()
        {
            try
            {
                if (File.Exists(filePath))
                {
                    logger.Debug("settings.json exists.");
                    string json = File.ReadAllText(filePath);
                    return JsonConvert.DeserializeObject<UserSettings>(json);
                }
                else
                {
                    logger.Debug("settings.json not found.");
                    var result = new UserSettings();
                    result.SetDefaultValues();
                    SaveSettings(result);
                    return result;
                }
            }
            //Jsonファイルが壊れていた場合既定値で作って返す。
            catch (JsonException ex)
            {
                logger.Warn("settings.json broken. return default.\\n{0}", ex.Message);
                var result = new UserSettings();
                result.SetDefaultValues();
                return result;
            }
        }

        /// <summary>
        /// 設定の書き込み
        /// </summary>
        /// <param name="userSettings">書き込むUserSetting</param>
        public void SaveSettings(UserSettings userSettings)
        {
            string json = JsonConvert.SerializeObject(userSettings, Formatting.Indented);
            File.WriteAllText(filePath, json);
        }
    }
}
