using Microsoft.VisualStudio.TestTools.UnitTesting;
using suppaman45;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Newtonsoft.Json;
using System.Diagnostics;

namespace suppaman45.Tests
{
    [TestClass()]
    public class SettingManagerTests
    {
        /// <summary>
        /// ファイルが存在する場合
        /// </summary>
        [TestMethod()]
        public void LoadSettingsTest_FileExists()
        {
            //期待値
            var expectedSettings = new UserSettings
            {
                ReadFileDir = "C:\\Users\\Hamu\\Documents\\生産指示\\",
                ReadFileExtention = ".xlsm",
                ReadSheetName = "ライン構成",
                NamedRange = "memTableHead",
                ReadIgnoreThrethold = 3,
                WriteFilepath = "C:\\Users\\Hamu\\Documents\\統計情報.xlsm",
                WriteSheetname = "Records",
                WriteTableName = "テーブル1",
                ArchiveDirPath = "C:\\Users\\Hamu\\Documents\\アーカイブ\\",
                UnprocessedDatesRangeName = "UnprocessedDatesRange",
                WaitTime = 5000,
                InvalidPatterns = new List<string>
                {
                    "[\\x00 -\\x7F]", "^[ぁ-ゞ]+$", "[（）、・]"
                }
            };

            string filePath = "../../../suppaman45Tests/TestFiles/SettingManager/期待される既定値.json";
            var settingManager = new SettingManager(filePath);
            var result = settingManager.LoadSettings();

            //期待値とファイルの中身が等しいか比較
            var exceptedJson = JsonConvert.SerializeObject(expectedSettings, Formatting.Indented);
            var resultJson = JsonConvert.SerializeObject(result, Formatting.Indented);
            Assert.IsTrue(exceptedJson == resultJson);
        }

        /// <summary>
        /// ファイルが存在しない場合
        /// </summary>
        [TestMethod()]
        public void LoadSettingsTest_FileNotExists()
        {
            //期待値（既定値）
            var expectedSettings = new UserSettings
            {
                ReadFileDir = "C:\\Users\\Hamu\\Documents\\生産指示\\",
                ReadFileExtention = ".xlsm",
                ReadSheetName = "ライン構成",
                NamedRange = "memTableHead",
                ReadIgnoreThrethold = 3,
                WriteFilepath = "C:\\Users\\Hamu\\Documents\\統計情報.xlsm",
                WriteSheetname = "Records",
                WriteTableName = "テーブル1",
                ArchiveDirPath = "C:\\Users\\Hamu\\Documents\\アーカイブ\\",
                UnprocessedDatesRangeName = "UnprocessedDatesRange",
                WaitTime = 5000,
                InvalidPatterns = new List<string>
                {
                    "[\\x00 -\\x7F]", "^[ぁ-ゞ]+$", "[（）、・]"
                }
            };

            string filepath = "ignorepath.json";
            var settingsManager = new SettingManager(filepath);
            var result = settingsManager.LoadSettings();

            //期待値と既定値が等しいか比較
            var exceptedJson = JsonConvert.SerializeObject(expectedSettings, Formatting.Indented);
            var resultJson = JsonConvert.SerializeObject(result, Formatting.Indented);
            Debug.WriteLine(exceptedJson);
            Debug.WriteLine(resultJson);
            Assert.IsTrue(exceptedJson == resultJson);

        }

        /// <summary>
        /// こわれたJson
        /// </summary>
        [TestMethod]
        public void LoadSettingTest_BlokenJson()
        {
            //期待値（既定値）
            var expectedSettings = new UserSettings
            {
                ReadFileDir = "C:\\Users\\Hamu\\Documents\\生産指示\\",
                ReadFileExtention = ".xlsm",
                ReadSheetName = "ライン構成",
                NamedRange = "memTableHead",
                ReadIgnoreThrethold = 3,
                WriteFilepath = "C:\\Users\\Hamu\\Documents\\統計情報.xlsm",
                WriteSheetname = "Records",
                WriteTableName = "テーブル1",
                ArchiveDirPath = "C:\\Users\\Hamu\\Documents\\アーカイブ\\",
                UnprocessedDatesRangeName = "UnprocessedDatesRange",
                WaitTime = 5000,
                InvalidPatterns = new List<string>
                {
                "[\\x00 -\\x7F]","^[ぁ-ゞ]+$","[（）、・]"
                }
            };

            string filepath = "../../../suppaman45Tests/TestFiles/SettingManager/壊れたJson.json";

            //ファイルがない場合には失敗扱い
            if (!File.Exists(filepath)) { Assert.Fail(); }

            var settingsManager = new SettingManager(filepath);
            var result = settingsManager.LoadSettings();

            //期待値と既定値が等しいか比較
            var exceptedJson = JsonConvert.SerializeObject(expectedSettings, Formatting.Indented);
            var resultJson = JsonConvert.SerializeObject(result, Formatting.Indented);
            Debug.WriteLine(exceptedJson);
            Debug.WriteLine(resultJson);
            Assert.IsTrue(exceptedJson == resultJson);
        }

        /// <summary>
        /// 保存のテスト
        /// </summary>
        [TestMethod()]
        public void SaveSettingsTest()
        {
            var resultPath = "../../../suppaman45Tests/TestFiles/SettingManager/出力Json.json";
            File.Delete(resultPath);

            var settings = new UserSettings();
            var settingManager = new SettingManager(resultPath);
            settingManager.SaveSettings(settings);

            Assert.IsTrue(File.Exists(resultPath));
        }
    }
}