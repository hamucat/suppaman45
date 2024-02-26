using Microsoft.VisualStudio.TestTools.UnitTesting;
using suppaman45;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace suppaman45.Tests
{
    [TestClass()]
    public class ArchiveManagerTests
    {
        /// <summary>
        /// 年で分割するテスト
        /// </summary>
        [TestMethod()]
        public void SplitByYearTest()
        {
            var testDatas = new List<ExcelData>
            {
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","キルア"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","シャルナーク"),
                new ExcelData(new DateTime(2023,12,25),"3P","パクノダ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","コルトピ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","アイザック"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","クラピカ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","アルカ"),
                new ExcelData(new DateTime(2023,12,25),"3P","シルバ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","ゴレイヌ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","セドカン"),
                new ExcelData(new DateTime(2024,12,25),"ツイストⅠ","ヒソカ"),
                new ExcelData(new DateTime(2024,12,25),"1Ｐ","イルミ"),
                new ExcelData(new DateTime(2024,12,25),"3P","ウイング"),
                new ExcelData(new DateTime(2024,12,25),"茶Ⅰ","ポックル"),
                new ExcelData(new DateTime(2024,12,25),"茶Ⅱ","キリコ"),
                new ExcelData(new DateTime(2024,12,25),"ツイストⅠ","フェイタン"),
                new ExcelData(new DateTime(2024,12,25),"1Ｐ","レオリオ"),
                new ExcelData(new DateTime(2024,12,25),"3P","ミト"),
                new ExcelData(new DateTime(2024,12,25),"茶Ⅰ","ニコル"),
                new ExcelData(new DateTime(2024,12,25),"茶Ⅱ","ネオン"),
                new ExcelData(new DateTime(2025,12,25),"ツイストⅠ","クロロ"),
                new ExcelData(new DateTime(2025,12,25),"1Ｐ","ビスケ"),
                new ExcelData(new DateTime(2025,12,25),"3P","ハンゾー"),
                new ExcelData(new DateTime(2025,12,25),"茶Ⅰ","トンパ"),
                new ExcelData(new DateTime(2025,12,25),"茶Ⅱ","バショウ"),
                new ExcelData(new DateTime(2025,12,25),"ツイストⅠ","ゴン"),
                new ExcelData(new DateTime(2025,12,25),"1Ｐ","シズク"),
                new ExcelData(new DateTime(2025,12,25),"3P","センリツ"),
                new ExcelData(new DateTime(2025,12,25),"茶Ⅰ","メンチ"),
                new ExcelData(new DateTime(2025,12,25),"茶Ⅱ","ヴェーゼ")
            };

            var archiveManager = new ArchiveManager(new UserSettings());
            var result = archiveManager.SplitByYear(testDatas);

            Assert.AreEqual(3, result.Count);
            foreach (var item in result)
            {
                Assert.AreEqual(10, item.Count);
            }
        }

        /// <summary>
        /// CSV書き込みのテスト　create
        /// </summary>
        [TestMethod()]
        public void WriteToCSVTest()
        {
            var testDatas = new List<ExcelData>
            {
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","キルア"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","シャルナーク"),
                new ExcelData(new DateTime(2023,12,25),"3P","パクノダ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","コルトピ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","アイザック"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","クラピカ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","アルカ"),
                new ExcelData(new DateTime(2023,12,25),"3P","シルバ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","ゴレイヌ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","セドカン"),
                new ExcelData(new DateTime(2024,12,25),"ツイストⅠ","ヒソカ"),
                new ExcelData(new DateTime(2024,12,25),"1Ｐ","イルミ"),
                new ExcelData(new DateTime(2024,12,25),"3P","ウイング"),
                new ExcelData(new DateTime(2024,12,25),"茶Ⅰ","ポックル"),
                new ExcelData(new DateTime(2024,12,25),"茶Ⅱ","キリコ"),
                new ExcelData(new DateTime(2024,12,25),"ツイストⅠ","フェイタン"),
                new ExcelData(new DateTime(2024,12,25),"1Ｐ","レオリオ"),
                new ExcelData(new DateTime(2024,12,25),"3P","ミト"),
                new ExcelData(new DateTime(2024,12,25),"茶Ⅰ","ニコル"),
                new ExcelData(new DateTime(2024,12,25),"茶Ⅱ","ネオン"),
                new ExcelData(new DateTime(2025,12,25),"ツイストⅠ","クロロ"),
                new ExcelData(new DateTime(2025,12,25),"1Ｐ","ビスケ"),
                new ExcelData(new DateTime(2025,12,25),"3P","ハンゾー"),
                new ExcelData(new DateTime(2025,12,25),"茶Ⅰ","トンパ"),
                new ExcelData(new DateTime(2025,12,25),"茶Ⅱ","バショウ"),
                new ExcelData(new DateTime(2025,12,25),"ツイストⅠ","ゴン"),
                new ExcelData(new DateTime(2025,12,25),"1Ｐ","シズク"),
                new ExcelData(new DateTime(2025,12,25),"3P","センリツ"),
                new ExcelData(new DateTime(2025,12,25),"茶Ⅰ","メンチ"),
                new ExcelData(new DateTime(2025,12,25),"茶Ⅱ","ヴェーゼ")
            };

            var settingManager = new SettingManager("../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var userSettings = settingManager.LoadSettings();

            if (Directory.Exists(userSettings.ArchiveDirPath))
            {
                Directory.Delete(userSettings.ArchiveDirPath, true);
            }

            var archiveManager = new ArchiveManager(userSettings);
            archiveManager.WriteToCSV(testDatas);

            Assert.AreEqual(3, Directory.GetFiles(userSettings.ArchiveDirPath).Length);
        }
    }
}