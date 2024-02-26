using Microsoft.VisualStudio.TestTools.UnitTesting;
using suppaman45;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;

namespace suppaman45.Tests
{
    [TestClass()]
    public class ExcelDataReaderTests
    {
        /// <summary>
        /// ファイルが存在する場合
        /// </summary>
        [TestMethod()]
        public void GetPathTest_FileExists()
        {
            var settingManager = new SettingManager(@"../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();

            var testDate = new DateTime(2023, 12, 25);
            var formattedDate = testDate.ToString("yyMMdd");
            var fileName = formattedDate + settings.ReadFileExtention;
            var filePath = Path.Combine(settings.ReadFileDir, fileName);
            using (File.Create(filePath))
            {
                var excelDataReader = new ExcelDataReader(settings);
                var result = excelDataReader.GetPath(testDate);
                Assert.AreEqual(filePath, result);
            }

            File.Delete(filePath);
        }

        /// <summary>
        /// ファイルが存在しない場合
        /// </summary>
        [TestMethod()]
        public void GetPathTest_FileNotExists()
        {
            var settingManager = new SettingManager(@"../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();

            var excelDataReader = new ExcelDataReader(settings);

            Assert.ThrowsException<FileNotFoundException>(() => { excelDataReader.GetPath(new DateTime(2024, 1, 1)); });
        }

        /// <summary>
        /// 読み込みテスト　成功
        /// </summary>
        [TestMethod()]
        public void GetExcelDatesTest_AssertSucces()
        {
            var exceptedData = new List<ExcelData>
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
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ヒソカ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","イルミ"),
                new ExcelData(new DateTime(2023,12,25),"3P","ウイング"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","ポックル"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","キリコ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","フェイタン"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","レオリオ"),
                new ExcelData(new DateTime(2023,12,25),"3P","ミト"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","ニコル"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","ネオン"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","クロロ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","ビスケ"),
                new ExcelData(new DateTime(2023,12,25),"3P","ハンゾー"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","トンパ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","バショウ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ゴン"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","シズク"),
                new ExcelData(new DateTime(2023,12,25),"3P","センリツ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","メンチ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","ヴェーゼ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","カイト"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","カルト"),
                new ExcelData(new DateTime(2023,12,25),"3P","ウボォーギン"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","マチ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","ジン"),
                new ExcelData(new DateTime(2023,12,25),"3P","ズシ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ノブナガ")
            };

            var settingManager = new SettingManager(@"../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();

            var excelDataReader = new ExcelDataReader(settings);
            var result = excelDataReader.GetExcelDatas(@"../../../suppaman45Tests/TestFiles/ExcelDataReader/AssertSuccess.xlsm", new DateTime(2023, 12, 25));

            Assert.IsNotNull(result);
            Assert.AreEqual(exceptedData.Count, result.Count);

            for (int i = 0; i < result.Count; i++)
            {
                Assert.AreEqual(exceptedData[i].Date, result[i].Date);
                Assert.AreEqual(exceptedData[i].LineName, result[i].LineName);
                Assert.AreEqual(exceptedData[i].EmploeeName, result[i].EmploeeName);
            }
        }

        /// <summary>
        /// ロック中の読み込みテスト　成功
        /// </summary>
        [TestMethod()]
        public void GetExcelDatesTest_WhenLockedAssertSucces()
        {
            var exceptedData = new List<ExcelData>
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
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ヒソカ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","イルミ"),
                new ExcelData(new DateTime(2023,12,25),"3P","ウイング"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","ポックル"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","キリコ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","フェイタン"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","レオリオ"),
                new ExcelData(new DateTime(2023,12,25),"3P","ミト"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","ニコル"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","ネオン"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","クロロ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","ビスケ"),
                new ExcelData(new DateTime(2023,12,25),"3P","ハンゾー"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","トンパ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","バショウ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ゴン"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","シズク"),
                new ExcelData(new DateTime(2023,12,25),"3P","センリツ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","メンチ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","ヴェーゼ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","カイト"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","カルト"),
                new ExcelData(new DateTime(2023,12,25),"3P","ウボォーギン"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","マチ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","ジン"),
                new ExcelData(new DateTime(2023,12,25),"3P","ズシ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ノブナガ")
            };

            var settingManager = new SettingManager(@"../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();

            var xlapp = new Application();
            xlapp.Workbooks.Open(Path.GetFullPath(@"../../../suppaman45Tests/TestFiles/ExcelDataReader/AssertSuccess.xlsm"));

            var excelDataReader = new ExcelDataReader(settings);
            var result = excelDataReader.GetExcelDatas(@"../../../suppaman45Tests/TestFiles/ExcelDataReader/AssertSuccess.xlsm", new DateTime(2023, 12, 25));

            xlapp.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlapp);

            Assert.IsNotNull(result);
            Assert.AreEqual(exceptedData.Count, result.Count);

            for (int i = 0; i < result.Count; i++)
            {
                Assert.AreEqual(exceptedData[i].Date, result[i].Date);
                Assert.AreEqual(exceptedData[i].LineName, result[i].LineName);
                Assert.AreEqual(exceptedData[i].EmploeeName, result[i].EmploeeName);
            }
        }

        /// <summary>
        /// 読み込みテスト(追記)　成功
        /// </summary>
        [TestMethod()]
        public void GetExcelDatesTest_Append_AssertSucces()
        {
            var exceptedData = new List<ExcelData>
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
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ヒソカ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","イルミ"),
                new ExcelData(new DateTime(2023,12,25),"3P","ウイング"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","ポックル"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","キリコ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","フェイタン"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","レオリオ"),
                new ExcelData(new DateTime(2023,12,25),"3P","ミト"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","ニコル"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","ネオン"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","クロロ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","ビスケ"),
                new ExcelData(new DateTime(2023,12,25),"3P","ハンゾー"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","トンパ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","バショウ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ゴン"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","シズク"),
                new ExcelData(new DateTime(2023,12,25),"3P","センリツ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","メンチ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","ヴェーゼ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","カイト"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","カルト"),
                new ExcelData(new DateTime(2023,12,25),"3P","ウボォーギン"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","マチ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","ジン"),
                new ExcelData(new DateTime(2023,12,25),"3P","ズシ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ノブナガ"),
                new ExcelData(new DateTime(2023,12,26),"ツイストⅠ","ルーデウス"),
                new ExcelData(new DateTime(2023,12,26),"1Ｐ","ノルン"),
                new ExcelData(new DateTime(2023,12,26),"3P","エリナリーゼ"),
                new ExcelData(new DateTime(2023,12,26),"ツイストⅡ","イゾルテ"),
                new ExcelData(new DateTime(2023,12,26),"茶Ⅰ","リニア"),
                new ExcelData(new DateTime(2023,12,26),"茶Ⅱ","ジノ"),
                new ExcelData(new DateTime(2023,12,26),"ツイストⅠ","ロキシー"),
                new ExcelData(new DateTime(2023,12,26),"1Ｐ","ルイジェルド"),
                new ExcelData(new DateTime(2023,12,26),"3P","タルバンド"),
                new ExcelData(new DateTime(2023,12,26),"ツイストⅡ","シャンドル"),
                new ExcelData(new DateTime(2023,12,26),"茶Ⅰ","プルセナ"),
                new ExcelData(new DateTime(2023,12,26),"茶Ⅱ","マルタ"),
                new ExcelData(new DateTime(2023,12,26),"ツイストⅠ","シルフィエット"),
                new ExcelData(new DateTime(2023,12,26),"1Ｐ","ザノバ"),
                new ExcelData(new DateTime(2023,12,26),"3P","ギース"),
                new ExcelData(new DateTime(2023,12,26),"ツイストⅡ","ドーガ"),
                new ExcelData(new DateTime(2023,12,26),"茶Ⅰ","ジュリ"),
                new ExcelData(new DateTime(2023,12,26),"茶Ⅱ","ランドルフ"),
                new ExcelData(new DateTime(2023,12,26),"ツイストⅠ","エリス"),
                new ExcelData(new DateTime(2023,12,26),"1Ｐ","クリフ"),
                new ExcelData(new DateTime(2023,12,26),"3P","サウロス"),
                new ExcelData(new DateTime(2023,12,26),"ツイストⅡ","オーベール"),
                new ExcelData(new DateTime(2023,12,26),"茶Ⅰ","ガル"),
                new ExcelData(new DateTime(2023,12,26),"茶Ⅱ","パックス"),
                new ExcelData(new DateTime(2023,12,26),"ツイストⅠ","パウロ"),
                new ExcelData(new DateTime(2023,12,26),"1Ｐ","ナナホシ"),
                new ExcelData(new DateTime(2023,12,26),"3P","フィリップ"),
                new ExcelData(new DateTime(2023,12,26),"ツイストⅡ","ピレモン"),
                new ExcelData(new DateTime(2023,12,26),"茶Ⅰ","ニナ"),
                new ExcelData(new DateTime(2023,12,26),"茶Ⅱ","ジンジャー"),
                new ExcelData(new DateTime(2023,12,26),"ツイストⅠ","ゼニス"),
                new ExcelData(new DateTime(2023,12,26),"1Ｐ","オルステッド"),
                new ExcelData(new DateTime(2023,12,26),"3P","トーマス"),
                new ExcelData(new DateTime(2023,12,26),"ツイストⅡ","トリスティーナ"),
                new ExcelData(new DateTime(2023,12,26),"ツイストⅠ","リーリャ"),
                new ExcelData(new DateTime(2023,12,26),"1Ｐ","ギレーヌ"),
                new ExcelData(new DateTime(2023,12,26),"3P","エドナ"),
                new ExcelData(new DateTime(2023,12,26),"ツイストⅡ","アリエル"),
                new ExcelData(new DateTime(2023,12,26),"ツイストⅠ","アイシャ"),
                new ExcelData(new DateTime(2023,12,26),"3P","レイダ"),
                new ExcelData(new DateTime(2023,12,26),"ツイストⅡ","ルーク")
            };

            var settingManager = new SettingManager(@"../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();

            var excelDataReader = new ExcelDataReader(settings);
            var result = excelDataReader.GetExcelDatas(@"../../../suppaman45Tests/TestFiles/ExcelDataReader/AssertSuccess.xlsm", new DateTime(2023, 12, 25));
            result = excelDataReader.GetExcelDatas(@"../../../suppaman45Tests/TestFiles/ExcelDataReader/AppendData.xlsm", new DateTime(2023, 12, 26), result);

            Assert.IsNotNull(result);
            Assert.AreEqual(exceptedData.Count, result.Count);

            for (int i = 0; i < result.Count; i++)
            {
                Assert.AreEqual(exceptedData[i].Date, result[i].Date);
                Assert.AreEqual(exceptedData[i].LineName, result[i].LineName);
                Assert.AreEqual(exceptedData[i].EmploeeName, result[i].EmploeeName);
            }

            ////ExceptedData作成用
            //foreach (var item in result)
            //{
            //    Debug.WriteLine("new ExcelData { Date = new DateTime(2023,12,26),LineName = \"" + item.LineName + "\",\"" + item.EmploeeName + "\"},");
            //}
        }

        /// <summary>
        /// CellsUsed()で空白セルをちゃんと飛ばせているかのテスト
        /// </summary>
        [TestMethod]
        public void GetExcelDataTest_CellsUsed()
        {
            var exceptedData = new List<ExcelData>
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
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ヒソカ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","イルミ"),
                new ExcelData(new DateTime(2023,12,25),"3P","ウイング"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","ポックル"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","キリコ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","フェイタン"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","レオリオ"),
                new ExcelData(new DateTime(2023,12,25),"3P","ミト"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","ニコル"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","ネオン"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","クロロ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","ビスケ"),
                new ExcelData(new DateTime(2023,12,25),"3P","ハンゾー"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","トンパ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","バショウ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ゴン"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","シズク"),
                new ExcelData(new DateTime(2023,12,25),"3P","センリツ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","メンチ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","ヴェーゼ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","カイト"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","カルト"),
                new ExcelData(new DateTime(2023,12,25),"3P","ウボォーギン"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","マチ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","ジン"),
                new ExcelData(new DateTime(2023,12,25),"3P","ズシ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ノブナガ")
            };

            var settingManager = new SettingManager(@"../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();

            var excelDataReader = new ExcelDataReader(settings);
            var result = excelDataReader.GetExcelDatas(@"../../../suppaman45Tests/TestFiles/ExcelDataReader/CellsUsed.xlsm", new DateTime(23, 12, 25));

            Assert.IsNotNull(result);
            Assert.AreEqual(exceptedData.Count, result.Count);
            foreach (var item in result)
            {
                if (string.IsNullOrEmpty(item.LineName))
                {
                    Assert.Fail();
                }
            }
        }

        /// <summary>
        /// 表記ゆれを置換
        /// </summary>
        [TestMethod()]
        public void GetExcelDatasTest_ReplaceEmploeeName()
        {
            var exceptedData = new List<ExcelData>
            {
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","キルア"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","シャルナーク"),
                new ExcelData(new DateTime(2023,12,25),"3P","パクノダ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","コルトピ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","アイザック"),//ネテロ
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","クラピカ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","アルカ"),
                new ExcelData(new DateTime(2023,12,25),"3P","シルバ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","ゴレイヌ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","セドカン"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ヒソカ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","イルミ"),//ギタラクル
                new ExcelData(new DateTime(2023,12,25),"3P","ウイング"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","ポックル"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","キリコ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","フェイタン"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","レオリオ"),
                new ExcelData(new DateTime(2023,12,25),"3P","ミト"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","ニコル"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","ネオン"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","クロロ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","ビスケ"),
                new ExcelData(new DateTime(2023,12,25),"3P","ハンゾー"),//半蔵
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","トンパ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","バショウ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ゴン"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","シズク"),
                new ExcelData(new DateTime(2023,12,25),"3P","センリツ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","メンチ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","ヴェーゼ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","カイト"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","カルト"),
                new ExcelData(new DateTime(2023,12,25),"3P","ウボォーギン"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","マチ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","ジン"),
                new ExcelData(new DateTime(2023,12,25),"3P","ズシ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ノブナガ")
            };

            var settingManager = new SettingManager(@"../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();

            var excelDataReader = new ExcelDataReader(settings);
            var result = excelDataReader.GetExcelDatas(@"../../../suppaman45Tests/TestFiles/ExcelDataReader/ReplaceEmploeeName.xlsm", new DateTime(2023, 12, 25));

            Assert.IsNotNull(result);
            Assert.AreEqual(exceptedData.Count, result.Count);

            for (int i = 0; i < result.Count; i++)
            {
                Assert.AreEqual(exceptedData[i].Date, result[i].Date);
                Assert.AreEqual(exceptedData[i].LineName, result[i].LineName);
                Assert.AreEqual(exceptedData[i].EmploeeName, result[i].EmploeeName);
            }
        }

        /// <summary>
        /// 重複を除外
        /// </summary>
        [TestMethod()]
        public void GetExcelDataTest_IsDeplicate()
        {
            var exceptedData = new List<ExcelData>
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
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ヒソカ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","イルミ"),
                new ExcelData(new DateTime(2023,12,25),"3P","ウイング"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","ポックル"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","キリコ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","フェイタン"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","レオリオ"),
                new ExcelData(new DateTime(2023,12,25),"3P","ミト"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","ニコル"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","ネオン"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","クロロ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","ビスケ"),
                new ExcelData(new DateTime(2023,12,25),"3P","ハンゾー"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","トンパ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","バショウ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ゴン"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","シズク"),
                new ExcelData(new DateTime(2023,12,25),"3P","センリツ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","メンチ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","ヴェーゼ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","カイト"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","カルト"),
                new ExcelData(new DateTime(2023,12,25),"3P","ウボォーギン"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","マチ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","ジン"),
                new ExcelData(new DateTime(2023,12,25),"3P","ズシ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ノブナガ")
            };

            var settingManager = new SettingManager(@"../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();

            var excelDataReader = new ExcelDataReader(settings);
            var result = excelDataReader.GetExcelDatas(@"../../../suppaman45Tests/TestFiles/ExcelDataReader/IsDuplicate.xlsm", new DateTime(2023, 12, 25));

            Assert.IsNotNull(result);
            Assert.AreEqual(exceptedData.Count, result.Count);

            for (int i = 0; i < result.Count; i++)
            {
                Assert.AreEqual(exceptedData[i].Date, result[i].Date);
                Assert.AreEqual(exceptedData[i].LineName, result[i].LineName);
                Assert.AreEqual(exceptedData[i].EmploeeName, result[i].EmploeeName);
            }
        }

        /// <summary>
        /// 正規表現にヒットする名前を除外
        /// </summary>
        [TestMethod()]
        public void GetExcelDatasTest_IsValidEmploeeName()
        {
            var exceptedData = new List<ExcelData>
            {
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","キルア"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","シャルナーク"),
                new ExcelData(new DateTime(2023,12,25),"3P","パクノダ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","コルトピ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","アイザック"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ノブナガ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","イルミ・ゾルディック")
            };

            var settingManager = new SettingManager(@"../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();

            var excelDataReader = new ExcelDataReader(settings);
            var result = excelDataReader.GetExcelDatas(@"../../../suppaman45Tests/TestFiles/ExcelDataReader/IsValidEmploeeName.xlsm", new DateTime(2023, 12, 25));

            Assert.IsNotNull(result);
            Assert.AreEqual(exceptedData.Count, result.Count);
        }

        /// <summary>
        /// 昼交代っぽい行を除外
        /// </summary>
        [TestMethod()]
        public void GetExcelDatasTest_isKoutai()
        {
            var exceptedData = new List<ExcelData>
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
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ヒソカ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","イルミ"),
                new ExcelData(new DateTime(2023,12,25),"3P","ウイング"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","ポックル"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","キリコ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","フェイタン"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","レオリオ"),
                new ExcelData(new DateTime(2023,12,25),"3P","ミト"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","ニコル"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","ネオン"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","クロロ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","ビスケ"),
                new ExcelData(new DateTime(2023,12,25),"3P","ハンゾー"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","トンパ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","バショウ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ゴン"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","シズク"),
                new ExcelData(new DateTime(2023,12,25),"3P","センリツ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","メンチ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","ヴェーゼ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","カイト"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","カルト"),
                new ExcelData(new DateTime(2023,12,25),"3P","ウボォーギン"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","マチ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","ジン"),
                new ExcelData(new DateTime(2023,12,25),"3P","ズシ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ノブナガ")
            };

            var settingManager = new SettingManager(@"../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();

            var excelDataReader = new ExcelDataReader(settings);
            var result = excelDataReader.GetExcelDatas(@"../../../suppaman45Tests/TestFiles/ExcelDataReader/isKoutai.xlsm", new DateTime(2023, 12, 25));

            Assert.IsNotNull(result);
            Assert.AreEqual(exceptedData.Count, result.Count);
        }

        /// <summary>
        /// ライン名を除外
        /// </summary>
        [TestMethod()]
        public void GetExcelDatasTest_isLineName()
        {
            var exceptedData = new List<ExcelData>
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
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ヒソカ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","イルミ"),
                new ExcelData(new DateTime(2023,12,25),"3P","ウイング"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","ポックル"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","キリコ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","フェイタン"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","レオリオ"),
                new ExcelData(new DateTime(2023,12,25),"3P","ミト"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","ニコル"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","ネオン"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","クロロ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","ビスケ"),
                new ExcelData(new DateTime(2023,12,25),"3P","ハンゾー"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","トンパ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","バショウ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ゴン"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","シズク"),
                new ExcelData(new DateTime(2023,12,25),"3P","センリツ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅰ","メンチ"),
                new ExcelData(new DateTime(2023,12,25),"茶Ⅱ","ヴェーゼ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","カイト"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","カルト"),
                new ExcelData(new DateTime(2023,12,25),"3P","ウボォーギン"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","マチ"),
                new ExcelData(new DateTime(2023,12,25),"1Ｐ","ジン"),
                new ExcelData(new DateTime(2023,12,25),"3P","ズシ"),
                new ExcelData(new DateTime(2023,12,25),"ツイストⅠ","ノブナガ")
            };

            var settingManager = new SettingManager(@"../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();

            var excelDataReader = new ExcelDataReader(settings);
            var result = excelDataReader.GetExcelDatas(@"../../../suppaman45Tests/TestFiles/ExcelDataReader/isLineName.xlsm", new DateTime(2023, 12, 25));

            Assert.IsNotNull(result);
            Assert.AreEqual(exceptedData.Count, result.Count);
        }

        /// <summary>
        /// データがない日付リストを取得するテスト　空白セルにぶち当たって成功　戻り値は中身ある
        /// </summary>
        [TestMethod()]
        public void GetUnprocessedDateListTest_AssertSuccess()
        {
            var settingManager = new SettingManager(@"../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();
            settings.WriteFilepath = @"../../../suppaman45Tests/TestFiles/ExcelDataReader/GetUnprocessedDateListTest/AssertSuccess.xlsm";

            var excelDataReader = new ExcelDataReader(settings);
            var result = excelDataReader.GetUnprocessedDateList();

            Assert.IsNotNull(result);
            Assert.AreEqual(result.Count, 29);
        }

        /// <summary>
        /// データがない日付リストを取得するテスト　空白セルにぶち当たって成功　戻り値は空のリスト
        /// </summary>
        [TestMethod()]
        public void GetUnprocessedDateListTest_AssertSuccess_ReturnEmpty()
        {
            var settingManager = new SettingManager(@"../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();
            settings.WriteFilepath = @"../../../suppaman45Tests/TestFiles/ExcelDataReader/GetUnprocessedDateListTest/AssertSuccess_ReturnEmpty.xlsm";

            var excelDataReader = new ExcelDataReader(settings);
            var result = excelDataReader.GetUnprocessedDateList();

            Assert.IsNotNull(result);
            Assert.AreEqual(result.Count, 0);
        }

        /// <summary>
        /// データがない日付リストを取得するテスト　空白でないセルにぶち当たって成功　戻り値は中身ある
        /// </summary>
        [TestMethod()]
        public void GetUnprocessedDateListTest_AssertSuccess_NextToDataContainsExits()
        {
            var settingManager = new SettingManager(@"../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();
            settings.WriteFilepath = @"../../../suppaman45Tests/TestFiles/ExcelDataReader/GetUnprocessedDateListTest/AssertSuccess_NextToDataContainsExits.xlsm";

            var excelDataReader = new ExcelDataReader(settings);
            var result = excelDataReader.GetUnprocessedDateList();

            Assert.IsNotNull(result);
            Assert.AreEqual(result.Count, 29);
        }

        /// <summary>
        /// データがない日付リストを取得するテスト　空白でないセルにぶち当たって成功　戻り値は空のリスト
        /// </summary>
        [TestMethod()]
        public void GetUnprocessedDateListTest_AssertSuccess_NextToDataContainsExits_ReturnEmpty()
        {
            var settingManager = new SettingManager(@"../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();
            settings.WriteFilepath = @"../../../suppaman45Tests/TestFiles/ExcelDataReader/GetUnprocessedDateListTest/AssertSuccess_NextToDataContainsExits_ReturnEmpty.xlsm";

            var excelDataReader = new ExcelDataReader(settings);
            var result = excelDataReader.GetUnprocessedDateList();

            Assert.IsNotNull(result);
            Assert.AreEqual(result.Count, 0);
        }

        /// <summary>
        /// データがない日付リストを取得するテスト　ファイルがない
        /// </summary>
        [TestMethod()]
        public void GetUnprocessedDateListTest_FileNotFound()
        {
            var settingManager = new SettingManager(@"../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();
            settings.WriteFilepath = @"../../../suppaman45Tests/TestFiles/ExcelDataReader/GetUnprocessedDateListTest/FileNotFound.xlsm";

            var excelDataReader = new ExcelDataReader(settings);

            Assert.ThrowsException<FileNotFoundException>(() => { excelDataReader.GetUnprocessedDateList(); });
        }

        /// <summary>
        /// データがない日付リストを取得するテスト　名前付き範囲がない
        /// </summary>
        [TestMethod()]
        public void GetUnprocessedDateListTest_NamedRangeNotFound()
        {
            var settingManager = new SettingManager(@"../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();
            settings.WriteFilepath = @"../../../suppaman45Tests/TestFiles/ExcelDataReader/GetUnprocessedDateListTest/NamedRangeNotFound.xlsm";

            var excelDataReader = new ExcelDataReader(settings);

            Assert.ThrowsException<NullReferenceException>(() => { excelDataReader.GetUnprocessedDateList(); });
        }
    }
}