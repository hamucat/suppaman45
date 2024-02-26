using CsvHelper;
using DocumentFormat.OpenXml.Office2010.Word;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using suppaman45;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.IO;


namespace suppaman45.Tests
{
    [TestClass()]
    public class ExcelDataWriterTests
    {
        /// <summary>
        /// Worksheetまで正しく取得できる成功パターン
        /// </summary>
        [TestMethod()]
        public void OpenTest_AssertSuccess()
        {
            var settingManager = new SettingManager("../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();
            settings.WriteFilepath = "../../../TestFiles/ExcelDataWriter/AssertSuccess.xlsm";

            var excelDataWriter = new ExcelDataWriter(settings);
            excelDataWriter.Open();
            Assert.IsNotNull(excelDataWriter.Worksheet);

            excelDataWriter.XlApp.DisplayAlerts = false;
            excelDataWriter.Workbook.Close();
            excelDataWriter.XlApp.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelDataWriter.XlApp);
        }

        /// <summary>
        /// 存在しないファイル名を指定されたパターン
        /// </summary>
        [TestMethod()]
        public void OpenTest_FileNotFound()
        {
            var settingManager = new SettingManager("../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();
            settings.WriteFilepath = "../../../TestFiles/ExcelDataWriter/FileNotFound.xlsm";

            var excelDataWriter = new ExcelDataWriter(settings);
            Assert.ThrowsException<FileNotFoundException>(() => { excelDataWriter.Open(); });

            excelDataWriter.XlApp.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelDataWriter.XlApp);
        }

        /// <summary>
        /// ワークシートが見つからないパターン
        /// </summary>
        [TestMethod]
        public void OpenTest_WorksheetNotFound()
        {
            var settingManager = new SettingManager("../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();
            settings.WriteFilepath = "../../../TestFiles/ExcelDataWriter/WorksheetNotFound.xlsm";

            var excelDataWriter = new ExcelDataWriter(settings);
            Assert.ThrowsException<System.Runtime.InteropServices.COMException>(() => { excelDataWriter.Open(); });

            excelDataWriter.XlApp.DisplayAlerts = false;
            excelDataWriter.Workbook.Close();
            excelDataWriter.XlApp.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelDataWriter.XlApp);
        }


        /// <summary>
        /// ListObjectを正しく取得できるパターン
        /// </summary>
        [TestMethod()]
        public void FindTableTest_AssertSuccess()
        {
            var settingManager = new SettingManager("../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();
            settings.WriteFilepath = "../../../TestFiles/ExcelDataWriter/AssertSuccess.xlsm";

            var excelDataWriter = new ExcelDataWriter(settings);
            excelDataWriter.Open();
            excelDataWriter.FindTable();
            Assert.IsNotNull(excelDataWriter.WriteTable);

            excelDataWriter.XlApp.DisplayAlerts = false;
            excelDataWriter.Workbook.Close();
            excelDataWriter.XlApp.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelDataWriter.XlApp);
        }

        /// <summary>
        /// テーブル1が見つからないパターン
        /// </summary>
        [TestMethod()]
        public void FindTableTest_TableNotFound()
        {
            var settingManager = new SettingManager("../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();
            settings.WriteFilepath = "../../../TestFiles/ExcelDataWriter/TableNotFound.xlsm";

            var excelDataWriter = new ExcelDataWriter(settings);
            excelDataWriter.Open();

            Assert.ThrowsException<System.Runtime.InteropServices.COMException>(() => excelDataWriter.FindTable());

            excelDataWriter.XlApp.DisplayAlerts = false;
            excelDataWriter.Workbook.Close();
            excelDataWriter.XlApp.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelDataWriter.XlApp);
        }

        /// <summary>
        /// カラのテーブルに書き込む場合に正しく書き込まれることを確認するテスト
        /// </summary>
        [TestMethod()]
        public void WriteDatasTest1_TableEmpty()
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
            var settingManager = new SettingManager("../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();
            settings.WriteFilepath = "../../../TestFiles/ExcelDataWriter/TableEmpty.xlsm";

            var excelDataWriter = new ExcelDataWriter(settings);
            excelDataWriter.XlApp.Visible = true;
            excelDataWriter.Open();
            excelDataWriter.FindTable();
            excelDataWriter.WriteDatas(testDatas);

            //書き込み後の行数で比較
            var lastRow = excelDataWriter.WriteTable.DataBodyRange[1, 1].End(XlDirection.xlDown);
            Assert.AreEqual(lastRow.Row, testDatas.Count + 1);

            excelDataWriter.XlApp.DisplayAlerts = false;
            excelDataWriter.Workbook.Close();
            excelDataWriter.XlApp.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelDataWriter.XlApp);
        }

        /// <summary>
        /// データが既に存在する場合に正しく書き込まれることを確認するテスト
        /// </summary>
        [TestMethod()]
        public void WriteDatasTest1_DataExists()
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
            var settingManager = new SettingManager("../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();
            settings.WriteFilepath = "../../../TestFiles/ExcelDataWriter/DataExists.xlsm";

            var excelDataWriter = new ExcelDataWriter(settings);
            excelDataWriter.XlApp.Visible = true;
            excelDataWriter.Open();
            excelDataWriter.FindTable();

            //テーブルのデータがあるところの一番下をあらかじめ取っておく
            var preLastRange = excelDataWriter.WriteTable.DataBodyRange[1, 3].End(XlDirection.xlDown);
            var preLastRangeRow = preLastRange.Row;
            var preLastEmployeeName = excelDataWriter.Worksheet.Cells(preLastRange.row, excelDataWriter.WriteTable.DataBodyRange[3].Column).Value2.ToString();

            excelDataWriter.WriteDatas(testDatas);

            //書き込み後の行数で比較
            var lastRow = excelDataWriter.WriteTable.DataBodyRange[1, 1].End(XlDirection.xlDown);
            Assert.AreEqual(lastRow.Row, testDatas.Count + preLastRangeRow);
            //書き込み前の最終行がつぶれていないかチェック
            Assert.AreEqual(preLastEmployeeName, preLastRange.Value2.ToString());
            //そのすぐ下がテストデータの1行目かチェック
            Assert.AreEqual(testDatas[0].EmploeeName, excelDataWriter.Worksheet.Cells(preLastRangeRow + 1, excelDataWriter.WriteTable.DataBodyRange[3].Column).Value2.ToString());

            excelDataWriter.XlApp.DisplayAlerts = false;
            excelDataWriter.Workbook.Close();
            excelDataWriter.XlApp.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelDataWriter.XlApp);
        }

        /// <summary>
        /// からのリストが渡されたときにテーブルが変化しないことを確認するテスト
        /// </summary>
        [TestMethod()]
        public void WriteDatasTest1_ListEmpty()
        {
            var testDatas = new List<ExcelData>();
            var settingManager = new SettingManager("../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();
            settings.WriteFilepath = "../../../TestFiles/ExcelDataWriter/DataExists.xlsm";

            var excelDataWriter = new ExcelDataWriter(settings);
            excelDataWriter.XlApp.Visible = true;
            excelDataWriter.Open();
            excelDataWriter.FindTable();

            //テーブルのデータがあるところの一番下をあらかじめ取っておく
            var preLastRange = excelDataWriter.WriteTable.DataBodyRange[1, 3].End(XlDirection.xlDown);
            var preLastRangeRow = preLastRange.Row;
            var preLastEmployeeName = excelDataWriter.Worksheet.Cells(preLastRange.row, excelDataWriter.WriteTable.DataBodyRange[3].Column).Value2.ToString();

            excelDataWriter.WriteDatas(testDatas);

            //書き込み後の行数で比較
            var lastRow = excelDataWriter.WriteTable.DataBodyRange[1, 1].End(XlDirection.xlDown);
            Assert.AreEqual(lastRow.Row, preLastRangeRow);
            //書き込み前の最終行に変化がないかチェック
            Assert.AreEqual(preLastEmployeeName, preLastRange.Value2.ToString());

            excelDataWriter.XlApp.DisplayAlerts = false;
            excelDataWriter.Workbook.Close();
            excelDataWriter.XlApp.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelDataWriter.XlApp);
        }

        /// <summary>
        /// ランダムなやつを正しくソートできたことを確認するテスト
        /// </summary>
        [TestMethod()]
        public void SotrTableTest_RandomData()
        {
            var testData = new List<ExcelData>
            {
                new ExcelData(new DateTime(45284),"茶Ⅱ" ,"御飯"),
                new ExcelData(new DateTime(45285),"茶Ⅰ" ,"ベジータ"),
                new ExcelData(new DateTime(45286),"１Ｐ" ,"チチ"),
                new ExcelData(new DateTime(45287),"１Ｐ" ,"ナム"),
                new ExcelData(new DateTime(45288),"ツイストⅠ" ,"17号"),
                new ExcelData(new DateTime(45289),"3P" ,"ピッコロ"),
                new ExcelData(new DateTime(45290),"１Ｐ" ,"バビディ"),
                new ExcelData(new DateTime(45291),"茶Ⅲ" ,"悟空"),
                new ExcelData(new DateTime(45292),"茶Ⅰ" ,"ミスターポポ"),
                new ExcelData(new DateTime(45293),"3P" ,"ブルマ"),
                new ExcelData(new DateTime(45294),"茶Ⅱ" ,"ランチ"),
                new ExcelData(new DateTime(45295),"ツイストⅠ" ,"サタン"),
                new ExcelData(new DateTime(45296),"3P" ,"ブロリー"),
                new ExcelData(new DateTime(45297),"茶Ⅱ" ,"牛魔王"),
                new ExcelData(new DateTime(45298),"ツイストⅠ" ,"18号"),
                new ExcelData(new DateTime(45299),"１Ｐ" ,"ナッパ"),
                new ExcelData(new DateTime(45300),"茶Ⅲ" ,"天津飯"),
                new ExcelData(new DateTime(45301),"茶Ⅰ" ,"ラディッツ"),
                new ExcelData(new DateTime(45302),"茶Ⅰ" ,"ヤムチャ"),
                new ExcelData(new DateTime(45303),"3P" ,"ビーデル"),
                new ExcelData(new DateTime(45304),"ツイストⅠ" ,"16号"),
                new ExcelData(new DateTime(45305),"ツイストⅠ" ,"クリリン"),
                new ExcelData(new DateTime(45306),"ツイストⅠ" ,"カリン"),
                new ExcelData(new DateTime(45307),"3P" ,"プーアル"),
                new ExcelData(new DateTime(45308),"茶Ⅲ" ,"餃子"),
                new ExcelData(new DateTime(45309),"茶Ⅲ" ,"魔人ブウ"),
                new ExcelData(new DateTime(45310),"3P" ,"ブラ"),
                new ExcelData(new DateTime(45311),"3P" ,"パン"),
                new ExcelData(new DateTime(45312),"茶Ⅱ" ,"亀仙人"),
                new ExcelData(new DateTime(45313),"１Ｐ" ,"ナナチ"),
                new ExcelData(new DateTime(45314),"3P" ,"ピラフ"),
                new ExcelData(new DateTime(45315),"茶Ⅰ" ,"ヤジロベー"),
                new ExcelData(new DateTime(45316),"ツイストⅠ" ,"セル"),
                new ExcelData(new DateTime(45317),"１Ｐ" ,"トランクス"),
                new ExcelData(new DateTime(45318),"１Ｐ" ,"デンデ"),
                new ExcelData(new DateTime(45319),"茶Ⅲ" ,"悟天")
            };
            var settingManager = new SettingManager("../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();
            settings.WriteFilepath = "../../../TestFiles/ExcelDataWriter/Sort_RandomData.xlsm";

            var excelDataWriter = new ExcelDataWriter(settings);
            try
            {
                try
                {
                    excelDataWriter.Open();
                    excelDataWriter.FindTable();
                    excelDataWriter.SortTableAsc();

                    for (int i = 1; true; i++)
                    {
                        if (excelDataWriter.WriteTable.DataBodyRange[i, 1].Value2 == null)
                        {
                            return;
                        }
                        var date = excelDataWriter.WriteTable.DataBodyRange[i, 1].Value2;
                        var lineName = excelDataWriter.WriteTable.DataBodyRange[i, 2].Value2.ToString();
                        var EmploeeName = excelDataWriter.WriteTable.DataBodyRange[i, 3].Value2.ToString();

                        Assert.AreEqual(testData[i - 1].Date, new DateTime((long)date));
                        Assert.AreEqual(testData[i - 1].LineName, lineName.ToString());
                        Assert.AreEqual(testData[i - 1].EmploeeName, EmploeeName.ToString());
                    }
                }
                finally
                {
                    excelDataWriter.XlApp.DisplayAlerts = false;
                    excelDataWriter.XlApp.Quit();
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelDataWriter.XlApp);
            }
        }

        /// <summary>
        /// ソート済みのデータが変化しないことを確認するテスト
        /// </summary>
        [TestMethod()]
        public void SotrTableTest_SortedData()
        {
            var testData = new List<ExcelData>
            {
                new ExcelData(new DateTime(45284),"茶Ⅱ" ,"御飯"),
                new ExcelData(new DateTime(45285),"茶Ⅰ" ,"ベジータ"),
                new ExcelData(new DateTime(45286),"１Ｐ" ,"チチ"),
                new ExcelData(new DateTime(45287),"１Ｐ" ,"ナム"),
                new ExcelData(new DateTime(45288),"ツイストⅠ" ,"17号"),
                new ExcelData(new DateTime(45289),"3P" ,"ピッコロ"),
                new ExcelData(new DateTime(45290),"１Ｐ" ,"バビディ"),
                new ExcelData(new DateTime(45291),"茶Ⅲ" ,"悟空"),
                new ExcelData(new DateTime(45292),"茶Ⅰ" ,"ミスターポポ"),
                new ExcelData(new DateTime(45293),"3P" ,"ブルマ"),
                new ExcelData(new DateTime(45294),"茶Ⅱ" ,"ランチ"),
                new ExcelData(new DateTime(45295),"ツイストⅠ" ,"サタン"),
                new ExcelData(new DateTime(45296),"3P" ,"ブロリー"),
                new ExcelData(new DateTime(45297),"茶Ⅱ" ,"牛魔王"),
                new ExcelData(new DateTime(45298),"ツイストⅠ" ,"18号"),
                new ExcelData(new DateTime(45299),"１Ｐ" ,"ナッパ"),
                new ExcelData(new DateTime(45300),"茶Ⅲ" ,"天津飯"),
                new ExcelData(new DateTime(45301),"茶Ⅰ" ,"ラディッツ"),
                new ExcelData(new DateTime(45302),"茶Ⅰ" ,"ヤムチャ"),
                new ExcelData(new DateTime(45303),"3P" ,"ビーデル"),
                new ExcelData(new DateTime(45304),"ツイストⅠ" ,"16号"),
                new ExcelData(new DateTime(45305),"ツイストⅠ" ,"クリリン"),
                new ExcelData(new DateTime(45306),"ツイストⅠ" ,"カリン"),
                new ExcelData(new DateTime(45307),"3P" ,"プーアル"),
                new ExcelData(new DateTime(45308),"茶Ⅲ" ,"餃子"),
                new ExcelData(new DateTime(45309),"茶Ⅲ" ,"魔人ブウ"),
                new ExcelData(new DateTime(45310),"3P" ,"ブラ"),
                new ExcelData(new DateTime(45311),"3P" ,"パン"),
                new ExcelData(new DateTime(45312),"茶Ⅱ" ,"亀仙人"),
                new ExcelData(new DateTime(45313),"１Ｐ" ,"ナナチ"),
                new ExcelData(new DateTime(45314),"3P" ,"ピラフ"),
                new ExcelData(new DateTime(45315),"茶Ⅰ" ,"ヤジロベー"),
                new ExcelData(new DateTime(45316),"ツイストⅠ" ,"セル"),
                new ExcelData(new DateTime(45317),"１Ｐ" ,"トランクス"),
                new ExcelData(new DateTime(45318),"１Ｐ" ,"デンデ"),
                new ExcelData(new DateTime(45319),"茶Ⅲ" ,"悟天")
            };
            var settingManager = new SettingManager("../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();
            settings.WriteFilepath = "../../../TestFiles/ExcelDataWriter/Sort_SortedData.xlsm";

            var excelDataWriter = new ExcelDataWriter(settings);
            try
            {
                try
                {
                    excelDataWriter.Open();
                    excelDataWriter.FindTable();
                    excelDataWriter.SortTableAsc();

                    for (int i = 1; true; i++)
                    {
                        if (excelDataWriter.WriteTable.DataBodyRange[i, 1].Value2 == null)
                        {
                            return;
                        }
                        var date = excelDataWriter.WriteTable.DataBodyRange[i, 1].Value2;
                        var lineName = excelDataWriter.WriteTable.DataBodyRange[i, 2].Value2.ToString();
                        var EmploeeName = excelDataWriter.WriteTable.DataBodyRange[i, 3].Value2.ToString();

                        Assert.AreEqual(testData[i - 1].Date, new DateTime((long)date));
                        Assert.AreEqual(testData[i - 1].LineName, lineName.ToString());
                        Assert.AreEqual(testData[i - 1].EmploeeName, EmploeeName.ToString());
                    }
                }
                finally
                {
                    excelDataWriter.XlApp.DisplayAlerts = false;
                    excelDataWriter.XlApp.Quit();
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelDataWriter.XlApp);
            }
        }

        /// <summary>
        /// アーカイブのテスト　成功
        /// </summary>
        [TestMethod()]
        public void StoreDataInArchiveTest()
        {
            var settingManager = new SettingManager("../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();
            settings.WriteFilepath = "../../../TestFiles/ExcelDataWriter/StoreDataInArchive.xlsm";

            if (Directory.Exists(settings.ArchiveDirPath))
            {
                Directory.Delete(settings.ArchiveDirPath, true);
            }

            var excelDataWriter = new ExcelDataWriter(settings);
            try
            {
                try
                {
                    excelDataWriter.Open();
                    excelDataWriter.XlApp.Visible = true;
                    excelDataWriter.FindTable();
                    excelDataWriter.StoreDataInArchive();

                    var lastRowValue = excelDataWriter.Worksheet.Range("c2").Value2;
                    Assert.AreEqual("18号", lastRowValue);

                    using (var streamReader = new StreamReader(settings.ArchiveDirPath + "2023.csv"))
                    {
                        var config = new CsvHelper.Configuration.CsvConfiguration(new System.Globalization.CultureInfo("ja-jp", false))
                        {
                            HasHeaderRecord = false
                        };
                        using (var csv = new CsvReader(streamReader, config))
                        {
                            List<ExcelData> val = csv.GetRecords<ExcelData>().ToList<ExcelData>();
                            Assert.AreEqual("牛魔王", val[val.Count - 1].EmploeeName);
                        }
                    }
                }
                finally
                {
                    excelDataWriter.XlApp.DisplayAlerts = false;
                    excelDataWriter.XlApp.Quit();
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelDataWriter.XlApp);

            }
        }

        /// <summary>
        /// アーカイブのテスト　全部アーカイブ
        /// </summary>
        [TestMethod()]
        public void StoreDataInArchiveTest_AllArchive()
        {
            var settingManager = new SettingManager("../../../suppaman45Tests/TestFiles/テスト用UserSettings.json");
            var settings = settingManager.LoadSettings();
            settings.WriteFilepath = "../../../TestFiles/ExcelDataWriter/AllArchive.xlsm";

            if (Directory.Exists(settings.ArchiveDirPath))
            {
                Directory.Delete(settings.ArchiveDirPath, true);
            }

            var excelDataWriter = new ExcelDataWriter(settings);
            try
            {
                try
                {
                    excelDataWriter.Open();
                    excelDataWriter.XlApp.Visible = true;
                    excelDataWriter.FindTable();
                    excelDataWriter.StoreDataInArchive();

                    var lastRowValue = excelDataWriter.Worksheet.Range("b2").Value2;
                    Assert.IsNull(lastRowValue);

                    using (var streamReader = new StreamReader(settings.ArchiveDirPath + "2023.csv"))
                    {
                        var config = new CsvHelper.Configuration.CsvConfiguration(new System.Globalization.CultureInfo("ja-jp", false))
                        {
                            HasHeaderRecord = false
                        };
                        using (var csv = new CsvReader(streamReader, config))
                        {
                            List<ExcelData> val = csv.GetRecords<ExcelData>().ToList<ExcelData>();
                            Assert.AreEqual("ヴェーゼ", val[val.Count - 1].EmploeeName);
                        }
                    }
                }
                finally
                {
                    excelDataWriter.XlApp.DisplayAlerts = false;
                    excelDataWriter.XlApp.Quit();
                }
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelDataWriter.XlApp);


            }
        }
    }
}
