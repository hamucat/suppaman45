using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Diagnostics;
using DocumentFormat.OpenXml.Bibliography;
using System.CodeDom;
using Microsoft.Office.Interop.Excel;

namespace suppaman45
{
    /// <summary>
    /// インポート
    /// </summary>
    public class ExcelDataReader
    {
        UserSettings userSettings;
        NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        public ExcelDataReader(UserSettings settings)
        {
            userSettings = settings ?? throw new ArgumentNullException(nameof(settings));
        }

        /// <summary>
        /// 日付から該当するファイルのパスを取得
        /// </summary>
        /// <param name="date">探す日付</param>
        /// <param name="readDate">読み取る日付</param>
        /// <returns>フルパス</returns>
        /// <exception cref="FileNotFoundException">ファイルが存在しない場合</exception>
        public string GetPath(DateTime readDate)
        {
            var formattedData = readDate.ToString("yyMMdd");
            var searchPattern = formattedData + userSettings.ReadFileExtention;

            string[] matchingFiles = Directory.GetFiles(userSettings.ReadFileDir, searchPattern, SearchOption.AllDirectories);

            if (matchingFiles.Length > 0)
            {
                return matchingFiles[0];
            }
            else
            {
                throw new FileNotFoundException("File not found " + nameof(searchPattern));
            }
        }


        /// <summary>
        /// エクセルを読み込む
        /// </summary>
        /// <param name="path">ファイルのパス</param>
        /// <param name="readDate">読み込み日付</param>
        /// <returns>読み取り結果</returns>
        /// <exception cref="NullReferenceException">指定された名前のシート、名前付き範囲が見つからない場合</exception>
        public List<ExcelData> GetExcelDatas(string path, DateTime readDate)
        {
            return AppendExcelDates(path, readDate, new List<ExcelData>());
        }

        /// <summary>
        /// エクセルを読み込む(追記)
        /// </summary>
        /// <param name="path">ファイルのパス</param>
        /// <param name="readDate">読み込み日付</param>
        /// <param name="excelDatas">追記したいリスト</param>
        /// <returns>読み取り結果</returns>
        /// <exception cref="NullReferenceException">指定された名前のシート、名前付き範囲が見つからない場合</exception>
        public List<ExcelData> GetExcelDatas(string path, DateTime readDate, List<ExcelData> excelDatas)
        {
            return AppendExcelDates(path, readDate, excelDatas);
        }

        private List<ExcelData> AppendExcelDates(string path, DateTime readDate, List<ExcelData> excelDatas)
        {
            //ファイルストリームで渡すとロック中でも開ける
            using (var workbook = new XLWorkbook(new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)))
            {
                var worksheet = workbook.Worksheet(userSettings.ReadSheetName) ?? throw new NullReferenceException("Worhsheet " + nameof(userSettings.ReadSheetName) + " is not found.");
                var namedRange = workbook.NamedRange(userSettings.NamedRange) ?? throw new NullReferenceException("NamedRange " + nameof(userSettings.NamedRange) + " is not found");

                var firstRange = namedRange.Ranges.First();
                var FirstRowNumber = firstRange.FirstCell().Address.RowNumber;

                var usedCells = namedRange.Ranges.CellsUsed();

                //昼交代から始まるケースで当該部分を除外するためのフラグ
                var isKoutai = true;

                //縦回転
                for (int i = FirstRowNumber + 1; i < worksheet.LastRowUsed().RowNumber(); i++)
                {
                    //昼交代っぽい行を除外する
                    //横一列みて名前の数が閾値以下ならその行を飛ばす
                    var first = worksheet.Cell(i, firstRange.FirstCell().Address.ColumnNumber).Address;
                    var last = worksheet.Cell(i, namedRange.Ranges.Last().LastCell().Address.ColumnNumber).Address;
                    var cells = worksheet.Range(first, last);
                    if (isKoutai == true && cells.CellsUsed().Count() <= userSettings.ReadIgnoreThrethold)
                    {
                        logger.Trace("昼交代continue i={0}", i);
                        continue;
                    }
                    else
                    {
                        isKoutai = false;
                    }

                    //横回転
                    foreach (var item in usedCells)
                    {
                        //行列が交わるところがnullなら飛ばす
                        var activeCellValue = worksheet.Cell(i, item.Address.ColumnNumber).Value.ToString();
                        if (string.IsNullOrEmpty(activeCellValue))
                        {
                            continue;
                        }

                        var excelData = new ExcelData();
                        excelData.Date = readDate;
                        excelData.LineName = item.Value.ToString();
                        excelData.EmploeeName = ReplaceEmploeeName(activeCellValue);

                        if (!IsDuplicate(readDate, excelData, excelDatas) &&
                            !IsValidEmploeeName(excelData.EmploeeName) &&
                            !IsLineName(worksheet.Cell(FirstRowNumber, item.Address.ColumnNumber).Value.ToString(), excelData.EmploeeName)
                        )
                        {
                            excelDatas.Add(excelData);
                            logger.Trace("Add:{0},{1}.{2}", excelData.Date.ToString("yyMMdd"), excelData.LineName, excelData.EmploeeName);
                        }
                    }//foreach
                }//for
            }//using

            return excelDatas;
        }

        //ReplacePatternsに従って表記ゆれを置換
        private string ReplaceEmploeeName(string emploeeName)
        {
            foreach (var item in userSettings.ReplacePatterns)
            {
                if (item.Key == emploeeName.Trim())
                {
                    return item.Value;
                }
            }
            return emploeeName;
        }

        //重複する名前を除外する判定 当日分に限る
        private bool IsDuplicate(DateTime readDate, ExcelData newData, List<ExcelData> existingDatas)
        {
            foreach (var item in existingDatas)
            {
                if (readDate == item.Date && newData.EmploeeName == item.EmploeeName)
                {
                    logger.Trace("IsDuplicate true {0} {1}", readDate.ToString("yyMMdd"), newData.EmploeeName);
                    return true;
                }
            }
            return false;
        }

        //正規表現パターンにマッチする名前を除外する判定
        private bool IsValidEmploeeName(string emploeeName)
        {
            foreach (var item in userSettings.InvalidPatterns)
            {
                if (System.Text.RegularExpressions.Regex.IsMatch(emploeeName, item))
                {
                    logger.Trace("Regex true {0} = {1}", emploeeName, item);
                    return true;
                }
            }
            return false;
        }

        //ライン名を除外する判定
        private bool IsLineName(string columnHead, string emploeeName)
        {
            if (columnHead == emploeeName)
            {
                logger.Trace("LineName true {0}", emploeeName);
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// データがない日付リストを取得する
        /// </summary>
        /// <returns>日付のリスト</returns>
        /// <exception cref="FileNotFoundException">書き込み用ファイルが見つからない場合</exception>
        /// <exception cref="NullReferenceException">名前付き範囲が見つからない場合</exception>
        public List<DateTime> GetUnprocessedDateList()
        {
            var unprocessedDates = new List<DateTime>();

            if (!File.Exists(userSettings.WriteFilepath))
            {
                throw new FileNotFoundException(userSettings.WriteFilepath + "is not found.");
            }

            var xlapp = new Microsoft.Office.Interop.Excel.Application();
            try
            {
                Workbook workbook = xlapp.Workbooks.Open(Path.GetFullPath(userSettings.WriteFilepath));
                dynamic worksheet = workbook.Sheets[userSettings.ManageSheetName];

                if (!IsNamedRangeExists(worksheet, userSettings.ManageSheetName + "!" + userSettings.UnprocessedDatesRangeName))
                {
                    throw new NullReferenceException();
                }
                var namedRange = worksheet.Range(userSettings.UnprocessedDatesRangeName);
                var firstRow = namedRange.Row;
                var firstColumn = namedRange.Column;


                var dateCell = worksheet.Cells(firstRow, firstColumn);
                var date = new DateTime();
                var i = 0;

                while (dateCell.Value2 != null && DateTime.TryParse(dateCell.Value.ToString(), out date))
                {
                    if (worksheet.Cells(firstRow + i, firstColumn + 1).Value2 == "ない")
                    {
                        unprocessedDates.Add(date);
                    }
                    i++;
                    dateCell = worksheet.Cells(firstRow + i, firstColumn);
                }
            }
            finally
            {
                xlapp.DisplayAlerts = false;
                xlapp.Quit();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlapp);
            }
            return unprocessedDates;
        }

        private bool IsNamedRangeExists(dynamic worksheet, string name)
        {
            // Namesコレクション内の名前付き範囲を調べる
            foreach (Microsoft.Office.Interop.Excel.Name item in worksheet.Names)
            {
                if (item.NameLocal == name)
                {
                    return true;
                }
            }
            return false;
        }
    }
}