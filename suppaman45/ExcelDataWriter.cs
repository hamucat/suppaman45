using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.ComponentModel.DataAnnotations.Schema;
using System.Diagnostics;
using NLog;

namespace suppaman45
{
    public class ExcelDataWriter
    {
        //テストメソッドから殺せるようにpublicにする
        public Microsoft.Office.Interop.Excel.Application XlApp { get; private set; }

        public Workbook Workbook { get; private set; }
        public dynamic Worksheet { get; private set; }
        public ListObject WriteTable { get; private set; }

        UserSettings settings;

        NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

#pragma warning disable CS8618 // null 非許容のフィールドには、コンストラクターの終了時に null 以外の値が入っていなければなりません。Null 許容として宣言することをご検討ください。
        public ExcelDataWriter(UserSettings settings)
#pragma warning restore CS8618 // null 非許容のフィールドには、コンストラクターの終了時に null 以外の値が入っていなければなりません。Null 許容として宣言することをご検討ください。
        {
            XlApp = new Microsoft.Office.Interop.Excel.Application();
            this.settings = settings;
        }

        /// <summary>
        /// ファイルとワークシートを開く
        /// </summary>
        /// <exception cref="FileNotFoundException">ファイルが見つからなかった場合</exception>
        /// <exception cref="IOException">指定された名前のシートが見つからなかった場合</exception>
        /// <exception cref="System.Runtime.InteropServices.COMException"></exception>
        public void Open()
        {
            //ファイルが存在するかチェック
            if (!File.Exists(settings.WriteFilepath))
            {
                throw new FileNotFoundException("{0} is not found.", settings.WriteFilepath);
            }

            //ファイルが誰かに開かれているかチェック
            if (File.Exists(Path.GetDirectoryName(settings.WriteFilepath) + "~$" + Path.GetFileName(settings.WriteFilepath)))
            {
                throw new IOException(settings.WriteFilepath + " is locked.");
            }

            Workbook = XlApp.Workbooks.Open(Path.GetFullPath(settings.WriteFilepath));

            //sheet
            try
            {
                Worksheet = Workbook.Sheets[settings.WriteSheetname];
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                throw new System.Runtime.InteropServices.COMException("Worksheet" + settings.WriteSheetname + "is not found.");
            }
        }

        /// <summary>
        /// ListObjectの中から書き込み用のテーブルを探す
        /// </summary>
        /// <exception cref="System.Runtime.InteropServices.COMException">テーブルが見つからなかった場合</exception>
        public void FindTable()
        {
            foreach (ListObject item in Worksheet.ListObjects)
            {
                if (item.Name == settings.WriteTableName)
                {
                    WriteTable = item;
                    return;
                }
            }

            throw new System.Runtime.InteropServices.COMException("Table " + settings.WriteTableName + " is not found.");
        }

        /// <summary>
        /// データを書き込む
        /// </summary>
        /// <param name="excelDatas">書き込み用のデータリスト</param>
        public void WriteDatas(List<ExcelData> excelDatas)
        {
            var startRow = GetStartRow();
            var startRange = WriteTable.DataBodyRange[1, 1];
            var startCol = startRange.Column;

            for (int i = 0; i < excelDatas.Count; i++)
            {
                Worksheet.Cells(startRow + i, startCol).Value = excelDatas[i].Date.ToString("yyyy/MM/dd");
                Worksheet.Cells(startRow + i, startCol + 1).Value = excelDatas[i].LineName;
                Worksheet.Cells(startRow + i, startCol + 2).Value = excelDatas[i].EmploeeName;

                logger.Trace("Write: {0},{1},{2}", excelDatas[i].Date.ToString("yyyy/MM/dd"), excelDatas[i].LineName, excelDatas[i].EmploeeName);
            }
        }

        /// <summary>
        /// 開始行番号を取得
        /// </summary>
        /// <returns>開始行番号</returns>
        private int GetStartRow()
        {
            var firstRange = WriteTable.DataBodyRange[1, 1];

            //テーブルが空の場合
            //Value2の中身はnullかdoubleかstring　doubleはnullでない　stringは””かもしれない
            if (firstRange.Value2 is null || firstRange.Value2 is string && !string.IsNullOrEmpty(firstRange.Value2))
            {
                logger.Debug("テーブルが空 GetStartRow() is {0}", firstRange.Row);
                return firstRange.Row;
            }
            else
            {
                var lastRange = firstRange.End(XlDirection.xlDown);
                logger.Debug("データがある GetStartRow() is {0}", lastRange.Row + 1);
                return lastRange.Row + 1;
            }
        }

        /// <summary>
        /// テーブルを昇順にソート　ソート列はハードコード
        /// </summary>
        public void SortTableAsc()
        {
            WriteTable.Sort.SortFields.Clear();
            WriteTable.Sort.SortFields.Add(WriteTable.HeaderRowRange[1], XlSortOn.xlSortOnValues, XlSortOrder.xlAscending, null, XlSortDataOption.xlSortNormal);
            WriteTable.Sort.SortFields.Add(WriteTable.HeaderRowRange[2], XlSortOn.xlSortOnValues, XlSortOrder.xlAscending, null, XlSortDataOption.xlSortNormal);
            WriteTable.Sort.Apply();

            Workbook.RefreshAll();
        }

        /// <summary>
        /// 古いデータをアーカイブするやつ
        /// </summary>
        public void StoreDataInArchive()
        {
            //テーブルが空なら何もしない
            if (WriteTable.DataBodyRange[1, 1].Value2 is null)
            {
                logger.Debug("テーブルが空");
                return;
            }

            //日付昇順でソート
            WriteTable.Sort.SortFields.Clear();
            WriteTable.Sort.SortFields.Add(WriteTable.DataBodyRange[1], XlSortOn.xlSortOnValues, XlSortOrder.xlAscending, null, XlSortDataOption.xlSortNormal);
            WriteTable.Sort.Apply();

            var minimumDate = GetMinimumDate();
            var (oldData, rowNumber) = GetOldDatasAndRowNumber(minimumDate);

            //アーカイブするデータがない時には何もしない
            if (oldData.Count == 0)
            {
                return;
            }

            var archiveManager = new ArchiveManager(settings);
            archiveManager.WriteToCSV(oldData);

            //なぜか－1しないと1行余計に持っていかれる テーブル下げても追尾する
            var deleteRange = Worksheet.Range(WriteTable.DataBodyRange[1, 1], WriteTable.DataBodyRange[rowNumber - 1, 3]);
            deleteRange.EntireRow.Delete();
        }

        /// <summary>
        /// 管理シートの日付一覧の一番下を取得
        /// </summary>
        /// <returns></returns>
        private DateTime GetMinimumDate()
        {
            var manageSheet = Workbook.Sheets[settings.ManageSheetName];
            var namedRange = manageSheet.Range(settings.UnprocessedDatesRangeName);
            var firstRow = namedRange.Row();
            var firstColumn = namedRange.Column();

            var dateCell = manageSheet.Cells(firstRow, firstColumn);
            var date = new DateTime(1900, 1, 1);
            var i = 0;

            //日付じゃなくなる一番下までdateに値入れながらループ　TryParseに失敗したらwhileから抜けるので例外処理は不要
            while (dateCell.Value2 != null && DateTime.TryParse(dateCell.Value.ToString(), out date))
            {
                i++;
                dateCell = manageSheet.Cells(firstRow + i, firstColumn);
            }

            return date;
        }

       　/// <summary>
        /// 閾値より古いの日付のデータをアーカイブ用リストに格納
        /// </summary>
        /// <param name="minimumDate">閾値の日付</param>
        /// <returns>アーカイブするデータのリスト</returns>
        /// <exception cref="System.FormatException">Recordsの日付をDateTime型に変換できなかった時</exception>
        private (List<ExcelData> archiveData, int row) GetOldDatasAndRowNumber(DateTime minimumDate)
        {
            //昇順にソート済みのテーブルを上から走査して指定日を超える最初の行を見つけたところで脱出
            var currentDate = new DateTime();
            var row = 1;
            while (WriteTable.DataBodyRange[row, 1].Value != null && DateTime.TryParse(WriteTable.DataBodyRange[row, 1].Value.ToString(), out currentDate))
            {
                if (currentDate >= minimumDate)
                {
                    break;
                }
                row++;
            }

            logger.Debug("指定日を超える最初の行 = {0}", row);
            
            //新しいデータしかなかった時
            if (row == 1)
            {
                return (new List<ExcelData>(), 1);
            }

            //アーカイブするデータをリストに格納
            var archiveData = new List<ExcelData>();
            for (int i = 1; i < row; i++)
            {
                try
                {
                    var excelData = new ExcelData();
                    excelData.Date = DateTime.Parse(WriteTable.DataBodyRange[i, 1].Value.ToString());
                    excelData.LineName = WriteTable.DataBodyRange[i, 2].Value2.ToString();
                    excelData.EmploeeName = WriteTable.DataBodyRange[i, 3].Value2.ToString();

                    archiveData.Add(excelData);
                }
                catch (System.FormatException e)
                {
                    throw new System.FormatException(WriteTable.DataBodyRange[i, 1].Address.ToString() + " can't cast DateTime.", e);
                }
            }

            return (archiveData, row);
        }
    }
}
