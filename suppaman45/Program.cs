using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace suppaman45
{
    class Program
    {
        static void Main(string[] args)
        {
            NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
            logger.Info("====================巡回開始===================");

            var settingManager = new SettingManager("settings.Json");
            var userSettings = settingManager.LoadSettings();

            var excelDataReader = new ExcelDataReader(userSettings);

            //ない日を取得
            logger.Info("ない日を取得");
            var fetchDates = excelDataReader.GetUnprocessedDateList();

            //読み込み
            var ExcelDatas = new List<ExcelData>();
            foreach (var item in fetchDates)
            {
                try
                {
                    logger.Info("{0} の読み込み開始", item.ToString());
                    ExcelDatas = excelDataReader.GetExcelDatas(excelDataReader.GetPath(item), item, ExcelDatas);
                    logger.Info("完了 {0} 件(累計）", ExcelDatas.Count);
                }
                //ファイルがない
                catch (FileNotFoundException ex) { logger.Error(ex); }
                //シートがない
                catch (NullReferenceException ex) { logger.Error(ex); }
            }

            var excelDataWriter = new ExcelDataWriter(userSettings);

            try
            {
                logger.Info("書き込み用ファイルを開いています");
                excelDataWriter.Open();
                excelDataWriter.XlApp.Visible = false;
                try
                {
                    logger.Info("テーブルを探しています");
                    excelDataWriter.FindTable();
                    logger.Info("データを書き込んでいます");
                    excelDataWriter.WriteDatas(ExcelDatas);
                    logger.Info("古いデータをアーカイブに移しています");
                    excelDataWriter.StoreDataInArchive();
                    logger.Info("ソートしています");
                    excelDataWriter.SortTableAsc();
                    logger.Info("保存しています");
                    excelDataWriter.Workbook.Save();
                }
                finally
                {
                    logger.Info("閉じています");
                    excelDataWriter.XlApp.DisplayAlerts = false;
                    excelDataWriter.XlApp.Quit();
                }
            }
            //ファイルがない
            catch (FileNotFoundException ex) { logger.Error(ex); }
            //ファイルがつかまれてる
            catch (IOException ex) { logger.Error(ex); }
            //テーブルまたはシートがない
            catch (System.Runtime.InteropServices.COMException ex) { logger.Error(ex); }
            finally
            {
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelDataWriter.XlApp);
            }
            logger.Info("完了");


        }
    }
}
