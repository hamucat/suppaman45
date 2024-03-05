using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace suppaman45
{
    public class ArchiveManager
    {
        UserSettings userSettings;
        NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        public ArchiveManager(UserSettings userSettings)
        {
            this.userSettings = userSettings;
            logger.Debug("ArchiveManager initiarize");
        }

        public void WriteToCSV(List<ExcelData> datas)
        {
            if (datas.Count == 0)
            {
                return;
            }
            var yarlyDataList = SplitByYear(datas);
            Directory.CreateDirectory(userSettings.ArchiveDirPath);

            foreach (var item in yarlyDataList)
            {
                using (var streamWriter = new StreamWriter(userSettings.ArchiveDirPath + item[0].Date.Year.ToString() + ".csv", true, System.Text.Encoding.UTF8))
                {
                    var config = new CsvHelper.Configuration.CsvConfiguration(new System.Globalization.CultureInfo("ja-jp", false));
                    config.HasHeaderRecord = false;

                    using (var csv = new CsvHelper.CsvWriter(streamWriter, config))
                    {
                        csv.WriteRecords(item);
                    }
                }
            }


        }

        /// <summary>
        /// リスト内のデータを年ごとのリストに分割
        /// </summary>
        /// <param name="datas">分割したいリスト</param>
        /// <returns>リスト＜分割したリスト＞</returns>
        public List<List<ExcelData>> SplitByYear(List<ExcelData> datas)
        {
            var yearlyDataList = new List<List<ExcelData>>();

            var currentYearDataList = new List<ExcelData>();
            currentYearDataList.Add(datas[0]);
            for (var i = 1; i < datas.Count; i++)
            {
                if (datas[i].Date.Year == currentYearDataList[0].Date.Year)
                {
                    currentYearDataList.Add(datas[i]);
                }
                else
                {
                    yearlyDataList.Add(currentYearDataList);
                    currentYearDataList = new List<ExcelData> { datas[i] };
                }
            }
            yearlyDataList.Add(currentYearDataList);

            return yearlyDataList;
        }
    }
}
