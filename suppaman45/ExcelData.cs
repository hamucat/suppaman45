using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace suppaman45
{

    /// <summary>
    /// 書き出すレコード1行分の値セット
    /// </summary>
    public class ExcelData
    {
        public ExcelData() { }

        public ExcelData(DateTime date, string lineName,string EmploeeName)
        {
            this.Date = date;
            this.LineName = lineName;
            this.EmploeeName = EmploeeName;
        }

        /// <summary>
        /// 日付
        /// </summary>
        [CsvHelper.Configuration.Attributes.Index(0)]
        public DateTime Date { get; set; }

        /// <summary>
        /// ライン名
        /// </summary>
        [CsvHelper.Configuration.Attributes.Index(1)]
        public string LineName { get; set; }

        /// <summary>
        /// 従業員名
        /// </summary>
        [CsvHelper.Configuration.Attributes.Index(2)]
        public string EmploeeName { get; set; }

        public override bool Equals(object obj)
        {
            return obj is ExcelData data &&
                   Date == data.Date &&
                   LineName == data.LineName &&
                   EmploeeName == data.EmploeeName;
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 17;
                hash = hash * 31 + Date.GetHashCode();
                hash = hash * 31 + (LineName != null ? LineName.GetHashCode() : 0);
                hash = hash * 31 + (EmploeeName != null ? EmploeeName.GetHashCode() : 0);
                return hash;
            }
        }
    }
}
