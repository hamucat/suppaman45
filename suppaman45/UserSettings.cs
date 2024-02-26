using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace suppaman45
{
    /// <summary>
    /// 設定を保持する
    /// </summary>
    [Serializable]
    public class UserSettings
    {
        /// <summary>
        /// コンストラクタが既定値を設定
        /// </summary>
        public UserSettings()
        {
            ReadFileDir = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal) + @"\生産指示\";
            ReadFileExtention = ".xlsm";
            ReadSheetName = "ライン構成";
            NamedRange = "memTableHead";
            ReadIgnoreThrethold = 3;
            WriteFilepath = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal) + @"\統計情報.xlsm";
            WriteSheetname = "Records";
            WriteTableName = "テーブル1";
            ArchiveDirPath = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal) + @"\アーカイブ\";
            ManageSheetName = "管理";
            UnprocessedDatesRangeName = "UnprocessedDatesRange";
            WaitTime = 5000;
            InvalidPatterns = new List<string>();
            ReplacePatterns = new Dictionary<string, string>();
        }

        /// <summary>
        /// リストはコンストラクタでAddするとロード時にそこに追記されて２重に値が存在してしまうので外から呼び出す感じにしてある。
        /// </summary>
        public void SetDefaultValues()
        {
            InvalidPatterns.Add("[\\x00 -\\x7F]");
            InvalidPatterns.Add("^[ぁ-ゞ]+$");
            InvalidPatterns.Add("[（）、・]");
        }

        public override bool Equals(object obj)
        {
            return obj is UserSettings settings &&
                   ReadFileDir == settings.ReadFileDir &&
                   ReadFileExtention == settings.ReadFileExtention &&
                   ReadSheetName == settings.ReadSheetName &&
                   NamedRange == settings.NamedRange &&
                   ReadIgnoreThrethold == settings.ReadIgnoreThrethold &&
                   WriteFilepath == settings.WriteFilepath &&
                   WriteSheetname == settings.WriteSheetname &&
                   WriteTableName == settings.WriteTableName &&
                   ArchiveDirPath == settings.ArchiveDirPath &&
                   ManageSheetName == settings.ManageSheetName &&
                   UnprocessedDatesRangeName == settings.UnprocessedDatesRangeName &&
                   WaitTime == settings.WaitTime &&
                   EqualityComparer<List<string>>.Default.Equals(InvalidPatterns, settings.InvalidPatterns) &&
                   EqualityComparer<Dictionary<string, string>>.Default.Equals(ReplacePatterns, settings.ReplacePatterns);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                int hash = 17;
                hash = hash * 31 + (ReadFileDir != null ? ReadFileDir.GetHashCode() : 0);
                hash = hash * 31 + (ReadFileExtention != null ? ReadFileExtention.GetHashCode() : 0);
                hash = hash * 31 + (ReadSheetName != null ? ReadSheetName.GetHashCode() : 0);
                hash = hash * 31 + (NamedRange != null ? NamedRange.GetHashCode() : 0);
                hash = hash * 31 + ReadIgnoreThrethold;
                hash = hash * 31 + (WriteFilepath != null ? WriteFilepath.GetHashCode() : 0);
                hash = hash * 31 + (WriteSheetname != null ? WriteSheetname.GetHashCode() : 0);
                hash = hash * 31 + (WriteTableName != null ? WriteTableName.GetHashCode() : 0);
                hash = hash * 31 + (ArchiveDirPath != null ? ArchiveDirPath.GetHashCode() : 0);
                hash = hash * 31 + (ManageSheetName != null ? ManageSheetName.GetHashCode() : 0);
                hash = hash * 31 + (UnprocessedDatesRangeName != null ? UnprocessedDatesRangeName.GetHashCode() : 0);
                hash = hash * 31 + WaitTime;
                hash = hash * 31 + (InvalidPatterns != null ? InvalidPatterns.GetHashCode() : 0);
                hash = hash * 31 + (ReplacePatterns != null ? ReplacePatterns.GetHashCode() : 0);
                return hash;
            }
        }

        /// <summary>
        /// 読み込むファイルを探すディレクトリのパス
        /// </summary>
        public string ReadFileDir { get; set; }

        /// <summary>
        /// 読み込むファイルの拡張子
        /// </summary>
        public string ReadFileExtention { get; set; }

        /// <summary>
        /// 読み込むシート名
        /// </summary>
        public string ReadSheetName { get; set; }

        /// <summary>
        /// ライン構成テーブルの見出し行範囲
        /// </summary>
        public string NamedRange { get; set; }

        /// <summary>
        /// 読み込みを除外する行の値数の閾値（昼交代っぽい行を除外するため）
        /// </summary>
        public int ReadIgnoreThrethold { get; set; }

        /// <summary>
        /// 書き込みファイルのパス
        /// </summary>
        public string WriteFilepath { get; set; }

        /// <summary>
        /// 書き込みシート名
        /// </summary>
        public string WriteSheetname { get; set; }

        /// <summary>
        /// 書き込みテーブル名
        /// </summary>
        public string WriteTableName { get; set; }

        /// <summary>
        /// アーカイブファイルのディレクトリ
        /// </summary>
        public string ArchiveDirPath { get; set; }

        /// <summary>
        /// 書き込みファイルの管理用シート名
        /// </summary>
        public string ManageSheetName { get; set; }

        /// <summary>
        /// 未処理の日付リストが並んでいる範囲の名前付き範囲
        /// </summary>
        public string UnprocessedDatesRangeName { get; set; }

        /// <summary>
        /// 自動巡回の待ち時間
        /// </summary>
        public int WaitTime { get; set; }

        /// <summary>
        /// 除外する正規表現のリスト
        /// </summary>
        public List<string> InvalidPatterns { get; set; }

        /// <summary>
        /// 表記ゆれの置換
        /// </summary>
        public Dictionary<string, string> ReplacePatterns { get; set; }
    }
}
