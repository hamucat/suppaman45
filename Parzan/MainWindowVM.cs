using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;

namespace Parzan
{
    public class MainWindowVM : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        /// <summary>
        /// 読み込むファイルを探すディレクトリのパス
        /// </summary>
        private string _ReadFileDir;
        public string ReadFileDir
        {
            get { return _ReadFileDir; }
            set
            {
                if (_ReadFileDir != value)
                {
                    _ReadFileDir = value;
                    OnPropertyChanged(nameof(ReadFileDir));
                }
            }
        }

        /// <summary>
        /// 読み込むファイルの拡張子
        /// </summary>
        private string _ReadFileExtention;
        public string ReadFileExtention
        {
            get { return _ReadFileExtention; }
            set
            {
                if (_ReadFileExtention != value)
                {
                    _ReadFileExtention = value;
                    OnPropertyChanged(nameof(ReadFileExtention));
                }
            }
        }

        /// <summary>
        /// 読み込むシート名
        /// </summary>
        private string _ReadSheetName;
        public string ReadSheetName
        {
            get { return _ReadSheetName; }
            set
            {
                if (_ReadSheetName != value)
                {
                    _ReadSheetName = value;
                    OnPropertyChanged(nameof(ReadSheetName));
                }
            }

        }

        /// <summary>
        /// ライン構成テーブルの見出し行範囲
        /// </summary>
        private string _NamedRange;
        public string Namedrange
        {
            get { return _NamedRange; }
            set
            {
                if (_NamedRange != value)
                {
                    _NamedRange = value;
                    OnPropertyChanged(nameof(Namedrange));
                }
            }
        }

        /// <summary>
        /// 読み込みを除外する行の閾値
        /// </summary>
        private int _ReadIgnoreThreshold;
        public int ReadIgnoreThreshold
        {
            get { return _ReadIgnoreThreshold; }
            set
            {
                if (_ReadIgnoreThreshold != value)
                {
                    _ReadIgnoreThreshold = value;
                    OnPropertyChanged(nameof(ReadIgnoreThreshold));
                }
            }
        }

        /// <summary>
        /// 書き込みファイルのパス
        /// </summary>
        private string _WriteFilePath;
        public string WriteFilePath
        {
            get { return _WriteFilePath; }
            set
            {
                if (_WriteFilePath != value)
                {
                    _WriteFilePath = value;
                    OnPropertyChanged(nameof(WriteFilePath));
                }
            }
        }

        /// <summary>
        /// 書き込みシート名
        /// </summary>
        private string _WriteSheetName;
        public string WriteSheetName
        {
            get { return _WriteSheetName; }
            set
            {
                if (_WriteSheetName != value)
                {
                    _WriteSheetName = value;
                    OnPropertyChanged($"{nameof(WriteSheetName)}");
                }
            }
        }

        /// <summary>
        /// 書き込みテーブル名
        /// </summary>
        private string _WriteTableName;
        public string WriteTableName
        {
            get { return _WriteTableName; }
            set
            {
                if (_WriteTableName != value)
                {
                    _WriteTableName = value;
                    OnPropertyChanged(nameof(WriteTableName));
                }
            }
        }

        /// <summary>
        /// アーカイブファイルのディレクトリ
        /// </summary>
        private string _ArchiveDirPath;
        public string ArchiveDirPath
        {
            get { return _ArchiveDirPath; }
            set
            {
                if (_ArchiveDirPath != value)
                {
                    _ArchiveDirPath = value;
                    OnPropertyChanged(nameof(_ArchiveDirPath));
                }
            }
        }

        /// <summary>
        /// 書き込みファイルの管理用シート名
        /// </summary>
        private string _ManageSheetName;
        public string ManageSheetName
        {
            get { return _ManageSheetName; }
            set
            {
                if (_ManageSheetName != value)
                {
                    _ManageSheetName = value;
                    OnPropertyChanged(nameof(ManageSheetName));
                }
            }
        }

        /// <summary>
        /// 未処理の日付が並んでいる範囲の名前付き範囲
        /// </summary>
        private string _UnprocessedDatesRangeName;
        public string UnprocessedDatesRangeName
        {
            get { return _UnprocessedDatesRangeName; }
            set
            {
                if (_UnprocessedDatesRangeName != value)
                {
                    _UnprocessedDatesRangeName = value;
                    OnPropertyChanged(nameof(UnprocessedDatesRangeName));
                }
            }
        }
    }
}
