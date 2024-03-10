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

        //読み込むファイルを探すディレクトリのパス
        private string _readFileDir;
        public string ReadFileDir
        {
            get { return _readFileDir; }
            set
            {
                if(_readFileDir != value)
                {
                    _readFileDir = value;
                    OnPropertyChanged(nameof(ReadFileDir));
                }
            }
        }
    }
}
