using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Parzan
{
    public class RegexWindowVM : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private List<string> _InvalidPatterns;
        public List<string> InvalidPatterns
        {
            get { return _InvalidPatterns; }
            set
            {
                if (value != _InvalidPatterns)
                {
                    _InvalidPatterns = value;
                    OnPropertyChanged(nameof(InvalidPatterns));
                }
            }
        }
    }
}
