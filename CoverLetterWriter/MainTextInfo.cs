using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoverLetterWriter
{
    public class MainTextInfo : INotifyPropertyChanged
    {
        public int Id { get; set; }
        private string _title;
        public string Title
        {
            get { return _title; }
            set
            {
                if (_title != value)
                {
                    _title = value;
                    OnPropertyChanged(nameof(Title), Title);
                }
            }
        }
        private string _mainText;
        public string MainText
        {
            get { return _mainText; }
            set
            {
                if (_mainText != value)
                {
                    _mainText = value;
                    OnPropertyChanged(nameof(MainText), MainText);
                }
            }
        }
        
        public event PropertyChangedEventHandler PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName, string value = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
