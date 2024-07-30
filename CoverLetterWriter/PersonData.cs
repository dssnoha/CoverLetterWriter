using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoverLetterWriter
{
    public class PersonData : INotifyPropertyChanged
    {
        public int Id { get; set; }
        private string _fullName;
        public string FullName
        {
            get { return _fullName; }
            set
            {
                if (_fullName != value)
                {
                    _fullName = value;
                    OnPropertyChanged(nameof(FullName), FullName);
                }
            }
        }
        private string _streetAddress;
        public string StreetAddress
        {
            get { return _streetAddress; }
            set
            {
                if (_streetAddress != value)
                {
                    _streetAddress = value;
                    OnPropertyChanged(nameof(StreetAddress), StreetAddress);
                }
            }
        }
        private string _cityAndPostcode;
        public string CityAndPostcode
        {
            get { return _cityAndPostcode; }
            set
            {
                if (_cityAndPostcode != value)
                {
                    _cityAndPostcode = value;
                    OnPropertyChanged(nameof(CityAndPostcode), CityAndPostcode);
                }
            }
        }
        private string _phoneNumber;
        public string PhoneNumber
        {
            get { return _phoneNumber; }
            set
            {
                if (_phoneNumber != value)
                {
                    _phoneNumber = value;
                    OnPropertyChanged(nameof(PhoneNumber), PhoneNumber);
                }
            }
        }
        private string _email;
        public string Email
        {
            get { return _email; }
            set
            {
                if (_email != value)
                {
                    _email = value;
                    OnPropertyChanged(nameof(Email), Email);
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
