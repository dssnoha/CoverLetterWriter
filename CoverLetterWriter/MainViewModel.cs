using DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media.Media3D;
using System.Windows.Threading;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace CoverLetterWriter
{
    public class MainViewModel : INotifyPropertyChanged
    {
        
        private string _companyName;
        public string CompanyName
        {
            get { return _companyName; }
            set
            {
                if (_companyName != value)
                {
                    _companyName = value;
                    OnPropertyChanged(nameof(CompanyName), CompanyName);
                }
            }
        }
        private string _companyAdresse;
        public string CompanyStreetAddress
        {
            get { return _companyAdresse; }
            set
            {
                
                if (_companyAdresse != value)
                {
                    _companyAdresse = value;
                    OnPropertyChanged(nameof(CompanyStreetAddress), CompanyStreetAddress);
                }
            }
        }
        private string _companyCountry;
        public string CompanyCityAndPostcode
        {
            get { return _companyCountry; }
            set
            {
                if (_companyCountry != value)
                {
                    _companyCountry = value;
                    OnPropertyChanged(nameof(CompanyCityAndPostcode), CompanyCityAndPostcode);
                }
            }
        }
        private MainTextInfo _mainText; 
        public MainTextInfo MainText
        {
            get { return _mainText; }
            set
            {
                if (value != _mainText)
                {
                    _mainText = value;
                    OnPropertyChanged(nameof(MainText), MainText.MainText);
                }
            }
        }
        private DateTime _date;
        public DateTime Date
        {
            get { return _date; }
            set
            {
                if (_date != value)
                {
                    _date = value;
                    CultureInfo germanCulture = new CultureInfo("de-DE");
                    OnPropertyChanged(nameof(Date), Date.ToString("dd.MMMM yyyy", germanCulture));
                }
            }
        }
        private string _day;
        public string Day
        {
            get { return _day; }
            set
            {
                
                if (_day != value)
                {
                    _day = value;
                    OnPropertyChanged(nameof(Day));
                }
            }
        }
        private string _month;
        public string Month
        {
            get { return _month; }
            set
            {
               
                if (_month != value)
                {
                    _month = value;
                    OnPropertyChanged(nameof(Month));
                }
            }
        }
        private string _year;
        public string Year
        {
            get { return _year; }
            set
            {
                
                if (_year != value)
                {
                    _year = value;
                    OnPropertyChanged(nameof(Year), Year);
                }
            }
        }
     
        private string _name;
        public string Name
        {
            get { return _name; }
            set
            {
                if (_name != value)
                {
                    _name = value;
                    OnPropertyChanged(nameof(Name));
                }
            }
        }
        private string _positionName;
        public string PositionName
        {
            get { return _positionName; }
            set
            {
                if (_positionName != value)
                {
                    _positionName = value;
                    OnPropertyChanged(nameof(PositionName), PositionName);
                }
            }
        }
        private string _positionFullName;
        public string PositionFullName
        {
            get { return _positionFullName; }
            set
            {
                if (_positionFullName != value)
                {
                    _positionFullName = value;
                    OnPropertyChanged(nameof(PositionFullName), PositionFullName);
                }
            }
        }
        
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
        private string _newMainText;
        public string NewMainText
        {
            get { return _newMainText; }
            set
            {
                if (_newMainText != value)
                {
                    _newMainText = value;
                    OnPropertyChanged(nameof(NewMainText), NewMainText);
                }
            }
        }
        private string _newTitle;
        public string NewTitle
        {
            get { return _newTitle; }
            set
            {
                if (_newTitle != value)
                {
                    _newTitle = value;
                    OnPropertyChanged(nameof(NewTitle), NewTitle);
                }
            }
        }
        private List<string> _options;
        public List<string> Options
        {
            get { return _options; }
            set
            {
                if (_options != value)
                {
                    _options = value;
                    OnPropertyChanged(nameof(Options));
                }
            }
        }
        private string _selectedOption;
        public string SelectedOption
        {
            get { return _selectedOption; }
            set
            {
                if (_selectedOption != value)
                {
                    switch (value)
                    {
                        case "Unkown":
                            _selectedOption = "geehrte Damen und Herren";
                            break;
                        case "Mr.":
                            _selectedOption = "geehrter Herr";
                            break;
                        case "Miss":
                            _selectedOption = "geehrte Frau";
                            break;
                    }           
                    OnPropertyChanged(nameof(SelectedOption), SelectedOption);
                }
            }
        }
        private ObservableCollection<PersonData> _personOptions;
        public ObservableCollection<PersonData> PersonOptions
        {
            get { return _personOptions; }
            set
            {
                    _personOptions = value;
                    OnPropertyChanged(nameof(PersonOptions));
            }
        }
        private Visibility _imgVis; 
        public Visibility ImgVis
        {
            get { return _imgVis; }
            set
            {
                if (value != _imgVis)
                {
                    _imgVis = value;
                    OnPropertyChanged(nameof(ImgVis));
                }
            }     
        }
        private Visibility _btnVis;
        public Visibility BtnVis
        {
            get { return _btnVis; }
            set
            {
                if (value != BtnVis)
                {
                    _btnVis = value;
                    OnPropertyChanged(nameof(BtnVis));
                }
            }
        }
        private PersonData _selectedPersonOption;
        public PersonData SelectedPersonOption
        {
            get { return _selectedPersonOption; }
            set
            {
                if (_selectedPersonOption != value)
                {
                    _selectedPersonOption = value;
                    PropertyInfo[] properties = typeof(PersonData).GetProperties();
                    foreach (PropertyInfo property in properties)
                    {
                        OnPropertyChanged(property.Name, property.GetValue(_selectedPersonOption)?.ToString());
                    }
                }
            }
        }
        private ObservableCollection<MainTextInfo> _mainTextOptions; 
        public ObservableCollection<MainTextInfo> MainTextOptions
        {
            get { return _mainTextOptions; }
            set
            {
                if (_mainTextOptions != value)
                {
                    _mainTextOptions = value;
                    OnPropertyChanged(nameof(MainTextOptions));
                }
            }
        } 
        private MainTextInfo _selectedMainTextOption;
        public MainTextInfo SelectedMainTextOption
        {
            get { return _selectedMainTextOption;  } 
            set
            {
                if (value != _selectedMainTextOption)
                {
                    _selectedMainTextOption = value;
                    OnPropertyChanged("MainText", SelectedMainTextOption.MainText);
                }
            }
        }
        private List<KeyValue> _newText;
        public List<KeyValue> NewText
        {
            get { return _newText; }
            set
            {
                if (_newText != value)
                {
                    _newText = value;
                }
            }
        }
        public MainViewModel()
        {
            string startupPath2 = Environment.CurrentDirectory;
            string startupPath = Directory.GetParent(startupPath2).Parent.Parent.FullName + "\\";
            ChangePropertyCommand = new RelayCommand(ChangeCommand);
            AddCommand = new RelayCommand(AddPerCommand);
            OpenNewWindowCommand = new RelayCommand(OpenWindow);
            OpenNewTextWindowCommand = new RelayCommand(OpenTextWindow);
            NewText = new List<KeyValue>();
            NewMainText = string.Empty;
            NewTitle = string.Empty;
            Date = DateTime.Now;
            Options = new List<string> { "Unkown", "Mr.", "Miss"};
            PersonOptions = new ObservableCollection<PersonData>(loadData<PersonData>(startupPath + "SavePersonData.Txt"));
            MainTextOptions = new ObservableCollection<MainTextInfo>(loadData<MainTextInfo>(startupPath + "SaveMaintextData.Txt"));
            AddTextCommand = new RelayCommand(AddText);
            ImgVis = Visibility.Collapsed;
            BtnVis = Visibility.Visible;
        }
        public event PropertyChangedEventHandler PropertyChanged;
        public ICommand ChangePropertyCommand { get; }
        public ICommand AddCommand { get; }
        public ICommand OpenNewWindowCommand { get; }
        public ICommand OpenNewTextWindowCommand { get; }
        public ICommand AddTextCommand { get; }


        protected virtual void OnPropertyChanged(string propertyName, string value = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
            if (value != null)
            {

                if (NewText.Where(x => x.Key == "{" + propertyName + "}").FirstOrDefault() != null) {
                    (NewText.Where(x => x.Key == "{" + propertyName + "}").FirstOrDefault()).Value = value;
                }
                else
                {
                    NewText.Add(new KeyValue { Key = "{" + propertyName + "}", Value = value });
                }
                
            }
            
        }
        public ObservableCollection<T> loadData<T>(string path)
        {
            if (System.IO.File.Exists(path))
            {
                string j = System.IO.File.ReadAllText(path);
                var io = JsonSerializer.Deserialize<ObservableCollection<T>>(j);
                return io;
            }
            return new ObservableCollection<T>();
        }
        private void OpenWindow()
        {
            PopUp addPersonWindow = new PopUp();
            addPersonWindow.DataContext = this;
            addPersonWindow.ShowDialog();
        }
        private void OpenTextWindow()
        {
            PopUpText popUpText = new PopUpText();
            popUpText.DataContext = this;
            popUpText.ShowDialog();
        }
        public void AddPerCommand()
        {

            string startupPath2 = Environment.CurrentDirectory;
            string startupPath = Directory.GetParent(startupPath2).Parent.Parent.FullName + "\\" + "SavePersonData.Txt";
            int id = PersonOptions.Count();
            ObservableCollection<PersonData> lo = PersonOptions;
            PersonData data = new PersonData { CityAndPostcode = CityAndPostcode, Email = Email, FullName = FullName, PhoneNumber = PhoneNumber, Id = id, StreetAddress = StreetAddress };
            lo.Add(data);
            PersonOptions = new ObservableCollection<PersonData> (lo);
            System.IO.File.WriteAllText(startupPath, JsonSerializer.Serialize(PersonOptions));
        }
        public void AddText()
        {
            string startupPath2 = Environment.CurrentDirectory;
            string startupPath = Directory.GetParent(startupPath2).Parent.Parent.FullName + "\\" + "SaveMaintextData.Txt";
            int id = MainTextOptions.Count();
            ObservableCollection<MainTextInfo> lo = MainTextOptions;
            MainTextInfo data = new MainTextInfo { MainText = NewMainText, Id = id, Title = NewTitle };
            lo.Add(data);
            MainTextOptions = new ObservableCollection<MainTextInfo>(lo);
            System.IO.File.WriteAllText(startupPath, JsonSerializer.Serialize(MainTextOptions));
        }
        public async void ChangeCommand()
        {
            BtnVis = Visibility.Collapsed;
            ImgVis = Visibility.Visible;
            PdfGenerator wordEditor = new PdfGenerator();
            NewText.Add(new KeyValue { Key = "{" + nameof(Name) + "}", Value = SelectedOption + " " +Name });
            wordEditor.EditWordDocument("CoverLetter.docx", NewText);
            BtnVis = Visibility.Visible;
            ImgVis = Visibility.Collapsed;
        }
    }
    public class RelayCommand : ICommand
    {
        private readonly Action _execute;

        public RelayCommand(Action execute)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
        }


        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }
        public bool CanExecute(object parameter) => true;

        public void Execute(object parameter) => _execute();
        
    }
}
