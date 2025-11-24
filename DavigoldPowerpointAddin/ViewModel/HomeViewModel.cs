using DavigoldExcel.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Text;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using ModuleModel = DavigoldExcel.Models.Module;

namespace DavigoldPowerpointAddin.ViewModel
{
    public class LabelViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private bool _selection;
        public bool Selection
        {
            get { return _selection; }
            set
            {
                if (_selection != value)
                {
                    _selection = value;
                    OnPropertyChanged(nameof(Selection));
                    //UpdateExcelSheet();
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

        private string _module;
        public string Module
        {
            get { return _module; }
            set
            {
                if (_module != value)
                {
                    _module = value;
                    OnPropertyChanged(nameof(Module));
                }
            }
        }

        private string _subModule;
        public string SubModule
        {
            get { return _subModule; }
            set
            {
                if (_subModule != value)
                {
                    _subModule = value;
                    OnPropertyChanged(nameof(SubModule));
                }
            }
        }

        public string Slug { set; get; }

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class KpiViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private bool _selection;
        public bool Selection
        {
            get { return _selection; }
            set
            {
                if (_selection != value)
                {
                    _selection = value;
                    OnPropertyChanged(nameof(Selection));
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

        private string _module;
        public string Module
        {
            get { return _module; }
            set
            {
                if (_module != value)
                {
                    _module = value;
                    OnPropertyChanged(nameof(Module));
                }
            }
        }

        private string _subModule;
        public string SubModule
        {
            get { return _subModule; }
            set
            {
                if (_subModule != value)
                {
                    _subModule = value;
                    OnPropertyChanged(nameof(SubModule));
                }
            }
        }

        public int Id { set; get; }

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class KeyFigureKpiViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private string _label;
        public string Label
        {
            get { return _label; }
            set
            {
                if (_label != value)
                {
                    _label = value;
                    OnPropertyChanged(nameof(Label));
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

        private string _tableName;
        public string TableName
        {
            get { return _tableName; }
            set
            {
                if (_tableName != value)
                {
                    _tableName = value;
                    OnPropertyChanged(nameof(TableName));
                }
            }
        }

        public int Id { set; get; }

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class ChartViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private int _id;
        public int Id
        {
            get { return _id; }
            set
            {
                if (_id != value)
                {
                    _id = value;
                    OnPropertyChanged(nameof(Id));
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

        private string _module;
        public string Module
        {
            get { return _module; }
            set
            {
                if (_module != value)
                {
                    _module = value;
                    OnPropertyChanged(nameof(Module));
                }
            }
        }

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

    public class ModuleList
    {
        private static string TransformString(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return input;

            // Split the input by spaces
            var words = input.Split(' ');

            // Keep the first word as is, convert remaining words to lowercase
            for (int i = 1; i < words.Length; i++)
            {
                if (words[i].ToLower() == "partners" || words[i].ToLower() == "breakdowns") continue;
                words[i] = words[i].ToLower();
            }

            // Join words back together with a space
            return string.Join(" ", words);
        }

        public static List<ModuleModel> GetModulesList()
        {
            User currentUser = Globals.ThisAddIn.GetUser();

            if (currentUser.Tenant.Modules == null)
            {
                return new List<ModuleModel>();
            }

            if(currentUser.Tenant.ClientType.Contains("Fund of funds"))
            {
                ModuleModel underlying = new ModuleModel();
                underlying.Name = "Underlyings";
                underlying.SubModules = new List<ModuleModel>();

                ModuleModel underlyingSubModule = new ModuleModel();
                underlyingSubModule.Name = "Charts";
                underlyingSubModule.SubModules = new List<ModuleModel>();

                underlying.SubModules.Add(underlyingSubModule);

                currentUser.Tenant.Modules.Add(underlying);
            }
            

            return currentUser.Tenant.Modules.Select(module =>
            {
                module.SubModules = module.SubModules.Select(subModule =>
                {
                    string label = subModule.Name == "NAV Breakdown" ? "NAV Breakdowns" : subModule.Name;
                    subModule.Label = TransformString(label);
                    return subModule;
                }).ToList();

                string currentLabel = module.Name == "Portfolio Fund" ? "Portfolio Funds" : module.Name;
                module.Label = TransformString(currentLabel);
                return module;
            }).ToList();
        }

        public static List<ComboBoxModel> GetKeyFigureTables()
        {
            User currentUser = Globals.ThisAddIn.GetUser();

            List<KeyFigureKpi> keyFigures = currentUser.KeyFigureKpis;

            List<ComboBoxModel> keyFigureTables = keyFigures.Select(f => new { Id = f.TableId, Label = f.TableName }).Distinct().Select(f => new ComboBoxModel() { Label = f.Label, Value = f.Id.ToString()}).ToList<ComboBoxModel>();

            return keyFigureTables;
        }

        public static List<ComboBoxModel> GetKeyFigureTableKpis(string tableId)
        {
            User currentUser = Globals.ThisAddIn.GetUser();

            List<KeyFigureKpi> keyFigures = currentUser.KeyFigureKpis;

            List<ComboBoxModel> keyFigureTableKpis = keyFigures.Where(kpi => kpi.TableId.ToString() == tableId).Select(f => new ComboBoxModel() { Label = f.Name, Value = f.Id.ToString() }).ToList<ComboBoxModel>();

            return keyFigureTableKpis;
        }

        public static List<ComboBoxModel> GetKpiDateValues()
        {
            List<ComboBoxModel> values = new List<ComboBoxModel>();

            values.Add(new ComboBoxModel() { Label = "Closing Year", Value = "closing-year" });
            values.Add(new ComboBoxModel() { Label = "Latest Year", Value = "latest-year" });

            return values;
        }

        public static List<ComboBoxModel> GetKpiTypeValues()
        {
            List<ComboBoxModel> values = new List<ComboBoxModel>();

            values.Add(new ComboBoxModel() { Label = "Actual", Value = "actual" });
            values.Add(new ComboBoxModel() { Label = "Budget", Value = "budget" });
            values.Add(new ComboBoxModel() { Label = "Forecast", Value = "forecast" });

            return values;
        }
        public static List<ComboBoxModel> GetAccountTypesList()
        {
            List<ComboBoxModel> accountTypes = new List<ComboBoxModel>();

            accountTypes.Add(new ComboBoxModel { Label = "Debit", Value = "debit" });
            accountTypes.Add(new ComboBoxModel { Label = "Credit", Value = "credit" });

            return accountTypes;
        }
        public static List<ComboBoxModel> GetPeriodList()
        {
            List<ComboBoxModel> periods = new List<ComboBoxModel>();

            periods.Add(new ComboBoxModel { Label = "N", Value = "n" });
            periods.Add(new ComboBoxModel { Label = "N-1", Value = "n1" });
            periods.Add(new ComboBoxModel { Label = "N-2", Value = "n2" });

            return periods;
        }
        public static List<ComboBoxModel> GetOperatorOptions()
        {
            List<ComboBoxModel> operatorTypes = new List<ComboBoxModel>();

            operatorTypes.Add(new ComboBoxModel { Label = "+", Value = "+" });
            operatorTypes.Add(new ComboBoxModel { Label = "-", Value = "-" });

            return operatorTypes;
        }



        public static List<ComboBoxModel> GetChartDateValues()
        {
            List<ComboBoxModel> values = new List<ComboBoxModel>();

            values.Add(new ComboBoxModel() { Label = "Last 8 quarters", Value = "last-8-quarters" });
            values.Add(new ComboBoxModel() { Label = "Last 4 quarters", Value = "last-4-quarters" });
            values.Add(new ComboBoxModel() { Label = "Since Beginning until last quarter", Value = "beginning-last-quarter" });
            values.Add(new ComboBoxModel() { Label = "Since Beginning till now", Value = "beginning-till-now" });
            values.Add(new ComboBoxModel() { Label = "Last 12M", Value = "last-12-months" });
            values.Add(new ComboBoxModel() { Label = "Year-to-date", Value = "year-to-date" });
            values.Add(new ComboBoxModel() { Label = "Last quarter", Value = "last-quarter" });
            values.Add(new ComboBoxModel() { Label = "Last month", Value = "last-month" });

            return values;
        }

        public static List<ComboBoxModel> GetUnderlyingSectorValues()
        {
            User currentUser = Globals.ThisAddIn.GetUser();

            if (currentUser.Tenant.Dropdowns != null)
            {
                Dropdown underlyingSectorDropdown = currentUser.Tenant.Dropdowns.FirstOrDefault(dropdown => dropdown.MainObject == "Companies" && dropdown.dropdown == "Underlying Sector");
                if (underlyingSectorDropdown != null)
                {
                    List<ComboBoxModel> currentOptions = new List<ComboBoxModel>();
                    currentOptions.Add(new ComboBoxModel() { Label = "None", Value = "" });
                    currentOptions.AddRange(underlyingSectorDropdown.Options.Select(option => new ComboBoxModel() { Label = option.EnLabel, Value = option.Slug }).ToList());
                    return currentOptions;
                }
            }

            return new List<ComboBoxModel>();
        }

        public static List<ComboBoxModel> GetFundQuarterlyValuesHeaderValues()
        {
            User currentUser = Globals.ThisAddIn.GetUser();

            if (currentUser.Tenant.Dropdowns != null)
            {
                Dropdown underlyingSectorDropdown = currentUser.Tenant.Dropdowns.FirstOrDefault(dropdown => dropdown.MainObject == "Funds" && dropdown.dropdown == "Header");
                if (underlyingSectorDropdown != null)
                {
                    List<ComboBoxModel> currentOptions = new List<ComboBoxModel>();
                    currentOptions.Add(new ComboBoxModel() { Label = "None", Value = "" });
                    currentOptions.AddRange(underlyingSectorDropdown.Options.Select(option => new ComboBoxModel() { Label = option.EnLabel, Value = option.Slug }).ToList());
                    return currentOptions;
                }
            }

            return new List<ComboBoxModel>();
        }

        public static List<ComboBoxModel> GetSecurityTypeValues()
        {
            User currentUser = Globals.ThisAddIn.GetUser();

            if (currentUser.Tenant.Dropdowns != null)
            {
                Dropdown underlyingSectorDropdown = currentUser.Tenant.Dropdowns.FirstOrDefault(dropdown => dropdown.MainObject == "Portfolio" && dropdown.dropdown == "Security Type");
                if (underlyingSectorDropdown != null)
                {
                    List<ComboBoxModel> currentOptions = new List<ComboBoxModel>();
                    currentOptions.Add(new ComboBoxModel() { Label = "None", Value = "" });
                    currentOptions.AddRange(underlyingSectorDropdown.Options.Select(option => new ComboBoxModel() { Label = option.EnLabel, Value = option.Slug }).ToList());
                    return currentOptions;
                }
            }

            return new List<ComboBoxModel>();
        }

        public static List<ComboBoxModel> GetUnderlyingStatusValues()
        {
            List<ComboBoxModel> values = new List<ComboBoxModel>();

            values.Add(new ComboBoxModel() { Label = "None", Value = "" });
            values.Add(new ComboBoxModel() { Label = "Active", Value = "active" });
            values.Add(new ComboBoxModel() { Label = "Realized", Value = "realized" });

            return values;
        }

        public static List<ComboBoxModel> GetSecurityStatusValues()
        {
            List<ComboBoxModel> values = new List<ComboBoxModel>();

            values.Add(new ComboBoxModel() { Label = "None", Value = "" });
            values.Add(new ComboBoxModel() { Label = "Active", Value = "active" });
            values.Add(new ComboBoxModel() { Label = "Realized", Value = "realized" });

            return values;
        }

        public static List<ComboBoxModel> GetUnderlyingTableFilterValues()
        {
            List<ComboBoxModel> values = new List<ComboBoxModel>();

            values.Add(new ComboBoxModel() { Label = "All", Value = "all" });
            values.Add(new ComboBoxModel() { Label = "Top 15", Value = "top-15" });
            values.Add(new ComboBoxModel() { Label = "All except top 15", Value = "all-excpet-top-15" });

            return values;
        }

        public static List<ComboBoxModel> GetShareValues()
        {
            User currentUser = Globals.ThisAddIn.GetUser();

            if (currentUser.Shares != null)
            {
                List<ComboBoxModel> shares = new List<ComboBoxModel>();
                shares.Add(new ComboBoxModel() { Label = "None", Value = "" });
                shares.AddRange(currentUser.Shares.Select(share => new ComboBoxModel() { Label = share.Name, Value = share.Name }).ToList());
                return shares;
            }

            return new List<ComboBoxModel>();
        }
    }

    public class HomeViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private ObservableCollection<LabelViewModel> _moduleLabels;
        public ObservableCollection<LabelViewModel> ModuleLabels
        {
            get { return _moduleLabels; }
            set
            {
                _moduleLabels = value;
                OnPropertyChanged(nameof(ModuleLabels));
            }
        }

        private ObservableCollection<KpiViewModel> _moduleKpis;
        public ObservableCollection<KpiViewModel> ModuleKpis
        {
            get { return _moduleKpis; }
            set
            {
                _moduleKpis = value;
                OnPropertyChanged(nameof(ModuleKpis));
            }
        }

        private ObservableCollection<Account> _moduleAccounts;
        public ObservableCollection<Account> ModuleAccounts
        {
            get { return _moduleAccounts; }
            set
            {
                _moduleAccounts = value;
                OnPropertyChanged(nameof(ModuleAccounts));
            }
        }

        private ObservableCollection<AccountGroup> _moduleAccountGroups;
        public ObservableCollection<AccountGroup> ModuleAccountGroups
        {
            get { return _moduleAccountGroups; }
            set
            {
                _moduleAccountGroups = value;
                OnPropertyChanged(nameof(ModuleAccountGroups));
            }
        }

        private ObservableCollection<ChartViewModel> _moduleCharts;
        public ObservableCollection<ChartViewModel> ModuleCharts
        {
            get { return _moduleCharts; }
            set
            {
                _moduleCharts = value;
                OnPropertyChanged(nameof(ModuleCharts));
            }
        }

        private ObservableCollection<KeyFigureKpiViewModel> _keyFigureKpis;
        public ObservableCollection<KeyFigureKpiViewModel> KeyFigureKpis
        {
            get { return _keyFigureKpis; }
            set
            {
                _keyFigureKpis = value;
                OnPropertyChanged(nameof(KeyFigureKpis));
            }
        }

        // Property for the items in the Module ComboBox
        private ObservableCollection<ModuleModel> _modulesComboBox;
        public ObservableCollection<ModuleModel> ModuleComboBoxItems
        {
            get { return _modulesComboBox; }
            set
            {
                _modulesComboBox = value;
                OnPropertyChanged(nameof(ModuleComboBoxItems));
            }
        }

        // Property for the selected item in the Module ComboBox
        private ModuleModel _selectedModuleComboxBoxItem;
        public ModuleModel SelectedModuleComboBoxItem
        {
            get { return _selectedModuleComboxBoxItem; }
            set
            {
                _selectedModuleComboxBoxItem = value;
                OnPropertyChanged(nameof(SelectedModuleComboBoxItem));
                UpdateSecondComboBoxItems();
                UpdateModuleKpi();
            }
        }

        // Property for the items in the Sub Module ComboBox
        private ObservableCollection<ModuleModel> _subModuleComboBoxItems;
        public ObservableCollection<ModuleModel> SubModuleComboBoxItems
        {
            get { return _subModuleComboBoxItems; }
            set
            {
                _subModuleComboBoxItems = value;
                OnPropertyChanged(nameof(SubModuleComboBoxItems));
            }
        }

        // Property for the selected item in the Sub Module ComboBox
        private ModuleModel _selectedSubModuleComboxBoxItem;
        public ModuleModel SelectedSubModuleComboBoxItem
        {
            get { return _selectedSubModuleComboxBoxItem; }
            set
            {
                _selectedSubModuleComboxBoxItem = value;
                OnPropertyChanged(nameof(SelectedSubModuleComboBoxItem));
                UpdateModuleLabels();
                UpdateModuleKpi();
                UpdateModuleKeyFigure();
                UpdateModuleKeyFigureKpi();
                UpdateModuleChart();
                UpdateExtraFields();
                UpdateAccountKpi();
                UpdateAccountGroupKpi();
                UpdateExtraAccountFields();
                UpdateFundHeader();
            }
        }

        private ObservableCollection<ComboBoxModel> _keyFigureTableComboBoxItems;
        public ObservableCollection<ComboBoxModel> KeyFigureTableComboBoxItems
        {
            get { return _keyFigureTableComboBoxItems; }
            set
            {
                _keyFigureTableComboBoxItems = value;
                OnPropertyChanged(nameof(KeyFigureTableComboBoxItems));
            }
        }

        private ComboBoxModel _selectedKeyFigureTableComboBoxItem;
        public ComboBoxModel SelectedKeyFigureTableComboBoxItem
        {
            get { return _selectedKeyFigureTableComboBoxItem; }
            set
            {
                _selectedKeyFigureTableComboBoxItem = value;
                OnPropertyChanged(nameof(SelectedKeyFigureTableComboBoxItem));
                UpdateModuleKeyFigureKpis();
            }
        }

        private ObservableCollection<ComboBoxModel> _keyFigureKpiComboBoxItems;
        public ObservableCollection<ComboBoxModel> KeyFigureKpiComboBoxItems
        {
            get { return _keyFigureKpiComboBoxItems; }
            set
            {
                _keyFigureKpiComboBoxItems = value;
                OnPropertyChanged(nameof(KeyFigureKpiComboBoxItems));
            }
        }

        private ComboBoxModel _selectedKeyFigureKpiComboBoxItem;
        public ComboBoxModel SelectedKeyFigureKpiComboBoxItem
        {
            get { return _selectedKeyFigureKpiComboBoxItem; }
            set
            {
                _selectedKeyFigureKpiComboBoxItem = value;
                OnPropertyChanged(nameof(SelectedKeyFigureKpiComboBoxItem));
            }
        }

        private ObservableCollection<ComboBoxModel> _keyFigureDateComboBoxItems;
        public ObservableCollection<ComboBoxModel> KeyFigureDateComboBoxItems
        {
            get { return _keyFigureDateComboBoxItems; }
            set
            {
                _keyFigureDateComboBoxItems = value;
                OnPropertyChanged(nameof(KeyFigureDateComboBoxItems));
            }
        }

        private ObservableCollection<ComboBoxModel> _underlyingTableFilterComboBoxItems;
        public ObservableCollection<ComboBoxModel> UnderlyingTableFilterComboBoxItems
        {
            get { return _underlyingTableFilterComboBoxItems; }
            set
            {
                _underlyingTableFilterComboBoxItems = value;
                OnPropertyChanged(nameof(UnderlyingTableFilterComboBoxItems));
            }
        }

        // Property for the selected item in the Key Figure ComboBox
        private ComboBoxModel _selectedKeyFigureDateComboBoxItem;
        public ComboBoxModel SelectedKeyFigureDateComboBoxItem
        {
            get { return _selectedKeyFigureDateComboBoxItem; }
            set
            {
                _selectedKeyFigureDateComboBoxItem = value;
                OnPropertyChanged(nameof(SelectedKeyFigureDateComboBoxItem));
            }
        }

        private ObservableCollection<ComboBoxModel> _keyFigureTypeComboBoxItems;
        public ObservableCollection<ComboBoxModel> KeyFigureTypeComboBoxItems
        {
            get { return _keyFigureTypeComboBoxItems; }
            set
            {
                _keyFigureTypeComboBoxItems = value;
                OnPropertyChanged(nameof(KeyFigureTypeComboBoxItems));
            }
        }

        // Property for the selected item in the Key Figure Type ComboBox
        private ComboBoxModel _selectedKeyFigureTypeComboBoxItem;
        public ComboBoxModel SelectedKeyFigureTypeComboBoxItem
        {
            get { return _selectedKeyFigureTypeComboBoxItem; }
            set
            {
                _selectedKeyFigureTypeComboBoxItem = value;
                OnPropertyChanged(nameof(SelectedKeyFigureTypeComboBoxItem));
            }
        }

        private ObservableCollection<ComboBoxModel> _chartDateComboBoxItems;
        public ObservableCollection<ComboBoxModel> ChartDateComboBoxItems
        {
            get { return _chartDateComboBoxItems; }
            set
            {
                _chartDateComboBoxItems = value;
                OnPropertyChanged(nameof(ChartDateComboBoxItems));
            }
        }

        // Property for the selected item in the Chart Date ComboBox
        private ComboBoxModel _selectedChartDateComboBoxItem;
        public ComboBoxModel SelectedChartDateComboBoxItem
        {
            get { return _selectedChartDateComboBoxItem; }
            set
            {
                _selectedChartDateComboBoxItem = value;
                OnPropertyChanged(nameof(SelectedChartDateComboBoxItem));
            }
        }

        private ObservableCollection<ComboBoxModel> _underlyingSectorComboBoxItems;
        public ObservableCollection<ComboBoxModel> UnderlyingSectorComboBoxItems
        {
            get { return _underlyingSectorComboBoxItems; }
            set
            {
                _underlyingSectorComboBoxItems = value;
                OnPropertyChanged(nameof(UnderlyingSectorComboBoxItems));
            }
        }

        private ComboBoxModel _selectedUnderlyingSectorComboBoxItem;
        public ComboBoxModel SelectedUnderlyingSectorComboBoxItem
        {
            get { return _selectedUnderlyingSectorComboBoxItem; }
            set
            {
                _selectedUnderlyingSectorComboBoxItem = value;
                OnPropertyChanged(nameof(SelectedUnderlyingSectorComboBoxItem));
            }
        }

        private ObservableCollection<ComboBoxModel> _underlyingStatusComboBoxItems;
        public ObservableCollection<ComboBoxModel> UnderlyingStatusComboBoxItems
        {
            get { return _underlyingStatusComboBoxItems; }
            set
            {
                _underlyingStatusComboBoxItems = value;
                OnPropertyChanged(nameof(UnderlyingStatusComboBoxItems));
            }
        }

        private ComboBoxModel _selectedUnderlyingStatusComboBoxItem;
        public ComboBoxModel SelectedUnderlyingStatusComboBoxItem
        {
            get { return _selectedUnderlyingStatusComboBoxItem; }
            set
            {
                _selectedUnderlyingStatusComboBoxItem = value;
                OnPropertyChanged(nameof(SelectedUnderlyingStatusComboBoxItem));
            }
        }

        private ComboBoxModel _selectedUnderlyingTableFilterComboBoxItem;
        public ComboBoxModel SelectedUnderlyingTableFilterComboBoxItem
        {
            get { return _selectedUnderlyingTableFilterComboBoxItem; }
            set
            {
                _selectedUnderlyingTableFilterComboBoxItem = value;
                OnPropertyChanged(nameof(SelectedUnderlyingTableFilterComboBoxItem));
            }
        }

        private ObservableCollection<ComboBoxModel> _sharesComboBoxItems;
        public ObservableCollection<ComboBoxModel> SharesComboBoxItems
        {
            get { return _sharesComboBoxItems; }
            set
            {
                _sharesComboBoxItems = value;
                OnPropertyChanged(nameof(SharesComboBoxItems));
            }
        }

        private ComboBoxModel _selectedShareComboBoxItem;
        public ComboBoxModel SelectedShareComboBoxItem
        {
            get { return _selectedShareComboBoxItem; }
            set
            {
                _selectedShareComboBoxItem = value;
                OnPropertyChanged(nameof(SelectedShareComboBoxItem));
            }
        }

        private ObservableCollection<ComboBoxModel> _securityTypeComboBoxItems;
        public ObservableCollection<ComboBoxModel> SecurityTypeComboBoxItems
        {
            get { return _securityTypeComboBoxItems; }
            set
            {
                _securityTypeComboBoxItems = value;
                OnPropertyChanged(nameof(SecurityTypeComboBoxItems));
            }
        }

        private ComboBoxModel _selectedSecurityTypeComboBoxItem;
        public ComboBoxModel SelectedSecurityTypeComboBoxItem
        {
            get { return _selectedSecurityTypeComboBoxItem; }
            set
            {
                _selectedSecurityTypeComboBoxItem = value;
                OnPropertyChanged(nameof(SelectedSecurityTypeComboBoxItem));
            }
        }

        private ObservableCollection<ComboBoxModel> _fundHeaderComboBoxItems;
        public ObservableCollection<ComboBoxModel> FundHeaderComboBoxItems
        {
            get { return _fundHeaderComboBoxItems; }
            set
            {
                _fundHeaderComboBoxItems = value;
                OnPropertyChanged(nameof(FundHeaderComboBoxItems));
            }
        }

        private ComboBoxModel _selectedFundHeaderComboBoxItem;
        public ComboBoxModel SelectedFundHeaderComboBoxItem
        {
            get { return _selectedFundHeaderComboBoxItem; }
            set
            {
                _selectedFundHeaderComboBoxItem = value;
                OnPropertyChanged(nameof(SelectedFundHeaderComboBoxItem));
            }
        }

        private ObservableCollection<ComboBoxModel> _securityStatusComboBoxItems;
        public ObservableCollection<ComboBoxModel> SecurityStatusComboBoxItems
        {
            get { return _securityStatusComboBoxItems; }
            set
            {
                _securityStatusComboBoxItems = value;
                OnPropertyChanged(nameof(SecurityStatusComboBoxItems));
            }
        }


        private ComboBoxModel _selectedSecurityStatusComboBoxItem;
        public ComboBoxModel SelectedSecurityStatusComboBoxItem
        {
            get { return _selectedSecurityStatusComboBoxItem; }
            set
            {
                _selectedSecurityStatusComboBoxItem = value;
                OnPropertyChanged(nameof(SelectedSecurityStatusComboBoxItem));
            }
        }



        // Property for the items in the Account Type ComboBox
        private ObservableCollection<ComboBoxModel> _accountTypeComboBoxItems;
        public ObservableCollection<ComboBoxModel> AccountTypeComboBoxItems
        {
            get { return _accountTypeComboBoxItems; }
            set
            {
                _accountTypeComboBoxItems = value;
                OnPropertyChanged(nameof(AccountTypeComboBoxItems));
            }
        }

        private ComboBoxModel _selectedAccountTypeComboBoxItem;
        public ComboBoxModel SelectedAccountTypeComboBoxItem
        {
            get { return _selectedAccountTypeComboBoxItem; }
            set
            {
                _selectedAccountTypeComboBoxItem = value;
                OnPropertyChanged(nameof(SelectedAccountTypeComboBoxItem));
            }
        }

        // Property for the items in the Period ComboBox
        private ObservableCollection<ComboBoxModel> _periodComboBoxItems;
        public ObservableCollection<ComboBoxModel> PeriodComboBoxItems
        {
            get { return _periodComboBoxItems; }
            set
            {
                _periodComboBoxItems = value;
                OnPropertyChanged(nameof(PeriodComboBoxItems));
            }
        }

        private ComboBoxModel _selectedPeriodComboBoxItem;
        public ComboBoxModel SelectedPeriodComboBoxItem
        {
            get { return _selectedPeriodComboBoxItem; }
            set
            {
                _selectedPeriodComboBoxItem = value;
                OnPropertyChanged(nameof(SelectedPeriodComboBoxItem));
            }
        }

        private ObservableCollection<ComboBoxModel> _operatorTypeComboBoxItems;
        public ObservableCollection<ComboBoxModel> OperatorTypeComboBoxItems
        {
            get { return _operatorTypeComboBoxItems; }
            set
            {
                _operatorTypeComboBoxItems = value;
                OnPropertyChanged(nameof(OperatorTypeComboBoxItems));
            }
        }


        public HomeViewModel()
        {
            List<ModuleModel> currentModules = ModuleList.GetModulesList();
            ModuleComboBoxItems = new ObservableCollection<ModuleModel>(currentModules);
            SubModuleComboBoxItems = new ObservableCollection<ModuleModel>();
            KeyFigureDateComboBoxItems = new ObservableCollection<ComboBoxModel>();
            KeyFigureTypeComboBoxItems = new ObservableCollection<ComboBoxModel>();
            ChartDateComboBoxItems = new ObservableCollection<ComboBoxModel>();
            ModuleKpis = new ObservableCollection<KpiViewModel>();
            KeyFigureKpis = new ObservableCollection<KeyFigureKpiViewModel>();
            ModuleCharts = new ObservableCollection<ChartViewModel>();
            UnderlyingSectorComboBoxItems = new ObservableCollection<ComboBoxModel>();
            UnderlyingStatusComboBoxItems = new ObservableCollection<ComboBoxModel>();
            UnderlyingTableFilterComboBoxItems = new ObservableCollection<ComboBoxModel>();
            SharesComboBoxItems = new ObservableCollection<ComboBoxModel>();
            SecurityTypeComboBoxItems = new ObservableCollection<ComboBoxModel>();
            SecurityStatusComboBoxItems = new ObservableCollection<ComboBoxModel>();
            KeyFigureTableComboBoxItems = new ObservableCollection<ComboBoxModel>();
            KeyFigureKpiComboBoxItems = new ObservableCollection<ComboBoxModel>();
            FundHeaderComboBoxItems = new ObservableCollection<ComboBoxModel>();
        }

        private void UpdateSecondComboBoxItems()
        {
            if (SelectedModuleComboBoxItem != null)
            {
                var currentSubModules = SelectedModuleComboBoxItem.SubModules;

                List<Label> TenantLabels = Globals.ThisAddIn.GetUser().Tenant.Labels;

                currentSubModules = currentSubModules.Where(mod =>
                {
                    if (SelectedModuleComboBoxItem.Name == "Funds" && (mod.Name == "Accounts" || mod.Name == "Account Groups"))
                    {
                        return true;
                    }

                    List<Label> ModuleLabelsList = TenantLabels.Where(label => label.Module == SelectedModuleComboBoxItem.Name && label.Form == mod.Name && label.ChangeIn == "Forms").ToList();

                    Label CurrentModuleLabel = ModuleLabelsList.FirstOrDefault();

                    return CurrentModuleLabel != null && CurrentModuleLabel?.IsAddin == true;
                }).ToList();

                ModuleModel existingKpiModule = currentSubModules.Where(a => a.Label == "KPIs").FirstOrDefault();
                if (existingKpiModule == null)
                {
                    currentSubModules.Add(new ModuleModel() { Label = "KPIs", Name = "KPIs", SubModules = new List<ModuleModel>() });
                }
                SubModuleComboBoxItems = new ObservableCollection<ModuleModel>(currentSubModules);
            }
        }

        public void UpdateModuleLabels()
        {
            if (SelectedSubModuleComboBoxItem != null)
            {
                List<Label> TenantLabels = Globals.ThisAddIn.GetUser().Tenant.Labels;
                List<Label> ModuleLabelsList = TenantLabels.Where(label => label.Module == SelectedModuleComboBoxItem.Name && label.Form == SelectedSubModuleComboBoxItem.Name && label.ChangeIn == "Forms").ToList();

                Label CurrentModuleLabel = ModuleLabelsList.FirstOrDefault();

                if (CurrentModuleLabel != null)
                {
                    List<LabelViewModel> viewModelLabels = CurrentModuleLabel.Labels.Select(currentLabel => new LabelViewModel() { Name = currentLabel.LabelEn, Module = SelectedModuleComboBoxItem.Name, SubModule = SelectedSubModuleComboBoxItem.Name, Slug = currentLabel.Slug, Selection = false }).ToList();
                    ModuleLabels = new ObservableCollection<LabelViewModel>(viewModelLabels);
                    return;
                }
            }

            ModuleLabels = new ObservableCollection<LabelViewModel>();
        }

        public void UpdateModuleKpi()
        {
            if (SelectedModuleComboBoxItem != null)
            {
                List<Kpi> TenantKpis = Globals.ThisAddIn.GetUser().Kpis;
                List<string> CurrentModuleNameList = SelectedModuleComboBoxItem.Name.ToLower().Split(' ').ToList();
                string CurrentModuleName = String.Join("-", CurrentModuleNameList);
                string CurrentSubModule = SelectedSubModuleComboBoxItem != null ? SelectedSubModuleComboBoxItem.Name.ToLower() : null;

                if (CurrentSubModule != null && CurrentSubModule == "kpis")
                {
                    List<Kpi> ModuleKpiList = TenantKpis.Where(label => label.Module == CurrentModuleName).ToList();

                    if (ModuleKpiList != null && ModuleKpiList.Count > 0)
                    {
                        List<KpiViewModel> viewModelKpi = ModuleKpiList.Select(currentKpi => new KpiViewModel() { Id = currentKpi.Id, Name = currentKpi.Name, Module = currentKpi.Module, SubModule = currentKpi.SubModule, Selection = false }).ToList();
                        ModuleKpis = new ObservableCollection<KpiViewModel>(viewModelKpi);
                        return;
                    }
                }
            }

            ModuleKpis = new ObservableCollection<KpiViewModel>();
        }

        public void UpdateModuleKeyFigureKpi()
        {
            if (SelectedModuleComboBoxItem != null)
            {
                List<KeyFigureKpi> TenantKeyFigureKpis = Globals.ThisAddIn.GetUser().KeyFigureKpis;
                string CurrentModuleName = SelectedModuleComboBoxItem != null ? SelectedModuleComboBoxItem.Name.ToLower() : null;
                string CurrentSubModule = SelectedSubModuleComboBoxItem != null ? SelectedSubModuleComboBoxItem.Name.ToLower() : null;

                if (CurrentModuleName != null && CurrentModuleName == "portfolio" && CurrentSubModule != null && CurrentSubModule == "key figures specific")
                {

                    if (TenantKeyFigureKpis != null && TenantKeyFigureKpis.Count > 0)
                    {
                        List<KeyFigureKpiViewModel> viewModelKpi = TenantKeyFigureKpis.Select(currentKpi => new KeyFigureKpiViewModel() { Id = currentKpi.Id, Name = currentKpi.Name, TableName = currentKpi.TableName, Label = $"{currentKpi.Name} ({currentKpi.TableName})" }).ToList();
                        KeyFigureKpis = new ObservableCollection<KeyFigureKpiViewModel>(viewModelKpi);

                        List<ComboBoxModel> dateValues = ModuleList.GetKpiDateValues();
                        KeyFigureDateComboBoxItems = new ObservableCollection<ComboBoxModel>(dateValues);

                        List<ComboBoxModel> typeValues = ModuleList.GetKpiTypeValues();
                        KeyFigureTypeComboBoxItems = new ObservableCollection<ComboBoxModel>(typeValues);
                        return;
                    }
                }
            }

            KeyFigureKpis = new ObservableCollection<KeyFigureKpiViewModel>();
        }

        public void UpdateModuleKeyFigure()
        {
            if (SelectedModuleComboBoxItem != null)
            {
                string CurrentModuleName = SelectedModuleComboBoxItem != null ? SelectedModuleComboBoxItem.Name.ToLower() : null;
                string CurrentSubModule = SelectedSubModuleComboBoxItem != null ? SelectedSubModuleComboBoxItem.Name.ToLower() : null;

                if (CurrentModuleName != null && CurrentModuleName == "portfolio" && CurrentSubModule != null && CurrentSubModule == "key figures")
                {
                    List<ComboBoxModel> tableValues = ModuleList.GetKeyFigureTables();
                    KeyFigureTableComboBoxItems = new ObservableCollection<ComboBoxModel>(tableValues);
                    return;
                }
            }

            KeyFigureTableComboBoxItems = new ObservableCollection<ComboBoxModel>();
            SelectedKeyFigureTableComboBoxItem = null;
        }

        public void UpdateModuleKeyFigureKpis()
        {
            if (SelectedModuleComboBoxItem != null && SelectedKeyFigureTableComboBoxItem != null)
            {
                string CurrentModuleName = SelectedModuleComboBoxItem != null ? SelectedModuleComboBoxItem.Name.ToLower() : null;
                string CurrentSubModule = SelectedSubModuleComboBoxItem != null ? SelectedSubModuleComboBoxItem.Name.ToLower() : null;

                if (CurrentModuleName != null && CurrentModuleName == "portfolio" && CurrentSubModule != null && CurrentSubModule == "key figures")
                {
                    List<ComboBoxModel> kpiValues = ModuleList.GetKeyFigureTableKpis(SelectedKeyFigureTableComboBoxItem.Value);
                    KeyFigureKpiComboBoxItems = new ObservableCollection<ComboBoxModel>(kpiValues);
                    return;
                }
            }

            KeyFigureKpiComboBoxItems = new ObservableCollection<ComboBoxModel>();
            SelectedKeyFigureKpiComboBoxItem = null;
        }

        public void UpdateFundHeader()
        {
            if (SelectedModuleComboBoxItem != null && SelectedSubModuleComboBoxItem != null)
            {
                string CurrentModuleName = SelectedModuleComboBoxItem != null ? SelectedModuleComboBoxItem.Name.ToLower() : null;
                string CurrentSubModule = SelectedSubModuleComboBoxItem != null ? SelectedSubModuleComboBoxItem.Name.ToLower() : null;

                if (CurrentModuleName != null && CurrentModuleName == "funds" && CurrentSubModule != null && CurrentSubModule == "quarterly updates")
                {
                    List<ComboBoxModel> headerValues = ModuleList.GetFundQuarterlyValuesHeaderValues();
                    FundHeaderComboBoxItems = new ObservableCollection<ComboBoxModel>(headerValues);
                    return;
                }
            }

            FundHeaderComboBoxItems = new ObservableCollection<ComboBoxModel>();
            SelectedFundHeaderComboBoxItem = null;
        }

        public void UpdateModuleChart()
        {
            if (SelectedModuleComboBoxItem != null)
            {
                List<Chart> TenantCharts = Globals.ThisAddIn.GetUser().Charts;
                List<string> CurrentModuleNameList = SelectedModuleComboBoxItem.Name.ToLower().Split(' ').ToList();
                string CurrentModuleName = String.Join("-", CurrentModuleNameList);
                string CurrentSubModule = SelectedSubModuleComboBoxItem != null ? SelectedSubModuleComboBoxItem.Name.ToLower() : null;

                if (CurrentSubModule != null && CurrentSubModule == "charts")
                {
                    List<Chart> ModuleChartList = TenantCharts.Where(label => label.Module == CurrentModuleName).ToList();

                    if (ModuleChartList != null && ModuleChartList.Count > 0)
                    {
                        List<ChartViewModel> viewModelKpi = ModuleChartList.Select(currentKpi => new ChartViewModel() { Id = currentKpi.Id, Name = currentKpi.Name, Module = currentKpi.Module }).ToList();
                        ModuleCharts = new ObservableCollection<ChartViewModel>(viewModelKpi);

                        List<ComboBoxModel> dateValues = ModuleList.GetChartDateValues();
                        ChartDateComboBoxItems = new ObservableCollection<ComboBoxModel>(dateValues);

                        ComboBoxModel defaultSelectedDate = dateValues.FirstOrDefault(date => date.Value == "beginning-till-now");
                        if (defaultSelectedDate != null)
                        {
                            SelectedChartDateComboBoxItem = defaultSelectedDate;
                        }
                        return;
                    }
                }
            }

            ModuleCharts = new ObservableCollection<ChartViewModel>();
        }

        public void UpdateAccountKpi()
        {
            if (SelectedModuleComboBoxItem != null)
            {
                List<Account> TenantAccounts = Globals.ThisAddIn.GetUser().Accounts;
                List<string> CurrentModuleNameList = SelectedModuleComboBoxItem.Name.ToLower().Split(' ').ToList();
                string CurrentModuleName = String.Join("-", CurrentModuleNameList);
                string CurrentSubModule = SelectedSubModuleComboBoxItem != null ? SelectedSubModuleComboBoxItem.Name.ToLower() : null;

                if (CurrentSubModule != null && CurrentSubModule == "accounts")
                {

                    if (TenantAccounts != null && TenantAccounts.Count > 0)
                    {
                        ModuleAccounts = new ObservableCollection<Account>(TenantAccounts);
                        return;
                    }
                }
            }

            ModuleAccounts = new ObservableCollection<Account>();
        }

        public void UpdateAccountGroupKpi()
        {
            if (SelectedModuleComboBoxItem != null)
            {
                List<string> TenantAccountGroups = Globals.ThisAddIn.GetUser().AccountGroups;
                List<string> CurrentModuleNameList = SelectedModuleComboBoxItem.Name.ToLower().Split(' ').ToList();
                string CurrentModuleName = String.Join("-", CurrentModuleNameList);
                string CurrentSubModule = SelectedSubModuleComboBoxItem != null ? SelectedSubModuleComboBoxItem.Name.ToLower() : null;

                if (CurrentSubModule != null && CurrentSubModule == "account groups")
                {

                    if (TenantAccountGroups != null && TenantAccountGroups.Count > 0)
                    {
                        ModuleAccountGroups = new ObservableCollection<AccountGroup>(TenantAccountGroups.OrderBy(accountNumber => accountNumber).Select(accountNumber => new AccountGroup() { AccountType = null, Period = null, Selection = false, Name = accountNumber}).ToList());
                        return;
                    }
                }
            }

            ModuleAccountGroups = new ObservableCollection<AccountGroup>();
        }

        public void UpdateExtraFields()
        {
            UnderlyingTableFilterComboBoxItems = new ObservableCollection<ComboBoxModel>();
            UnderlyingSectorComboBoxItems = new ObservableCollection<ComboBoxModel>();
            UnderlyingStatusComboBoxItems = new ObservableCollection<ComboBoxModel>();
            SharesComboBoxItems = new ObservableCollection<ComboBoxModel>();

            if (SelectedModuleComboBoxItem != null && SelectedSubModuleComboBoxItem != null)
            {
                if (SelectedModuleComboBoxItem.Name.ToLower() == "portfolio fund")
                {
                    if(SelectedSubModuleComboBoxItem.Name.ToLower() == "kpis")
                    {
                        UnderlyingSectorComboBoxItems = new ObservableCollection<ComboBoxModel>(ModuleList.GetUnderlyingSectorValues());
                        return;
                    }

                    if (SelectedModuleComboBoxItem.Name.ToLower() == "portfolio fund" && SelectedSubModuleComboBoxItem.Name.ToLower() == "underlying investments")
                    {
                        UnderlyingTableFilterComboBoxItems = new ObservableCollection<ComboBoxModel>(ModuleList.GetUnderlyingTableFilterValues());
                        return;
                    }

                }

                if (SelectedModuleComboBoxItem.Name.ToLower() == "funds")
                {
                    if (SelectedSubModuleComboBoxItem.Name.ToLower() == "kpis" || SelectedSubModuleComboBoxItem.Name.ToLower() == "charts")
                    {
                        SharesComboBoxItems = new ObservableCollection<ComboBoxModel>(ModuleList.GetShareValues());
                    }
                    return;
                }

                if (SelectedModuleComboBoxItem.Name.ToLower() == "portfolio")
                {
                    if (SelectedSubModuleComboBoxItem.Name.ToLower() == "kpis")
                    {
                        SecurityTypeComboBoxItems = new ObservableCollection<ComboBoxModel>(ModuleList.GetSecurityTypeValues());
                        SecurityStatusComboBoxItems = new ObservableCollection<ComboBoxModel>(ModuleList.GetSecurityStatusValues());
                    }

                    if (SelectedSubModuleComboBoxItem.Name.ToLower() == "target general")
                    {
                        SecurityStatusComboBoxItems = new ObservableCollection<ComboBoxModel>(ModuleList.GetSecurityStatusValues());
                    }
                    return;
                }

            }
        }

        public void UpdateExtraAccountFields()
        {
            string CurrentSubModule = SelectedSubModuleComboBoxItem != null ? SelectedSubModuleComboBoxItem.Name.ToLower() : null;
            if (CurrentSubModule != null && (CurrentSubModule == "accounts" || CurrentSubModule == "account groups"))
            {
                AccountTypeComboBoxItems = new ObservableCollection<ComboBoxModel>(ModuleList.GetAccountTypesList());
                PeriodComboBoxItems = new ObservableCollection<ComboBoxModel>(ModuleList.GetPeriodList());
                OperatorTypeComboBoxItems = new ObservableCollection<ComboBoxModel>(ModuleList.GetOperatorOptions());
                return;
            }

            AccountTypeComboBoxItems = new ObservableCollection<ComboBoxModel>();
            PeriodComboBoxItems = new ObservableCollection<ComboBoxModel>();
            OperatorTypeComboBoxItems = new ObservableCollection<ComboBoxModel>();
        }

        public void ResetAccounts()
        {
            var user = Globals.ThisAddIn.GetUser();
            ModuleAccounts = new ObservableCollection<Account>(user.Accounts);
            ModuleAccountGroups = new ObservableCollection<AccountGroup>(user.AccountGroups.OrderBy(accountNumber => accountNumber).Select(accountNumber => new AccountGroup() { AccountType = null, Period = null, Selection = false, Name = accountNumber }).ToList());
        }


        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

}
