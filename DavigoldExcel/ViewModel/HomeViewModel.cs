using DavigoldExcel.Models;
using DavigoldExcel.Service;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Label = DavigoldExcel.Models.Label;
using ModuleModel = DavigoldExcel.Models.Module;

namespace DavigoldExcel.ViewModel
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
                    UpdateExcelSheet();
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

        private void UpdateExcelSheet()
        {

            if (Selection)
            {
                ExcelService.AddColumnAtLastPosition(this);
            }
            else
            {
                ExcelService.RemoveColumn(this);
            }
        }

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
                    UpdateExcelSheet();
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

        private void UpdateExcelSheet()
        {

            if (Selection)
            {
                ExcelService.AddKpiAtLastPosition(this);
            }
            else
            {
                ExcelService.RemoveKpi(this);
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

            return currentUser.Tenant.Modules.Select(module =>
            {
                module.SubModules = module.SubModules.Where(subModule => subModule.Name != "Charts").Where(subModule => subModule.Name != "Key Figures Specific").Select(subModule =>
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

        public static List<ComboBoxModel> GetValueTypeList()
        {
            List<ComboBoxModel> valueTypes = new List<ComboBoxModel>();

            valueTypes.Add(new ComboBoxModel { Label = "Table", Value = "list" });
            valueTypes.Add(new ComboBoxModel { Label = "Value", Value = "value" });

            return valueTypes;
        }

        public static List<ComboBoxModel> GetDataOperationTypesList()
        {
            List<ComboBoxModel> dataOperationTypes = new List<ComboBoxModel>();

            dataOperationTypes.Add(new ComboBoxModel { Label = "Download", Value = "D" });
            dataOperationTypes.Add(new ComboBoxModel { Label = "Upload", Value = "U" });

            return dataOperationTypes;
        }

        public static List<ComboBoxModel> GetDataTypesList()
        {
            List<ComboBoxModel> dataTypes = new List<ComboBoxModel>();

            dataTypes.Add(new ComboBoxModel { Value = "fields", Label = "Fields" });
            dataTypes.Add(new ComboBoxModel { Value = "kpis", Label = "KPIs" });

            return dataTypes;
        }

        public static List<ComboBoxModel> GetFundValues()
        {
            User currentUser = Globals.ThisAddIn.GetUser();

            if (currentUser.Funds != null)
            {
                List<ComboBoxModel> funds = new List<ComboBoxModel>();
                funds.Add(new ComboBoxModel() { Label = "All Funds", Value = "" });
                funds.AddRange(currentUser.Funds.Select(fund => new ComboBoxModel() { Label = fund.Name, Value = fund.Id.ToString() }).ToList());
                return funds;
            }

            return new List<ComboBoxModel>();
        }
    }

    public class HomeViewModel : INotifyPropertyChanged
    {
        bool onlyKeyFiguresUpload = false;
        public event PropertyChangedEventHandler PropertyChanged;

        private bool _isEnglish = true;
        public bool IsEnglish
        {
            get { return _isEnglish; }
            set
            {
                if (_isEnglish != value)
                {
                    _isEnglish = value;
                    OnPropertyChanged(nameof(IsEnglish));
                }
            }
        }

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

        private ObservableCollection<string> _moduleAccountGroups;
        public ObservableCollection<string> ModuleAccountGroups
        {
            get { return _moduleAccountGroups; }
            set
            {
                _moduleAccountGroups = value;
                OnPropertyChanged(nameof(ModuleAccountGroups));
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
                //UpdateKpiItems();

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
                UpdateAccountKpi();
                UpdateAccountGroupKpi();
                UpdateExtraFields();
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

        // Property for the items in the Period ComboBox
        private ObservableCollection<ComboBoxModel> _valueTypeComboBoxItems;
        public ObservableCollection<ComboBoxModel> ValueTypeComboxBoxItems
        {
            get { return _valueTypeComboBoxItems; }
            set
            {
                _valueTypeComboBoxItems = value;
                OnPropertyChanged(nameof(ValueTypeComboxBoxItems));
            }
        }

        private ComboBoxModel _selectedValueTypeComboBoxItem;
        public ComboBoxModel SelectedValueTypeComboBoxItem
        {
            get { return _selectedValueTypeComboBoxItem; }
            set
            {
                _selectedValueTypeComboBoxItem = value;
                OnPropertyChanged(nameof(SelectedValueTypeComboBoxItem));
            }
        }

        private ObservableCollection<ComboBoxModel> _dataOperationTypeComboBoxItems;
        public ObservableCollection<ComboBoxModel> DataOperationTypeComboBoxItems
        {
            get { return _dataOperationTypeComboBoxItems; }
            set
            {
                _dataOperationTypeComboBoxItems = value;
                OnPropertyChanged(nameof(DataOperationTypeComboBoxItems));
            }
        }

        private ComboBoxModel _selectedDataOperationTypeComboBoxItem;
        public ComboBoxModel SelectedDataOperationTypeComboBoxItem
        {
            get { return _selectedDataOperationTypeComboBoxItem; }
            set
            {
                _selectedDataOperationTypeComboBoxItem = value;
                OnPropertyChanged(nameof(SelectedDataOperationTypeComboBoxItem));

                UpdateComboBoxItems();
                UpdateModuleLabels();
                UpdateModuleKpi();
                UpdateAccountKpi();
                UpdateAccountGroupKpi();
                UpdateExtraFields();
                UpdateSecondComboBoxItems();
                UpdateDataTypeComboBox();
            }
        }

        private ObservableCollection<ComboBoxModel> _fundsComboBoxItems;
        public ObservableCollection<ComboBoxModel> FundsComboBoxItems
        {
            get { return _fundsComboBoxItems; }
            set
            {
                _fundsComboBoxItems = value;
                OnPropertyChanged(nameof(FundsComboBoxItems));
            }
        }

        private ComboBoxModel _selectedFundComboBoxItem;
        public ComboBoxModel SelectedFundComboBoxItem
        {
            get { return _selectedFundComboBoxItem; }
            set
            {
                _selectedFundComboBoxItem = value;
                OnPropertyChanged(nameof(SelectedFundComboBoxItem));
            }
        }

        private ObservableCollection<ComboBoxModel> _dataTypesComboBoxItems;
        public ObservableCollection<ComboBoxModel> DataTypesComboBoxItems
        {
            get { return _dataTypesComboBoxItems; }
            set
            {
                _dataTypesComboBoxItems = value;
                OnPropertyChanged(nameof(DataTypesComboBoxItems));

            }
        }

        private ComboBoxModel _selectedDataTypeComboBoxItem;
        public ComboBoxModel SelectedDataTypeComboBoxItem
        {
            get { return _selectedDataTypeComboBoxItem; }
            set
            {
                _selectedDataTypeComboBoxItem = value;
                OnPropertyChanged(nameof(SelectedDataTypeComboBoxItem));
                if (onlyKeyFiguresUpload)
                {
                    UpdateComboBoxItems();
                }
                UpdateModuleLabels();
                UpdateModuleKpi();
                UpdateAccountKpi();
                UpdateAccountGroupKpi();
                UpdateExtraFields();
                UpdateSecondComboBoxItems();
            }
        }

        private bool _isUpload;
        public bool isUpload
        {
            get { return _isUpload; }
            set
            {
                if (_isUpload != value)
                {
                    _isUpload = value;
                    OnPropertyChanged(nameof(isUpload));
                }
            }
        }

        public HomeViewModel()
        {
            List<ModuleModel> currentModules = ModuleList.GetModulesList();
            ModuleComboBoxItems = new ObservableCollection<ModuleModel>(currentModules);
            ModuleAccounts = new ObservableCollection<Account>();
            ModuleAccountGroups = new ObservableCollection<string>();
            SubModuleComboBoxItems = new ObservableCollection<ModuleModel>();
            ModuleKpis = new ObservableCollection<KpiViewModel>();

            var funds = ModuleList.GetFundValues();
            FundsComboBoxItems = new ObservableCollection<ComboBoxModel>(funds);

            var dataTypes = ModuleList.GetDataTypesList();
            DataTypesComboBoxItems = new ObservableCollection<ComboBoxModel>(dataTypes);

            var selectedDataType = dataTypes.FirstOrDefault();
            if (selectedDataType != null)
            {
                SelectedDataTypeComboBoxItem = selectedDataType;
            }

            var selectedFund = funds.FirstOrDefault();
            if (selectedFund != null)
            {
                SelectedFundComboBoxItem = selectedFund;
            }

            var valueTypes = ModuleList.GetValueTypeList();
            ValueTypeComboxBoxItems = new ObservableCollection<ComboBoxModel>(valueTypes);

            var selectedValueType = valueTypes.FirstOrDefault();
            if (selectedValueType != null)
            {
                SelectedValueTypeComboBoxItem = selectedValueType;
            }

            var dataOperationTypes = ModuleList.GetDataOperationTypesList();
            DataOperationTypeComboBoxItems = new ObservableCollection<ComboBoxModel>(dataOperationTypes);

            var selectedDataOperationType = dataOperationTypes.FirstOrDefault();
            if (selectedDataOperationType != null)
            {
                SelectedDataOperationTypeComboBoxItem = selectedDataOperationType;
            }

        }

        private void UpdateDataTypeComboBox()
        {
           if(SelectedDataOperationTypeComboBoxItem != null && SelectedDataOperationTypeComboBoxItem.Value == "U")
            {
                var dataTypes = ModuleList.GetDataTypesList();
                var currentDataTypes = dataTypes.Where(d => d.Value != "kpis").ToList();
                DataTypesComboBoxItems = new ObservableCollection<ComboBoxModel>(currentDataTypes);

                SelectedDataTypeComboBoxItem = currentDataTypes.FirstOrDefault();
            } else
            {
                var dataTypes = ModuleList.GetDataTypesList();
                DataTypesComboBoxItems = new ObservableCollection<ComboBoxModel>(dataTypes);

                SelectedDataTypeComboBoxItem = dataTypes.FirstOrDefault();
            }
            
            if (onlyKeyFiguresUpload) { 
                SelectedSubModuleComboBoxItem = null;
            } 
        }

        private void UpdateComboBoxItems()
        {
            if (onlyKeyFiguresUpload && SelectedDataOperationTypeComboBoxItem != null && SelectedDataOperationTypeComboBoxItem.Value == "U")
            {
                List<ModuleModel> currentModules = ModuleList.GetModulesList();
                ModuleComboBoxItems = new ObservableCollection<ModuleModel>(currentModules.Where(mod => mod.Label == "Portfolio"));
            }
            else
            {
                List<ModuleModel> currentModules = ModuleList.GetModulesList();
                ModuleComboBoxItems = new ObservableCollection<ModuleModel>(currentModules);
            }
        }

        private void UpdateSecondComboBoxItems()
        {
            if (SelectedModuleComboBoxItem != null && SelectedDataTypeComboBoxItem != null)
            {
                if (SelectedDataTypeComboBoxItem.Value == "fields")
                { 
                    List<ModuleModel> currentSubModules = SelectedModuleComboBoxItem.SubModules;

                    if(currentSubModules.Count > 0)
                    {
                        if(SelectedDataOperationTypeComboBoxItem != null)
                        {
                            List<Label> TenantLabels = Globals.ThisAddIn.GetUser().Tenant.Labels;

                            if (SelectedDataOperationTypeComboBoxItem.Value == "U")
                            {

                                currentSubModules = currentSubModules.Where(mod =>
                                {
                                    if (SelectedModuleComboBoxItem.Name == "Funds" && (mod.Name == "Accounts" || mod.Name == "Account Groups"))
                                    {
                                        return true;
                                    }

                                    List<Label> ModuleLabelsList = TenantLabels.Where(label => label.Module == SelectedModuleComboBoxItem.Name && label.Form == mod.Name && label.ChangeIn == "Forms").ToList();

                                    Label CurrentModuleLabel = ModuleLabelsList.FirstOrDefault();

                                    return CurrentModuleLabel != null && CurrentModuleLabel?.IsAddinUpload == true;
                                }).ToList();
                            } else if (SelectedDataOperationTypeComboBoxItem.Value == "D")
                            {
                                currentSubModules = currentSubModules.Where(mod =>
                                {

                                    if(SelectedModuleComboBoxItem.Name == "Funds" && (mod.Name == "Accounts" || mod.Name == "Account Groups"))
                                    {
                                        return true;
                                    }

                                    List<Label> ModuleLabelsList = TenantLabels.Where(label => label.Module == SelectedModuleComboBoxItem.Name && label.Form == mod.Name && label.ChangeIn == "Forms").ToList();

                                    Label CurrentModuleLabel = ModuleLabelsList.FirstOrDefault();

                                    return CurrentModuleLabel != null && CurrentModuleLabel?.IsAddin == true;
                                }).ToList();
                            }
                        }
                    }

                    if (onlyKeyFiguresUpload && SelectedDataOperationTypeComboBoxItem.Value == "U")
                    {
                        SubModuleComboBoxItems = new ObservableCollection<ModuleModel>(currentSubModules.Where(mod => mod.Label == "Key figures"));
                    }

                    else if (onlyKeyFiguresUpload && (SelectedModuleComboBoxItem.Label == "Limited Partners"|| SelectedModuleComboBoxItem.Label == "Funds"))
                    {
                        SubModuleComboBoxItems = new ObservableCollection<ModuleModel>(currentSubModules.Where(mod => mod.Label != "Capital calls breakdown" && mod.Label != "Distributions breakdown" && mod.Label != "Accounts" && mod.Label != "Account groups").ToList());

                    }
                    else if (SelectedDataOperationTypeComboBoxItem.Value == "D" && SelectedModuleComboBoxItem.Label == "Limited Partners")
                    {
                        SubModuleComboBoxItems = new ObservableCollection<ModuleModel>(currentSubModules.Where(mod => mod.Label != "Capital calls breakdown" && mod.Label != "Distributions breakdown").ToList());
                    }   
                    else
                    {
                        SubModuleComboBoxItems = new ObservableCollection<ModuleModel>(currentSubModules);
                    }
                }
                else if (SelectedDataTypeComboBoxItem.Value == "kpis")
                {
                    if (SelectedModuleComboBoxItem.Label == "Deal flow")
                    {
                        List<ModuleModel> options = new List<ModuleModel>();
                        options.Add(new ModuleModel() { Label = "Deal flow", Name = "deal-flow", SubModules = null });

                        SubModuleComboBoxItems = new ObservableCollection<ModuleModel>(options);

                    }
                    else if (SelectedModuleComboBoxItem.Label == "Portfolio")
                    {
                        List<ModuleModel> options = new List<ModuleModel>();
                        options.Add(new ModuleModel() { Label = "Operations", Name = "operations", SubModules = null });
                        options.Add(new ModuleModel() { Label = "Securities", Name = "securities", SubModules = null });
                        options.Add(new ModuleModel() { Label = "Ownership", Name = "ownership", SubModules = null });
                        options.Add(new ModuleModel() { Label = "Teams", Name = "teams", SubModules = null });

                        SubModuleComboBoxItems = new ObservableCollection<ModuleModel>(options);

                    }
                    else if (SelectedModuleComboBoxItem.Label == "Funds")
                    {
                        List<ModuleModel> options = new List<ModuleModel>();
                        options.Add(new ModuleModel() { Label = "Funds", Name = "funds", SubModules = null });
                        options.Add(new ModuleModel() { Label = "Shares", Name = "shares", SubModules = null });

                        SubModuleComboBoxItems = new ObservableCollection<ModuleModel>(options);
                    }
                    else if (SelectedModuleComboBoxItem.Label == "Limited Partners")
                    {
                        List<ModuleModel> options = new List<ModuleModel>();
                        options.Add(new ModuleModel() { Label = "Limited Partners", Name = "limited-partners", SubModules = null });
                        options.Add(new ModuleModel() { Label = "Shares", Name = "shares", SubModules = null });
                        options.Add(new ModuleModel() { Label = "Retail", Name = "Retail", SubModules = null });

                        SubModuleComboBoxItems = new ObservableCollection<ModuleModel>(options);
                    }
                    else if (SelectedModuleComboBoxItem.Label.ToLower() == "portfolio funds")
                    {
                        List<ModuleModel> options = new List<ModuleModel>();
                        options.Add(new ModuleModel() { Label = "Portfolio fund", Name = "portfolio-fund", SubModules = null });
                        options.Add(new ModuleModel() { Label = "Underlyings", Name = "underlyings", SubModules = null });

                        SubModuleComboBoxItems = new ObservableCollection<ModuleModel>(options);
                    }
                    else
                    {
                        SubModuleComboBoxItems = new ObservableCollection<ModuleModel>();
                    }
                }
                else
                {
                    SubModuleComboBoxItems = new ObservableCollection<ModuleModel>();
                }
            }
        }

        public void UpdateModuleLabels()
        {
            if (SelectedSubModuleComboBoxItem != null && SelectedModuleComboBoxItem != null && SelectedDataTypeComboBoxItem != null && SelectedDataTypeComboBoxItem.Value == "fields")
            {
                List<Label> TenantLabels = Globals.ThisAddIn.GetUser().Tenant.Labels;
                List<Label> ModuleLabelsList = TenantLabels.Where(label => label.Module == SelectedModuleComboBoxItem.Name && label.Form == SelectedSubModuleComboBoxItem.Name && label.ChangeIn == "Forms").ToList();

                Label CurrentModuleLabel = ModuleLabelsList.FirstOrDefault();

                if (CurrentModuleLabel != null && SelectedDataTypeComboBoxItem.Value == "fields")
                {
                    // Access the active worksheet
                    Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Worksheet;

                    // Get the range representing the first row
                    Range firstRowRange = (Range)activeSheet.Rows[1];
                    Range fullRange = (Range)firstRowRange.Cells[1, activeSheet.Columns.Count];

                    // Find the last column with data in the first row
                    int lastColumn = fullRange.End[XlDirection.xlToLeft].Column;

                    List<string> currentColumnsSlug = new List<string>();

                    // Iterate through each cell in the first row
                    for (int col = 1; col <= lastColumn; col++)
                    {
                        // Read the value of the cell
                        Range currentRange = (Range)activeSheet.Cells[1, col];
                        string cellValue = currentRange.Comment != null ? currentRange.Comment.Text() : null;

                        if (!String.IsNullOrWhiteSpace(cellValue))
                        {
                            currentColumnsSlug.Add(cellValue);
                        }
                    }

                    List<LabelViewModel> viewModelLabels = CurrentModuleLabel.Labels.Select(currentLabel => new LabelViewModel() { Name = currentLabel.LabelEn, Module = SelectedModuleComboBoxItem.Name, SubModule = SelectedSubModuleComboBoxItem.Name, Slug = currentLabel.Slug, Selection = currentColumnsSlug.Where(column => column == currentLabel.Slug).FirstOrDefault() != null }).ToList();
                    ModuleLabels = new ObservableCollection<LabelViewModel>(viewModelLabels);
                    return;
                }
            }

            ModuleLabels = new ObservableCollection<LabelViewModel>();
        }

        public void UpdateModuleKpi()
        {
            if (SelectedModuleComboBoxItem != null && SelectedDataTypeComboBoxItem != null && SelectedSubModuleComboBoxItem != null && SelectedDataTypeComboBoxItem.Value == "kpis")
            {
                List<Kpi> TenantKpis = Globals.ThisAddIn.GetUser().Kpis;
                List<string> CurrentModuleNameList = SelectedModuleComboBoxItem.Name.ToLower().Split(' ').ToList();
                string CurrentModuleName = String.Join("-", CurrentModuleNameList);
                string CurrentSubModule = SelectedSubModuleComboBoxItem.Name.ToLower();

                if (CurrentSubModule != null && CurrentModuleName == "deal-flow")
                {
                    CurrentSubModule = "deal-flow";
                }

                if (CurrentSubModule == "security")
                {
                    CurrentSubModule = "securities";
                }

                if (CurrentSubModule == "portfolio")
                {
                    CurrentSubModule = "";
                }

                if (CurrentSubModule == "operations")
                {
                    CurrentSubModule = "portfolio";
                }

                if (CurrentSubModule == "portfolio fund")
                {
                    CurrentSubModule = "portfolio-fund";
                }

                if (CurrentSubModule == "underlying investments")
                {
                    CurrentSubModule = "underlyings";
                }

                if (CurrentSubModule == "fund profile")
                {
                    CurrentSubModule = "funds";
                }

                if (CurrentSubModule == "lp details")
                {
                    CurrentSubModule = "limited-partners";
                }

                if (CurrentSubModule == "lp operations")
                {
                    CurrentSubModule = "shares";
                }

                if (CurrentSubModule != null && SelectedDataTypeComboBoxItem.Value == "kpis")
                {
                    List<Kpi> ModuleKpiList = TenantKpis.Where(label => label.Module == CurrentModuleName && label.SubModule == CurrentSubModule).ToList();

                    if (ModuleKpiList != null && ModuleKpiList.Count > 0)
                    {
                        // Access the active worksheet
                        Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Worksheet;

                        // Get the range representing the first row
                        Range firstRowRange = (Range)activeSheet.Rows[1];
                        Range fullRange = (Range)firstRowRange.Cells[1, activeSheet.Columns.Count];

                        // Find the last column with data in the first row
                        int lastColumn = fullRange.End[XlDirection.xlToLeft].Column;

                        List<string> currentColumnsSlug = new List<string>();

                        // Iterate through each cell in the first row
                        for (int col = 1; col <= lastColumn; col++)
                        {
                            // Read the value of the cell
                            Range currentRange = (Range)activeSheet.Cells[1, col];
                            string cellValue = currentRange.Comment != null ? currentRange.Comment.Text() : null;

                            if (!String.IsNullOrWhiteSpace(cellValue))
                            {
                                currentColumnsSlug.Add(cellValue);
                            }
                        }

                        List<KpiViewModel> viewModelKpi = ModuleKpiList.Select(currentKpi => new KpiViewModel() { Id = currentKpi.Id, Name = currentKpi.Name, Module = currentKpi.Module, SubModule = currentKpi.SubModule, Selection = currentColumnsSlug.Where(column => column == currentKpi.Id.ToString() + "-kpi").FirstOrDefault() != null }).ToList();
                        ModuleKpis = new ObservableCollection<KpiViewModel>(viewModelKpi);
                        return;
                    }
                }
            }

            ModuleKpis = new ObservableCollection<KpiViewModel>();
        }

        public void UpdateModuleAccounts()
        {
            List<Account> TenantAccounts = Globals.ThisAddIn.GetUser().Accounts;
            List<string> CurrentModuleNameList = SelectedModuleComboBoxItem.Name.ToLower().Split(' ').ToList();
            string CurrentModuleName = String.Join("-", CurrentModuleNameList);
            string CurrentSubModule = SelectedSubModuleComboBoxItem != null ? SelectedSubModuleComboBoxItem.Name.ToLower() : null;

            if (CurrentSubModule != null && CurrentSubModule == "accounts" && TenantAccounts.Count > 0)
            {
                // Access the active worksheet
                Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet as Worksheet;

                Range usedRange = activeSheet.UsedRange;

                int rowCount = usedRange.Rows.Count;
                int colCount = usedRange.Columns.Count;

                for (int i = 0; i <= rowCount; i++)
                {
                    for (int j = 0; j <= colCount; j++)
                    {
                        Range currentRange = usedRange.Cells[i, j] as Range;
                        if (currentRange != null)
                        {
                            string cellValue = currentRange.Value2 as string;

                            if (cellValue != null && !String.IsNullOrEmpty(cellValue))
                            {
                                var dropppedAccount = TenantAccounts.Where(account => account.AccountName == cellValue).FirstOrDefault();
                                if (dropppedAccount != null)
                                {
                                    //string period = SelectedPeriodComboBoxItem.Value;
                                    //string currentLabel = $"Funds:Accounts:value:{dropppedAccount.Id}:{accountType}:{period}";
                                }
                            }
                        }
                    }
                }
            }

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
                        ModuleAccountGroups = new ObservableCollection<string>(TenantAccountGroups.OrderBy(accountNumber => accountNumber).ToList());
                        return;
                    }
                }
            }

            ModuleAccountGroups = new ObservableCollection<string>();
        }

        public void UpdateExtraFields()
        {
            string CurrentSubModule = SelectedSubModuleComboBoxItem != null ? SelectedSubModuleComboBoxItem.Name.ToLower() : null;
            if (CurrentSubModule != null && (CurrentSubModule == "accounts" || CurrentSubModule == "account groups"))
            {
                AccountTypeComboBoxItems = new ObservableCollection<ComboBoxModel>(ModuleList.GetAccountTypesList());
                PeriodComboBoxItems = new ObservableCollection<ComboBoxModel>(ModuleList.GetPeriodList());
                return;
            }

            AccountTypeComboBoxItems = new ObservableCollection<ComboBoxModel>();
            PeriodComboBoxItems = new ObservableCollection<ComboBoxModel>();
        }

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
