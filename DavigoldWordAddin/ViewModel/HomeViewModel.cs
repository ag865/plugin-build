using DavigoldExcel.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ModuleModel = DavigoldExcel.Models.Module;

namespace DavigoldWordAddin.ViewModel
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

        public int Id { set; get; }

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
            }
        }

        public HomeViewModel()
        {
            List<ModuleModel> currentModules = ModuleList.GetModulesList();
            ModuleComboBoxItems = new ObservableCollection<ModuleModel>(currentModules);
            SubModuleComboBoxItems = new ObservableCollection<ModuleModel>();
            ModuleKpis = new ObservableCollection<KpiViewModel>();
        }

        private void UpdateSecondComboBoxItems()
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

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }

}
