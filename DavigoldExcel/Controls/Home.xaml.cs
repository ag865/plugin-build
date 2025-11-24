using DavigoldExcel.ViewModel;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using ExcelWorkbook = Microsoft.Office.Interop.Excel.Workbook;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;
using ExcelSheet = Microsoft.Office.Interop.Excel.Worksheet;
using DavigoldExcel.Service;
using System.Globalization;
using ModuleModel = DavigoldExcel.Models.Module;
using Label = DavigoldExcel.Models.Label;
using System.Collections.Generic;
using System.Linq;
using CustomTaskPane = Microsoft.Office.Tools.CustomTaskPane;
using DavigoldExcel.Models;

namespace DavigoldExcel.Controls
{
    public class CountToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is int count)
            {
                return count > 0 ? Visibility.Visible : Visibility.Collapsed;
            }

            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class MultiListCountToVisibilityConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values != null && values.Length >= 2 &&
                values[0] is int count1 && values[1] is int count2)
            {
                return (count1 > 0 || count2 > 0) ? Visibility.Visible : Visibility.Collapsed;
            }
            return Visibility.Collapsed;
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    /// <summary>
    /// Interaction logic for Home.xaml
    /// </summary>
    public partial class Home : UserControl
    {

        private Point startPoint;
        private HomeViewModel homeViewModel;

        public static HomeViewModel Instance { get; private set; }
        
        // Public property to access FilterDatePicker
        public System.Windows.Controls.DatePicker GetFilterDatePicker()
        {
            return FilterDatePicker;
        }

        public Home()
        {
            InitializeComponent();
            InitializeData();
            InitializeLanguage();

            Globals.ThisAddIn.Application.SheetChange += Workbook_SheetChange;
            Globals.ThisAddIn.Application.SheetActivate += Workbook_SheetActivate;
        }

        private void InitializeData()
        {
            homeViewModel = new HomeViewModel();
            DataContext = homeViewModel;

            Instance = homeViewModel;
        }

        private void InitializeLanguage()
        {
            LanguageToggle.IsChecked = homeViewModel.IsEnglish;
            UpdateLanguageText();
        }

        private void UpdateLanguageText()
        {
            LanguageText.Text = homeViewModel.IsEnglish ? "English" : "Français";
        }

        private void LanguageToggle_Checked(object sender, RoutedEventArgs e)
        {
            homeViewModel.IsEnglish = true;
            UpdateLanguageText();
            // Add any additional language-specific logic here
        }

        private void LanguageToggle_Unchecked(object sender, RoutedEventArgs e)
        {
            homeViewModel.IsEnglish = false;
            UpdateLanguageText();
            // Add any additional language-specific logic here
        }

        private void Workbook_SheetChange(object sheet, ExcelRange range)
        {
            CustomTaskPane mainTaskPane = Globals.ThisAddIn.GetCurrentTaskPane();

            var dataContext = this.DataContext as HomeViewModel;

            if (mainTaskPane != null && mainTaskPane.Visible)
            {
                var selectedValueType = dataContext.SelectedValueTypeComboBoxItem;
                var isUpload = dataContext.SelectedDataOperationTypeComboBoxItem;

                if (homeViewModel.ModuleLabels != null && selectedValueType != null && homeViewModel.ModuleLabels.Count > 0)
                {
                    ExcelService.SyncColumnsComments(homeViewModel.ModuleLabels, selectedValueType, isUpload.Value);
                }

                if (homeViewModel.ModuleKpis != null && selectedValueType != null && homeViewModel.ModuleKpis.Count > 0)
                {
                    string date = "";
                    if (MyDatePicker.SelectedDate != null)
                    {
                        date = MyDatePicker.SelectedDate.Value.ToString("dd/MM/yyyy");
                    }
                    ExcelService.SyncColumnsKpis(homeViewModel.ModuleKpis, selectedValueType, date);
                }
            }
        }

        private void Workbook_SheetActivate(object sheet)
        {
            //homeViewModel.UpdateModuleLabels();
            //homeViewModel.UpdateModuleKpi();
        }

        private void FeatureListView_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // Get current mouse position
            startPoint = e.GetPosition(null);
        }

        private static T FindAnchestor<T>(DependencyObject current)
            where T : DependencyObject
        {
            do
            {
                if (current is T)
                {
                    return (T)current;
                }
                current = VisualTreeHelper.GetParent(current);
            }
            while (current != null);
            return null;
        }

        private void FeatureListView_MouseMove(object sender, MouseEventArgs e)
        {
            // Get the current mouse position
            Point mousePos = e.GetPosition(null);
            Vector diff = startPoint - mousePos;

            if (e.LeftButton == MouseButtonState.Pressed &&
                (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance ||
                       Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance))
            {
                // Get the dragged ListViewItem
                ListView listView = sender as ListView;
                ListViewItem listViewItem = FindAnchestor<ListViewItem>((DependencyObject)e.OriginalSource);
                if (listViewItem == null) return;           // Abort
                                                            // Find the data behind the ListViewItem
                LabelViewModel label = (LabelViewModel)listView.ItemContainerGenerator.ItemFromContainer(listViewItem);
                if (label == null) return;                   // Abort
                                                             // Initialize the drag & drop operation

                DragDrop.DoDragDrop(listViewItem, label.Name.ToString(), DragDropEffects.Copy);
            }
        }

        private void ListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ListViewItem listViewItem = FindAnchestor<ListViewItem>((DependencyObject)e.OriginalSource);
            var dataContext = this.DataContext as HomeViewModel;

            var selectedValueType = dataContext.SelectedValueTypeComboBoxItem;
            var isUpload = dataContext.SelectedDataOperationTypeComboBoxItem;
            string downloadValue = isUpload.Value;

            if (listViewItem != null && selectedValueType != null)
            {
                LabelViewModel Label = listViewItem.DataContext as LabelViewModel;
                ExcelRange selectedRange = (ExcelRange)Globals.ThisAddIn.Application.Selection;

                if (selectedRange != null)
                {
                    selectedRange.Value2 = Label.Name;

                    string currentComment = selectedRange.Comment != null ? selectedRange.Comment.Text() : null;
                    string slug = $"{downloadValue}:{Label.Module}:{Label.SubModule}:{Label.Slug}:{selectedValueType.Value}";

                    if (currentComment != null && !String.IsNullOrEmpty(currentComment))
                    {
                        selectedRange.ClearComments();
                    }

                    selectedRange.AddComment(slug);
                    // Move to the next cell to the right
                    ExcelRange nextCell = selectedRange.Offset[0, 1];
                    nextCell.Select();
                }
            }
        }

        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            LabelList.Items.Filter = FilterLabels;
            KpiList.Items.Filter = FilterKpiLabels;
            AccountList.Items.Filter = FilterAccountsLabels;
            AccountGroupList.Items.Filter = FilterAccountGroupsLabels;
        }

        private bool FilterLabels(object obj)
        {
            LabelViewModel label = (LabelViewModel)obj;

            CultureInfo myCIintl = new CultureInfo("en-GB", false);
            return myCIintl.CompareInfo.IndexOf(label.Name, SearchTextBox.Text, CompareOptions.IgnoreCase) >= 0;
        }

        private bool FilterKpiLabels(object obj)
        {
            KpiViewModel label = (KpiViewModel)obj;

            CultureInfo myCIintl = new CultureInfo("en-GB", false);
            return myCIintl.CompareInfo.IndexOf(label.Name, SearchTextBox.Text, CompareOptions.IgnoreCase) >= 0;
        }

        private bool FilterAccountsLabels(object obj)
        {
            Account account = (Account)obj;

            CultureInfo myCIintl = new CultureInfo("en-GB", false);
            return myCIintl.CompareInfo.IndexOf(account.AccountName, SearchTextBox.Text, CompareOptions.IgnoreCase) >= 0 || myCIintl.CompareInfo.IndexOf(account.AccountNumber, SearchTextBox.Text, CompareOptions.IgnoreCase) == 0;
        }

        private bool FilterAccountGroupsLabels(object obj)
        {
            string accountGroup = (string)obj;

            CultureInfo myCIintl = new CultureInfo("en-GB", false);
            return myCIintl.CompareInfo.IndexOf(accountGroup, SearchTextBox.Text, CompareOptions.IgnoreCase) == 0;
        }

        private bool FilterModules(object obj)
        {
            ModuleModel module = (ModuleModel)obj;

            CultureInfo myCIintl = new CultureInfo("en-GB", false);

            List<Label> Labels = Globals.ThisAddIn.GetUser().Tenant.Labels.Where(label => label.Module == module.Name && label.Labels.Where(moduleLabel => moduleLabel != null && moduleLabel.NameEn != null && myCIintl.CompareInfo.IndexOf(moduleLabel.NameEn, SearchTextBox.Text, CompareOptions.IgnoreCase) >= 0).ToList().Count() > 0).ToList();


            return Labels.Count > 0;
        }

        private void KpiList_MouseMove(object sender, MouseEventArgs e)
        {
            // Get the current mouse position
            Point mousePos = e.GetPosition(null);
            Vector diff = startPoint - mousePos;

            if (e.LeftButton == MouseButtonState.Pressed &&
                (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance ||
                       Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance))
            {
                // Get the dragged ListViewItem
                ListView listView = sender as ListView;
                ListViewItem listViewItem = FindAnchestor<ListViewItem>((DependencyObject)e.OriginalSource);
                if (listViewItem == null) return;           // Abort
                                                            // Find the data behind the ListViewItem
                KpiViewModel label = (KpiViewModel)listView.ItemContainerGenerator.ItemFromContainer(listViewItem);
                if (label == null) return;                   // Abort
                                                             // Initialize the drag & drop operation

                DragDrop.DoDragDrop(listViewItem, label.Name.ToString(), DragDropEffects.Copy);
            }
        }

        private void KpiList_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // Get current mouse position
            startPoint = e.GetPosition(null);
        }

        private void KpiList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ListViewItem listViewItem = FindAnchestor<ListViewItem>((DependencyObject)e.OriginalSource);
            DateTime? selectedDate = MyDatePicker.SelectedDate;
            if (listViewItem != null)
            {
                KpiViewModel Label = listViewItem.DataContext as KpiViewModel;
                var dataContext = this.DataContext as HomeViewModel;

                var selectedValueType = dataContext.SelectedValueTypeComboBoxItem;
                ExcelRange selectedRange = (ExcelRange)Globals.ThisAddIn.Application.Selection;

                if (selectedRange != null && selectedValueType != null)
                {
                    selectedRange.Value2 = Label.Name;

                    string currentComment = selectedRange.Comment != null ? selectedRange.Comment.Text() : null;
                    string slug = $"{Label.Id}-kpi:{selectedValueType.Value}";

                    if (selectedDate != null)
                    {
                        slug += ":" + selectedDate.Value.ToString("dd/MM/yyyy");
                    }

                    if (currentComment != null && !String.IsNullOrEmpty(currentComment))
                    {
                        selectedRange.ClearComments();
                    }

                    selectedRange.AddComment(slug);
                    // Move to the next cell to the right
                    ExcelRange nextCell = selectedRange.Offset[0, 1];
                    nextCell.Select();

                }

                //if (!Label.Selection)
                //{
                //    ExcelService.AddKpiAtLastPosition(Label);
                //}
                //else
                //{
                //    ExcelService.RemoveKpi(Label);
                //}
            }
        }

        private void AccountList_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // Get current mouse position
            startPoint = e.GetPosition(null);
        }

        private void AccountList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ListViewItem listViewItem = FindAnchestor<ListViewItem>((DependencyObject)e.OriginalSource);
            var dataContext = this.DataContext as HomeViewModel;

            var selectedValueType = dataContext.SelectedValueTypeComboBoxItem;

            if (selectedValueType == null)
            {
                MessageBox.Show("Please Selected Value Type.");
                return;
            }

            if (dataContext.SelectedAccountTypeComboBoxItem == null)
            {
                MessageBox.Show("Please Selected Debit/Credit.");
                return;
            }

            if (dataContext.SelectedPeriodComboBoxItem == null)
            {
                MessageBox.Show("Please Selected Period.");
                return;
            }

            if (listViewItem != null)
            {
                Account SelectedAccount = listViewItem.DataContext as Account;
                ExcelRange selectedRange = (ExcelRange)Globals.ThisAddIn.Application.Selection;

                if (selectedRange != null)
                {
                    string accountType = dataContext.SelectedAccountTypeComboBoxItem.Value == "debit" ? "D" : "C";
                    string period = dataContext.SelectedPeriodComboBoxItem.Value;
                    string currentValue = selectedRange.Value2 as string;
                    if (currentValue != null && !String.IsNullOrEmpty(currentValue))
                    {
                        List<string> currentAccounts = currentValue.Split('+').ToList();
                        currentAccounts.Add(accountType + SelectedAccount.AccountNumber);
                        selectedRange.Value2 = String.Join("+", currentAccounts);
                    }
                    else
                    {
                        selectedRange.Value2 = accountType + SelectedAccount.AccountNumber;
                    }

                    string currentComment = selectedRange.Comment != null ? selectedRange.Comment.Text() : null;
                    string currentLabel = $"Funds:Accounts:{SelectedAccount.Id}:{selectedValueType.Value}:{accountType}:{period}";

                    if (currentComment != null && !String.IsNullOrEmpty(currentComment))
                    {
                        List<string> currentAccounts = currentComment.Split(',').ToList();
                        currentAccounts.Add(currentLabel);
                        string commentString = String.Join(",", currentAccounts);
                        selectedRange.Comment.Text(commentString);
                    }
                    else
                    {
                        selectedRange.AddComment(currentLabel);
                    }
                }
            }


        }

        private void AccountGroupList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ListViewItem listViewItem = FindAnchestor<ListViewItem>((DependencyObject)e.OriginalSource);
            var dataContext = this.DataContext as HomeViewModel;

            var selectedValueType = dataContext.SelectedValueTypeComboBoxItem;

            if (selectedValueType == null)
            {
                MessageBox.Show("Please Selected Value Type.");
                return;
            }

            if (dataContext.SelectedAccountTypeComboBoxItem == null)
            {
                MessageBox.Show("Please Selected Debit/Credit.");
                return;
            }

            if (dataContext.SelectedPeriodComboBoxItem == null)
            {
                MessageBox.Show("Please Selected Period.");
                return;
            }

            if (listViewItem != null)
            {
                string SelectedAccountGroup = listViewItem.DataContext as string;
                ExcelRange selectedRange = (ExcelRange)Globals.ThisAddIn.Application.Selection;

                if (selectedRange != null)
                {
                    string accountType = dataContext.SelectedAccountTypeComboBoxItem.Value == "debit" ? "D" : "C";
                    string period = dataContext.SelectedPeriodComboBoxItem.Value;
                    string currentValue = selectedRange.Value2 as string;
                    if (currentValue != null && !String.IsNullOrEmpty(currentValue))
                    {
                        List<string> currentAccounts = currentValue.Split('+').ToList();
                        currentAccounts.Add(accountType + SelectedAccountGroup);
                        selectedRange.Value2 = String.Join("+", currentAccounts);
                    }
                    else
                    {
                        selectedRange.Value2 = accountType + SelectedAccountGroup;
                    }

                    string currentComment = selectedRange.Comment != null ? selectedRange.Comment.Text() : null;
                    string currentLabel = $"Funds:AccountGroups:{SelectedAccountGroup}:{selectedValueType.Value}:{accountType}:{period}";

                    if (currentComment != null && !String.IsNullOrEmpty(currentComment))
                    {
                        List<string> currentAccounts = currentComment.Split(',').ToList();
                        currentAccounts.Add(currentLabel);
                        string commentString = String.Join(",", currentAccounts);
                        selectedRange.Comment.Text(commentString);
                    }
                    else
                    {
                        selectedRange.AddComment(currentLabel);
                    }
                }
            }

        }
    }
}
