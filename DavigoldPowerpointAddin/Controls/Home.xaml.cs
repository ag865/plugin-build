using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using DavigoldExcel.Models;
using DavigoldPowerpointAddin.ViewModel;

namespace DavigoldPowerpointAddin.Controls
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
                Console.WriteLine($"Count1: {count1}, Count2: {count2}");
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
        private HomeViewModel homeViewModel;
        private Point startPoint;

        public Home()
        {
            InitializeComponent();
            homeViewModel = new HomeViewModel();
            DataContext = homeViewModel;
        }

        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            LabelList.Items.Filter = FilterLabels;
            KpiList.Items.Filter = FilterKpiLabels;
            KeyFigureKpiList.Items.Filter = FilterKeyFigureKpiLabels;
            //ModuleComboBox.Items.Filter = FilterModules;
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
        
        private bool FilterKeyFigureKpiLabels(object obj)
        {
            KeyFigureKpiViewModel label = (KeyFigureKpiViewModel)obj;

            CultureInfo myCIintl = new CultureInfo("en-GB", false);
            return myCIintl.CompareInfo.IndexOf(label.Label, SearchTextBox.Text, CompareOptions.IgnoreCase) >= 0;
        }

        private void FeatureListView_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // Get current mouse position
            startPoint = e.GetPosition(null);
        }

        private static T FindAnchestor<T>(DependencyObject current) where T : DependencyObject
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
                // Get the dragged ListViewItemGeneral Partner
                ListView listView = sender as ListView;
                ListViewItem listViewItem = FindAnchestor<ListViewItem>((DependencyObject)e.OriginalSource);
                if (listViewItem == null) return;           // Abort
                                                            // Find the data behind the ListViewItem
                LabelViewModel label = (LabelViewModel)listView.ItemContainerGenerator.ItemFromContainer(listViewItem);
                if (label == null) return;                   // Abort
                                                             // Initialize the drag & drop operatGeneral Partnerion

                DragDrop.DoDragDrop(listViewItem, label.Name.ToString(), DragDropEffects.Copy);
            }

        }

        private void FeatureListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ListViewItem listViewItem = FindAnchestor<ListViewItem>((DependencyObject)e.OriginalSource);

            if (listViewItem != null)
            {
                LabelViewModel Label = listViewItem.DataContext as LabelViewModel;
                var dataContext = this.DataContext as HomeViewModel;

                var selectedUnderlyingSector = dataContext.SelectedUnderlyingSectorComboBoxItem;
                var selectedUnderlyingStatus = dataContext.SelectedUnderlyingStatusComboBoxItem;
                var selectedShare = dataContext.SelectedShareComboBoxItem;
                var selectedSecurityType = dataContext.SelectedSecurityTypeComboBoxItem;
                var selectedSecurityStatus = dataContext.SelectedSecurityStatusComboBoxItem;
                var selectedUnderlyingTableFilterComboBoxItem = dataContext.SelectedUnderlyingTableFilterComboBoxItem;
                var selectedKeyFigureTableComboBoxItem = dataContext.SelectedKeyFigureTableComboBoxItem;
                var selectedKeyFigureKpiComboBoxItem = dataContext.SelectedKeyFigureKpiComboBoxItem;
                var selectedFundHeader = dataContext.SelectedFundHeaderComboBoxItem;

                string LabelKey = "";
                string displayText = Label.Name;

                if (dataContext.SelectedSubModuleComboBoxItem.Label.ToLower() == "underlying investments" && Label.Slug == "top-15-underlying")
                {
                    LabelKey = $"{{{Label.Slug}}}";
                } else
                {
                    List<string> placeholderList = new List<string>();
                    placeholderList.Add(Label.Module.ToString());
                    placeholderList.Add(Label.Slug.ToString());
                    placeholderList.Add(Label.SubModule.ToString());

                    if (selectedUnderlyingSector != null && !String.IsNullOrEmpty(selectedUnderlyingSector.Value))
                    {
                        placeholderList.Add("underlyingsector-" + selectedUnderlyingSector.Value);
                    }

                    if (selectedUnderlyingStatus != null && !String.IsNullOrEmpty(selectedUnderlyingStatus.Value))
                    {
                        placeholderList.Add("underlyingstatus-" + selectedUnderlyingStatus.Value);
                    }

                    if (selectedShare != null && !String.IsNullOrEmpty(selectedShare.Value))
                    {
                        placeholderList.Add("share-" + selectedShare.Value);
                    }

                    if (selectedShare != null && !String.IsNullOrEmpty(selectedShare.Value))
                    {
                        placeholderList.Add("share-" + selectedShare.Value);
                    }

                    if (selectedSecurityStatus != null && !String.IsNullOrEmpty(selectedSecurityStatus.Value))
                    {
                        placeholderList.Add("status-" + selectedSecurityStatus.Value);
                    }

                    if (selectedSecurityType != null && !String.IsNullOrEmpty(selectedSecurityType.Value))
                    {
                        placeholderList.Add("securityType-" + selectedSecurityType.Value);
                    }

                    if (selectedUnderlyingTableFilterComboBoxItem != null && !String.IsNullOrEmpty(selectedUnderlyingTableFilterComboBoxItem.Value))
                    {
                        placeholderList.Add("extra-" + selectedUnderlyingTableFilterComboBoxItem.Value);
                    }

                    if(selectedFundHeader != null)
                    {
                        placeholderList.Add("fundheader-" + selectedFundHeader.Value);
                    }

                    if(selectedKeyFigureTableComboBoxItem != null && selectedKeyFigureKpiComboBoxItem != null && !String.IsNullOrEmpty(selectedKeyFigureTableComboBoxItem.Value) && !String.IsNullOrEmpty(selectedKeyFigureKpiComboBoxItem.Value))
                    {
                        placeholderList.Add("kftable-" + selectedKeyFigureTableComboBoxItem.Value);
                        placeholderList.Add("kfkpi-" + selectedKeyFigureKpiComboBoxItem.Value);
                    }

                    string placeholder = String.Join(":", placeholderList);

                    LabelKey = $"{{{placeholder}}}";
                }

               
                // Get the active presentation
                Microsoft.Office.Interop.PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;

                // Get the current selection
                Microsoft.Office.Interop.PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                // Check if the selection type is text
                if (selection.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText)
                {
                    // Get the current text range
                    Microsoft.Office.Interop.PowerPoint.TextRange textRange = selection.TextRange;
                    int originalLength = textRange.Text.Length;

                    // Insert your text at the end of the selection
                    textRange.InsertAfter(displayText);

                    Microsoft.Office.Interop.PowerPoint.TextRange newRange =
                       textRange.Characters(
                           originalLength + 1,     // 1-based index of first new char
                           displayText.Length      // length of the newly inserted text
                       );


                    // Configure the Click action on that range to be a hyperlink
                    var clickAction = newRange.ActionSettings[
                        Microsoft.Office.Interop.PowerPoint.PpMouseActivation.ppMouseClick
                    ];
                    clickAction.Action = Microsoft.Office.Interop.PowerPoint.PpActionType.ppActionHyperlink;
                    clickAction.Hyperlink.Address = LabelKey;    // the target of your hyperlink
                    //clickAction.Hyperlink.ScreenTip = $"{LabelKey}";
                }
            }
        }

        private void KpiListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ListViewItem listViewItem = FindAnchestor<ListViewItem>((DependencyObject)e.OriginalSource);
            var dataContext = this.DataContext as HomeViewModel;

            var selectedUnderlyingSector = dataContext.SelectedUnderlyingSectorComboBoxItem;
            var selectedUnderlyingStatus = dataContext.SelectedUnderlyingStatusComboBoxItem;
            var selectedShare = dataContext.SelectedShareComboBoxItem;
            var selectedSecurityType = dataContext.SelectedSecurityTypeComboBoxItem;
            var selectedSecurityStatus = dataContext.SelectedSecurityStatusComboBoxItem;

            if (listViewItem != null)
            {
                KpiViewModel Label = listViewItem.DataContext as KpiViewModel;

                List<string> placeholderList = new List<string>();
                placeholderList.Add("kpi");
                placeholderList.Add(Label.Id.ToString());
                placeholderList.Add(Label.Name);

                if (selectedUnderlyingSector != null && !String.IsNullOrEmpty(selectedUnderlyingSector.Value)) {
                    placeholderList.Add("underlyingsector-" + selectedUnderlyingSector.Value);
                }

                if (selectedUnderlyingStatus != null && !String.IsNullOrEmpty(selectedUnderlyingStatus.Value))
                {
                    placeholderList.Add("underlyingstatus-" + selectedUnderlyingStatus.Value);
                }

                if(selectedShare != null && !String.IsNullOrEmpty(selectedShare.Value))
                {
                    placeholderList.Add("share-" + selectedShare.Value);
                }

                if (selectedShare != null && !String.IsNullOrEmpty(selectedShare.Value))
                {
                    placeholderList.Add("share-" + selectedShare.Value);
                }

                if (selectedSecurityStatus != null && !String.IsNullOrEmpty(selectedSecurityStatus.Value))
                {
                    placeholderList.Add("status-" + selectedSecurityStatus.Value);
                }

                if (selectedSecurityType != null && !String.IsNullOrEmpty(selectedSecurityType.Value))
                {
                    placeholderList.Add("securityType-" + selectedSecurityType.Value);
                }

                string placeholder = String.Join(":", placeholderList);

                string LabelKey = $"{{{placeholder}}}";

                string displayText = Label.Name;

                // Get the active presentation
                Microsoft.Office.Interop.PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;

                // Get the current selection
                Microsoft.Office.Interop.PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                // Check if the selection type is text
                if (selection.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText)
                {
                    // Get the current text range
                    Microsoft.Office.Interop.PowerPoint.TextRange textRange = selection.TextRange;
                    int originalLength = textRange.Text.Length;

                    // Insert your text at the end of the selection
                    textRange.InsertAfter(displayText);

                    Microsoft.Office.Interop.PowerPoint.TextRange newRange =
                       textRange.Characters(
                           originalLength + 1,     // 1-based index of first new char
                           displayText.Length      // length of the newly inserted text
                       );


                    // Configure the Click action on that range to be a hyperlink
                    var clickAction = newRange.ActionSettings[
                        Microsoft.Office.Interop.PowerPoint.PpMouseActivation.ppMouseClick
                    ];
                    clickAction.Action = Microsoft.Office.Interop.PowerPoint.PpActionType.ppActionHyperlink;
                    clickAction.Hyperlink.Address = LabelKey;    // the target of your hyperlink
                    clickAction.Hyperlink.ScreenTip = $"{LabelKey}";
                }
            }
        }

        private void KeyFigureKpiList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ListViewItem listViewItem = FindAnchestor<ListViewItem>((DependencyObject)e.OriginalSource);
            var dataContext = this.DataContext as HomeViewModel;

            var selectedType = dataContext.SelectedKeyFigureTypeComboBoxItem;
            var selectedDate = dataContext.SelectedKeyFigureDateComboBoxItem;

            if (selectedDate == null) {
                MessageBox.Show("Please Select Date");
                return;
            }

            if (selectedType == null)
            {
                MessageBox.Show("Please Select Type");
                return;
            }

            if (listViewItem != null && selectedType != null && selectedDate != null)
            {
                KeyFigureKpiViewModel Label = listViewItem.DataContext as KeyFigureKpiViewModel;

                string LabelKey = $"{{Portfolio:KeyFigureSpecific:{Label.Id.ToString()}:{selectedType.Value}:{selectedDate.Value}}}";

                string displayText = Label.Name;

                // Get the active presentation
                Microsoft.Office.Interop.PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;

                // Get the current selection
                Microsoft.Office.Interop.PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                // Check if the selection type is text
                if (selection.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText)
                {
                    // Get the current text range
                    Microsoft.Office.Interop.PowerPoint.TextRange textRange = selection.TextRange;
                    int originalLength = textRange.Text.Length;

                    // Insert your text at the end of the selection
                    textRange.InsertAfter(displayText);

                    Microsoft.Office.Interop.PowerPoint.TextRange newRange =
                       textRange.Characters(
                           originalLength + 1,     // 1-based index of first new char
                           displayText.Length      // length of the newly inserted text
                       );


                    // Configure the Click action on that range to be a hyperlink
                    var clickAction = newRange.ActionSettings[
                        Microsoft.Office.Interop.PowerPoint.PpMouseActivation.ppMouseClick
                    ];
                    clickAction.Action = Microsoft.Office.Interop.PowerPoint.PpActionType.ppActionHyperlink;
                    clickAction.Hyperlink.Address = LabelKey;    // the target of your hyperlink
                    clickAction.Hyperlink.ScreenTip = $"{LabelKey}";
                }
            }
        }

        private void ChartList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ListViewItem listViewItem = FindAnchestor<ListViewItem>((DependencyObject)e.OriginalSource);
            var dataContext = this.DataContext as HomeViewModel;

            var selectedDate = dataContext.SelectedChartDateComboBoxItem;
            var selectedShare = dataContext.SelectedShareComboBoxItem;

            if (selectedDate == null)
            {
                MessageBox.Show("Please Select Date");
                return;
            }


            if (listViewItem != null && selectedDate != null)
            {
                ChartViewModel Label = listViewItem.DataContext as ChartViewModel;

                List<string> placeholderList = new List<string>();
                placeholderList.Add("chart");
                placeholderList.Add(Label.Id.ToString());
                placeholderList.Add(Label.Name);
                placeholderList.Add(selectedDate.Value);


                if (selectedShare != null && !String.IsNullOrEmpty(selectedShare.Value))
                {
                    placeholderList.Add("share-" + selectedShare.Value);
                }

                string placeholder = String.Join(":", placeholderList);

                string LabelKey = $"{{{placeholder}}}";

                string displayText = Label.Name;

                // Get the active presentation
                Microsoft.Office.Interop.PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;

                // Get the current selection
                Microsoft.Office.Interop.PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                // Check if the selection type is text
                if (selection.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText)
                {
                    // Get the current text range
                    Microsoft.Office.Interop.PowerPoint.TextRange textRange = selection.TextRange;

                    int originalLength = textRange.Text.Length;

                    // Insert your text at the end of the selection
                    textRange.InsertAfter(displayText);

                    Microsoft.Office.Interop.PowerPoint.TextRange newRange =
                     textRange.Characters(
                         originalLength + 1,     // 1-based index of first new char
                         displayText.Length      // length of the newly inserted text
                     );


                    // Configure the Click action on that range to be a hyperlink
                    var clickAction = newRange.ActionSettings[
                        Microsoft.Office.Interop.PowerPoint.PpMouseActivation.ppMouseClick
                    ];
                    clickAction.Action = Microsoft.Office.Interop.PowerPoint.PpActionType.ppActionHyperlink;
                    clickAction.Hyperlink.Address = LabelKey;    // the target of your hyperlink
                    clickAction.Hyperlink.ScreenTip = $"{LabelKey}";
                }
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

            var selectedPeriod = dataContext.SelectedPeriodComboBoxItem;

            if (selectedPeriod == null)
            {
                MessageBox.Show("Please Select Period");
                return;
            }


            var selectedeAccounts = dataContext.ModuleAccounts.Where(account => account.Selection == true && account.AccountType != null).ToList();

            if (selectedeAccounts.Count > 0) 
            {

                List<string> accountList = new List<string>();

                foreach (var account in selectedeAccounts)
                {
                    accountList.Add($"Funds:Accounts:{account.Id}:{account.AccountType.Value}:{selectedPeriod.Value}:{account.OperatorType.Value}");
                }

                string placeholder = String.Join(",", accountList);

                string LabelKey = $"{{{placeholder}}}";

                string displayText = "Funds Accounts";

                // Get the active presentation
                Microsoft.Office.Interop.PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;

                // Get the current selection
                Microsoft.Office.Interop.PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                // Check if the selection type is text
                if (selection.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText)
                {
                    // Get the current text range
                    Microsoft.Office.Interop.PowerPoint.TextRange textRange = selection.TextRange;
                    int originalLength = textRange.Text.Length;

                    // Insert your text at the end of the selection
                    textRange.InsertAfter(LabelKey);

                    Microsoft.Office.Interop.PowerPoint.TextRange newRange =
                        textRange.Characters(
                            originalLength + 1,     // 1-based index of first new char
                            displayText.Length      // length of the newly inserted text
                        );


                    // Configure the Click action on that range to be a hyperlink
                    var clickAction = newRange.ActionSettings[
                        Microsoft.Office.Interop.PowerPoint.PpMouseActivation.ppMouseClick
                    ];
                    clickAction.Action = Microsoft.Office.Interop.PowerPoint.PpActionType.ppActionHyperlink;
                    clickAction.Hyperlink.Address = LabelKey;    // the target of your hyperlink
                    clickAction.Hyperlink.ScreenTip = $"{LabelKey}";


                }

                //dataContext.ResetAccounts();
            }
        }

        private void AccountGroupList_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ListViewItem listViewItem = FindAnchestor<ListViewItem>((DependencyObject)e.OriginalSource);
            var dataContext = this.DataContext as HomeViewModel;

            var selectedPeriod = dataContext.SelectedPeriodComboBoxItem;

            if (selectedPeriod == null)
            {
                MessageBox.Show("Please Select Period");
                return;
            }

            var selectedeAccounts = dataContext.ModuleAccountGroups.Where(account => account.Selection == true && account.AccountType != null).ToList();

            if (selectedeAccounts.Count > 0)
            {

                List<string> accountList = new List<string>();

                foreach (var account in selectedeAccounts)
                {
                    accountList.Add($"Funds:AccountGroups:{account.Name}:{account.AccountType.Value}:{selectedPeriod.Value}:{account.OperatorType.Value}");
                }

                string placeholder = String.Join(",", accountList);

                string LabelKey = $"{{{placeholder}}}";

                string displayText = "Funds Account Groups";

                // Get the active presentation
                Microsoft.Office.Interop.PowerPoint.Presentation presentation = Globals.ThisAddIn.Application.ActivePresentation;

                // Get the current selection
                Microsoft.Office.Interop.PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;

                // Check if the selection type is text
                if (selection.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionText)
                {
                    // Get the current text range
                    Microsoft.Office.Interop.PowerPoint.TextRange textRange = selection.TextRange;
                    int originalLength = textRange.Text.Length;


                    // Insert your text at the end of the selection
                    textRange.InsertAfter(LabelKey);

                    Microsoft.Office.Interop.PowerPoint.TextRange newRange =
                        textRange.Characters(
                            originalLength + 1,     // 1-based index of first new char
                            displayText.Length      // length of the newly inserted text
                        );


                    // Configure the Click action on that range to be a hyperlink
                    var clickAction = newRange.ActionSettings[
                        Microsoft.Office.Interop.PowerPoint.PpMouseActivation.ppMouseClick
                    ];
                    clickAction.Action = Microsoft.Office.Interop.PowerPoint.PpActionType.ppActionHyperlink;
                    clickAction.Hyperlink.Address = LabelKey;    // the target of your hyperlink
                    clickAction.Hyperlink.ScreenTip = $"{LabelKey}";
                }

                //dataContext.ResetAccounts();
            }
        }

    }
}
