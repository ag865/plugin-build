using DavigoldWordAddin.ViewModel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace DavigoldWordAddin.Controls
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
    /// <summary>
    /// Interaction logic for Home.xaml
    /// </summary>
    public partial class Home : UserControl
    {
        private HomeViewModel homeViewModel;
        private System.Windows.Point startPoint;

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
             System.Windows.Point mousePos = e.GetPosition(null);
            Vector diff = startPoint - mousePos;

            if (e.LeftButton == MouseButtonState.Pressed &&
                (Math.Abs(diff.X) > SystemParameters.MinimumHorizontalDragDistance ||
                       Math.Abs(diff.Y) > SystemParameters.MinimumVerticalDragDistance))
            {
                // Get the dragged ListViewItemGeneral Partner
                ListView listView = sender as ListView;
                
                ListViewItem listViewItem = FindAnchestor<ListViewItem>((DependencyObject)e.OriginalSource);
                if (listViewItem == null) return;           

                LabelViewModel label = (LabelViewModel)listView.ItemContainerGenerator.ItemFromContainer(listViewItem);
                if (label == null) return;                   

                DragDrop.DoDragDrop(listViewItem, label.Name.ToString(), DragDropEffects.Copy);
            }

        }

        private void FeatureListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ListViewItem listViewItem = FindAnchestor<ListViewItem>((DependencyObject)e.OriginalSource);

            if (listViewItem != null)
            {
                LabelViewModel Label = listViewItem.DataContext as LabelViewModel;

                string LabelKey = "{" + Label.SubModule.ToString() + ":" + Label.Slug.ToString() + "}";
                string LabelName = Label.Name.ToString();

                // Get the active document
                Microsoft.Office.Interop.Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;

                // Get the current selection
                Microsoft.Office.Interop.Word.Selection selection = Globals.ThisAddIn.Application.Selection;

                try
                { 
                    int start = selection.Start;
                     
                    selection.TypeText(LabelKey);
                     
                    int end = selection.Start;
                     
                    Range insertedRange = document.Range(start, end);
                     
                    Microsoft.Office.Interop.Word.Hyperlink hyperlink = document.Hyperlinks.Add(
                        Anchor: insertedRange,
                        Address: LabelKey,
                        ScreenTip: null,
                        TextToDisplay: LabelName
                    );
                     
                    insertedRange.Font.Underline = WdUnderline.wdUnderlineNone;
                    insertedRange.Font.Color = WdColor.wdColorAutomatic;
                    selection.SetRange(end, end);
                }
                catch (Exception ex)
                {
                    // Handle exceptions
                    System.Windows.Forms.MessageBox.Show("An error occurred: " + ex.Message);
                }
            }
        }

        private void KpiListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ListViewItem listViewItem = FindAnchestor<ListViewItem>((DependencyObject)e.OriginalSource);

            if (listViewItem != null)
            {
                KpiViewModel Label = listViewItem.DataContext as KpiViewModel;

                string LabelKey = $"{{kpi:{Label.Id}:{Label.Name}}}";
                string LabelName = Label.Name.ToString();
                 
                Microsoft.Office.Interop.Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
                 
                Microsoft.Office.Interop.Word.Selection selection = Globals.ThisAddIn.Application.Selection;

                try
                { 
                    int start = selection.Start;
                     
                    selection.TypeText(LabelKey);
                     
                    int end = selection.Start;
                     
                    Range insertedRange = document.Range(start, end);
                     
                    Microsoft.Office.Interop.Word.Hyperlink hyperlink = document.Hyperlinks.Add(
                        Anchor: insertedRange,
                        Address: LabelKey,
                        ScreenTip: null,
                        TextToDisplay: LabelName
                    );
                     
                    insertedRange.Font.Underline = WdUnderline.wdUnderlineNone;
                    insertedRange.Font.Color = WdColor.wdColorAutomatic;
                    selection.SetRange(end, end);
                }
                catch (Exception ex)
                {
                    // Handle exceptions
                    System.Windows.Forms.MessageBox.Show("An error occurred: " + ex.Message);
                }
            }
        }
    }
}
