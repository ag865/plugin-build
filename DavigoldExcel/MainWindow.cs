using DavigoldExcel.Controls;
using System.Windows.Forms;

namespace DavigoldExcel
{
    public partial class MainWindow : UserControl
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        public void ShowTokenPage()
        {
            this.mainFormHost.Child = new UserTokenForm();
        }

        public void ShowHomePage()
        {
            this.mainFormHost.Child = new Home();
        }

    }
}
