using DavigoldPowerpointAddin.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DavigoldPowerpointAddin
{
    public partial class PPMainWindow : UserControl
    {
        public PPMainWindow()
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
