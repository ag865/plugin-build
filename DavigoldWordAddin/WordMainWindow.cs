using DavigoldWordAddin.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DavigoldWordAddin
{
    public partial class WordMainWindow : UserControl
    {
        public WordMainWindow()
        {
            InitializeComponent();
        }

        public void ShowTokenPage()
        {
            this.elementHost1.Child = new UserTokenForm();
        }

        public void ShowHomePage()
        {
            this.elementHost1.Child = new Home();
        }
    }
}
