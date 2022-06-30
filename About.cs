using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;

namespace ExRaspViewer
{
    public partial class About : Form
    {
        public About()
        {
            InitializeComponent();
            lbVersion.Text = AppVersion();
        }

        private string AppVersion()
        {
            return Assembly.GetExecutingAssembly().GetName().Version.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
