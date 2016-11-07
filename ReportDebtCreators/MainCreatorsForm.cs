using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ReportDebtCreators.enginer;
using ReportDebtCreators.model;

namespace ReportDebtCreators
{
    public partial class MainCreatorsForm : Form
    {
        public MainCreatorsForm(IReadOnlyCollection<StructExelModel> temp, IReadOnlyCollection<StructExelModel> pack)
        {

            InitializeComponent();

            TemplateLasts.DataSource = temp;
            TemplateLasts.DisplayMember = "Name";
            TemplateLasts.ValueMember = "AbsolutPatch";

            PackageLasts.DataSource = pack;
            PackageLasts.DisplayMember = "Name";
            PackageLasts.ValueMember = "AbsolutPatch";
        }

        private void TemplateLasts_SelectedValueChanged(object sender, EventArgs e)
        {
            //MessageBox.Show("Changed");
        }

        private void TemplateLasts_SelectedIndexChanged(object sender, EventArgs e)
        {
            //MessageBox.Show("Changed");
        }

        private void CloseApp_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
