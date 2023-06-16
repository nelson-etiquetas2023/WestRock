using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Reporting.WinForms;
using System.IO;
namespace WestRockDataPonchesPRO
{
    public partial class ReporteView : Form
    {
        public ReporteView()
        {
            InitializeComponent();
        }
        public ReportDataSource data = new ReportDataSource();
        public ReportParameter [] RpParams { get; set; }

        private void RepoprteView_Load(object sender, EventArgs e)
        {
            
            //this.reportViewer1.LocalReport.SetParameters(RpParams);
            
        }
    }
}
