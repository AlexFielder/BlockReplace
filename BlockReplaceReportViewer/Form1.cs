using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BlockReplaceReportViewer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            ds.ReadXml(@"C:\temp\Drawings.xml");
            ds.WriteXmlSchema(@"C:\temp\Drawings.xsd");
            //ds = null;
            DataSet snds = new DataSet();
            snds.ReadXml(@"C:\temp\Snapshots.xml");
            snds.WriteXmlSchema(@"C:\temp\Snapshots.xsd");
            DrawingBindingSource.DataSource = ds;
            //snapshotsBindingSource.DataSource = snds;
            //reportViewer1.LocalReport.SubreportProcessing += new Microsoft.Reporting.WinForms.SubreportProcessingEventHandler(SetSubDataSource);
            //SnapshotBindingSource.DataSource = ds.Tables["snapshot"];
            this.reportViewer1.RefreshReport();
        }

        //private void SetSubDataSource(object sender, Microsoft.Reporting.WinForms.SubreportProcessingEventArgs e)
        //{
        //    e.DataSources.Add(new Microsoft.Reporting.WinForms.ReportDataSource("snapshot",this.SnapshotBindingSource));
        //}       
    }
}
