using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using Microsoft.Reporting.WinForms;

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
            string asmpath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            DirectoryInfo dir = new DirectoryInfo(asmpath + @"\temp");
            FileInfo[] files = dir.GetFiles("*.xml");
            foreach (FileInfo file in files)
            {
                comboBox1.Items.Add(file);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string asmpath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string exeFolder = Path.GetDirectoryName(Application.ExecutablePath);
            DataSet ds = new DataSet();
            ds.ReadXml(asmpath + "\\temp\\" + comboBox1.SelectedItem);
            //ds.ReadXml(asmpath + @"\temp\Drawings.xml");
            string schemaName = Path.GetFileNameWithoutExtension(comboBox1.SelectedItem.ToString()) + ".xsd";
            ds.WriteXmlSchema(asmpath + "\\temp\\" + schemaName);
            //ds.WriteXmlSchema(asmpath + @"\temp\Drawings.xsd");
            //ds = null;
            //DataSet snds = new DataSet();
            //snds.ReadXml(asmpath + @"\temp\Snapshots.xml");
            //snds.WriteXmlSchema(asmpath + @"\temp\Snapshots.xsd");
            DrawingBindingSource.DataSource = ds;
            this.reportViewer1.LocalReport.EnableExternalImages = true;
            ReportParameter AppPath = new ReportParameter("AppPath", exeFolder);
            this.reportViewer1.LocalReport.SetParameters(AppPath);
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
