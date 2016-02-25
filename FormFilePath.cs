using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace MIL_TEST_REPORTER
{
    public partial class FormFilePath : Form
    {
        public Form1 ParenMaintForm;
     
        public FormFilePath()
        {
            InitializeComponent(); 


                txtRSPath.Text = System.Configuration.ConfigurationManager.AppSettings["RSpecFilePath"];
                txtASPath.Text = System.Configuration.ConfigurationManager.AppSettings["ASpecFilePath"];
                
                txtBLPath.Text = System.Configuration.ConfigurationManager.AppSettings["BaselineFolderPath"];
               
                txtTReqPhase2.Text = System.Configuration.ConfigurationManager.AppSettings["TestRequirementPhase2"];
                txtTReqPhase3.Text = System.Configuration.ConfigurationManager.AppSettings["TestRequirementPhase3"];
        }

        private void btnDocFile_Click(object sender, EventArgs e)
        {   
            
           
           
        }

        private void btnTextFile_Click(object sender, EventArgs e)
        {
           
        }

        private void FormFilePath_Load(object sender, EventArgs e)
        {

        }

       

        private void button6_Click(object sender, EventArgs e)
        {
            this.Close();
        }

      

       

        private void btnSave_Click(object sender, EventArgs e)
        {
            System.Configuration.ConfigurationManager.AppSettings["RSpecFilePath"] = txtRSPath.Text;
            System.Configuration.ConfigurationManager.AppSettings["ASpecFilePath"] = txtASPath.Text;

            System.Configuration.ConfigurationManager.AppSettings["BaselineFolderPath"] = txtBLPath.Text;

            System.Configuration.ConfigurationManager.AppSettings["TestRequirementPhase2"] = txtTReqPhase2.Text;
            System.Configuration.ConfigurationManager.AppSettings["TestRequirementPhase3"] = txtTReqPhase3.Text;

            Form1 _main = (Form1)this.Tag;
            _main.ASFilePath = txtASPath.Text;
            _main.RSFilePath = txtRSPath.Text;
           
        }

        private void btnBLPath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog _fldr = new FolderBrowserDialog();
            _fldr.ShowDialog();
            txtBLPath.Text = _fldr.SelectedPath;

        }

        private void btnRSPath_Click(object sender, EventArgs e)
        {
            OpenFileDialog _dlg = new OpenFileDialog();
            _dlg.ShowDialog();
            txtRSPath.Text = _dlg.FileName;
        }

        private void btnASPath_Click(object sender, EventArgs e)
        {
            OpenFileDialog _dlg = new OpenFileDialog();
            _dlg.ShowDialog();
            txtASPath.Text = _dlg.FileName;
        }

        private void btnTRPhase2_Click(object sender, EventArgs e)
        {
            OpenFileDialog _dlg = new OpenFileDialog();
            _dlg.ShowDialog();
            txtTReqPhase2.Text = _dlg.FileName;
        }

        private void btnTRPhase3_Click(object sender, EventArgs e)
        {
            OpenFileDialog _dlg = new OpenFileDialog();
            _dlg.ShowDialog();
            txtTReqPhase3.Text = _dlg.FileName;
        }

       
        



    }
}
