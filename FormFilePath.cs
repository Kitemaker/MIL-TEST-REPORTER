using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Configuration;
using System.Reflection;

namespace MIL_TEST_REPORTER
{
    public partial class FormFilePath : Form
    {
        public Form1 ParenMaintForm;
        System.Configuration.Configuration config;

        public FormFilePath()
        {
            InitializeComponent();

            try
            {
                string appPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                string configFile = System.IO.Path.Combine(appPath, "MIL TEST REPORTER.exe.config");
                ExeConfigurationFileMap configFileMap = new ExeConfigurationFileMap();
                configFileMap.ExeConfigFilename = configFile;
                config = ConfigurationManager.OpenMappedExeConfiguration(configFileMap, ConfigurationUserLevel.None);
                
                txtRSPath.Text = config.AppSettings.Settings["RSpecFilePath"].Value;
                txtASPath.Text = config.AppSettings.Settings["ASpecFilePath"].Value;
                txtBLPath.Text = config.AppSettings.Settings["BaselineFolderPath"].Value;
                txtTReqPhase2.Text = config.AppSettings.Settings["TestRequirementPhase2"].Value;
                txtTReqPhase3.Text =   config.AppSettings.Settings["TestRequirementPhase3"].Value;

            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "FormFilePath");
            }
        }

              
        //Close the form
        private void button6_Click(object sender, EventArgs e)
        {
            this.Close();
        } 

       
        //Save the changes made
        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                Form1 _main = (Form1)this.Tag;
                _main.ASFilePath = txtASPath.Text;
                _main.RSFilePath = txtRSPath.Text;               

                config.AppSettings.Settings["RSpecFilePath"].Value = txtRSPath.Text;
                config.AppSettings.Settings["ASpecFilePath"].Value = txtASPath.Text;

                config.AppSettings.Settings["BaselineFolderPath"].Value = txtBLPath.Text;

                config.AppSettings.Settings["TestRequirementPhase2"].Value = txtTReqPhase2.Text;
                config.AppSettings.Settings["TestRequirementPhase3"].Value = txtTReqPhase3.Text;
                config.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message, "Error in method btnSave_Click");                       
            }
           
        }

        private void btnBLPath_Click(object sender, EventArgs e)
        {
            try
            {
                FolderBrowserDialog _fldr = new FolderBrowserDialog();
                _fldr.ShowDialog();
                txtBLPath.Text = _fldr.SelectedPath;
            }
            catch (Exception exc)
            {
                MessageBox.Show(this, exc.Message, "btnBLPath_Click");
            }
        }

        private void btnRSPath_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog _dlg = new OpenFileDialog();
                _dlg.ShowDialog();
                txtRSPath.Text = _dlg.FileName;
            }
            catch (Exception exc)
            {
                MessageBox.Show(this, exc.Message, "btnRSPath_Click");
            }
        }

        private void btnASPath_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog _dlg = new OpenFileDialog();
                _dlg.ShowDialog();
                txtASPath.Text = _dlg.FileName;
            }
            catch (Exception exc)
            {
                MessageBox.Show(this, exc.Message, "btnASPath_Click");
            }
        }

        private void btnTRPhase2_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog _dlg = new OpenFileDialog();
                _dlg.ShowDialog();
                txtTReqPhase2.Text = _dlg.FileName;
            }
            catch (Exception exc)
            {
                MessageBox.Show(this, exc.Message, "btnTRPhase2_Click");
            }
        }

        private void btnTRPhase3_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog _dlg = new OpenFileDialog();
                _dlg.ShowDialog();
                txtTReqPhase3.Text = _dlg.FileName;
            }
            catch (Exception exc)
            {
                MessageBox.Show(this, exc.Message, "btnTRPhase3_Click");
            }
        }

       
        



    }
}
