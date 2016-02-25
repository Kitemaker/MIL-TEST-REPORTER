using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using TestStandAPI = NationalInstruments.TestStand.Interop.API;
using TAdapter = NationalInstruments.TestStand.Interop.AdapterAPI;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Configuration;


namespace MIL_TEST_REPORTER
{
    public partial class Form1 : Form
    {
        //Project Setting
        string ProjectName = string.Empty;

        //File Paths  
        string RSpecFilePath = string.Empty;
        string ASpecFilePath = string.Empty;
        string ASTextFilePath = string.Empty;
        string RSTextFilePath = string.Empty;
        string BaselineFolderPath = string.Empty;
        //string Ph3InstSeqPath = string.Empty;
        string TestRequirementPhase2 = string.Empty;
        string TestRequirementPhase3 = string.Empty;
        string genericseqpath = string.Empty;
        string InstSeqPath = string.Empty;
        string Phase2PrinciplePath = string.Empty;

        public string rsFilePath = "";
        public string rsTextFilePath = "";
        TestStandAPI.Engine Eng;
        Excel.Application ExcelApp;
        Excel.Workbook TestSpec;
        Excel.Worksheet TestReqSheet;
        Excel.Range range;
        string[,] RequirementPair;
        System.IO.DirectoryInfo InstanSeqDir;
        DirectoryInfo selectedFunc;
        int countTotal = 0, countPass = 0, countFailed = 0, CountTerm = 0, CountError = 0;
        double totExecTime = 0;
        System.Collections.Generic.List<FileInfo> testReports;
        int Phase = 0;
       // string genericseqpath =  @"R:\Phase 3\Testing\Test sequences\Generic sequences";
       // string InstSeqPath = @"R:\Phase 3\Testing\Test sequences\Instantiated sequences";
        //string Phase2PrinciplePath = @"R:\Phase 2\Testing\Set_Of_Principles_Test";

        ProgressForm ProgDialog;
        DirectoryInfo GenSeqDirInfo, InstSeqDirInfo;
        string[,] testReqList;
        string[] specText;  // stores text of requiremnt specification in text array
        string[,] chapterIndex;  // stores the chapters of requiremnt specification

        public Form1()
        {
            InitializeComponent();
            this.Height = Screen.PrimaryScreen.Bounds.Height;
            this.Width = Screen.PrimaryScreen.Bounds.Width;

            //Layout Sizing
            ResizeLayout();

            SetConfiguration();

            CreateDriveR();

            // this.btnReq.Enabled = false;
            GenSeqDirInfo = new DirectoryInfo(genericseqpath);
            InstSeqDirInfo = new DirectoryInfo(InstSeqPath);
           
        
           
        }


        private string[] GetReportInfo(string ReportPath)
        {
            String[] ReportInfo = new string[5];
            try
            {

                XmlNodeList nlist;
                XmlDocument xdoc = new XmlDocument();
                using (StreamReader reader = new StreamReader(ReportPath))
                {
                    String line = string.Empty;
                    while ((line = reader.ReadLine()) != null)
                    {
                        if (line.Contains("UUT Result:"))
                        {
                            if (line.Contains("Passed"))
                            {
                                ReportInfo[0] = "Passed";
                                countPass = countPass + 1;
                            }
                            else if (line.Contains("Failed"))
                            {
                                ReportInfo[0] = "Failed";
                                countFailed++;
                            }
                            else if (line.Contains("Terminated"))
                            {
                                ReportInfo[0] = "Terminated";
                                CountTerm++;
                            }
                            else
                            {
                                ReportInfo[0] = "Error";
                                CountError++;
                            }
                        }

                        else if (line.Contains("Date:"))
                        {
                            xdoc.LoadXml(line);
                            nlist = xdoc.GetElementsByTagName("td");
                            ReportInfo[1] = nlist[1].InnerText;

                        }
                        else if (line.Contains(">Time:"))
                        {
                            xdoc.LoadXml(line);
                            nlist = xdoc.GetElementsByTagName("td");
                            ReportInfo[2] = nlist[1].InnerText;
                        }
                        else if (line.Contains("Execution Time:"))
                        {
                            xdoc.LoadXml(line);
                            nlist = xdoc.GetElementsByTagName("td");
                            ReportInfo[3] = nlist[1].InnerText.Substring(0, 5);
                            totExecTime = totExecTime + Convert.ToDouble(nlist[1].InnerText.Substring(0, 5));
                        }
                        else if (line.Contains("Number of Results:"))
                        {
                            xdoc.LoadXml(line);
                            nlist = xdoc.GetElementsByTagName("td");
                            ReportInfo[4] = nlist[1].InnerText;
                        }
                        else if (line.Contains("MainSequence"))
                        { break; }
                    }
                }

                return ReportInfo;
            }
            catch
            {
                MessageBox.Show("Error");
                return ReportInfo;
            }
        }
             

        private string GetExecTime(double ExecTimeInSec)
        {
            string exectime = "0.0";
            try
            {
                TimeSpan t = TimeSpan.FromSeconds(ExecTimeInSec);
                exectime = t.Hours.ToString() + "Hr " + t.Minutes.ToString() + "min  " + Convert.ToString(t.Seconds) + "sec";

            }
            catch (System.Exception exc)
            {
                MessageBox.Show(exc.Message + "GetExecTime");
            }
            return exectime;
        }

        private void btnBrowsePhase2_Click(object sender, EventArgs e)
        {
            try
            {
                FolderBrowserDialog flg = new FolderBrowserDialog();
                flg.ShowNewFolderButton = false;
                flg.ShowDialog();
                this.txtBoxPathPh2.Text = flg.SelectedPath;
                fillTreeView(flg.SelectedPath);
            }
            catch (System.Exception exc)
            {
                MessageBox.Show(exc.Message + "btnBrowsePhase2_Click");
            }
        }


        private void fillTreeView(string Path)
        {
            DirectoryInfo[] childFol;
            DirectoryInfo[] tseq;
            DirectoryInfo dir = new DirectoryInfo(Path);
            childFol = dir.GetDirectories();

            try
            {
                foreach (DirectoryInfo dinfo in childFol)
                {
                    TreeNode tnode = treeView1.TopNode.Nodes.Add(dinfo.Name);
                    tseq = dinfo.GetDirectories("Test Sequences*", SearchOption.AllDirectories);
                    foreach (DirectoryInfo tdir in tseq)
                    {
                        TreeNode TestSeqParent = tnode.Nodes.Add(tdir.Parent.Name);
                        TestSeqParent.Tag = tdir.FullName;
                        TestSeqParent.ToolTipText = tdir.FullName;
                    }
                }
            }
            catch (System.Exception exc)
            {
                MessageBox.Show(exc.Message + "fillTreeView");
            }
        }


        private void UpdateList(string Path)
        {
            string _path = Path;

            FunctionListBox.Items.Clear();
            if (Directory.Exists(_path))
            {
                InstanSeqDir = new System.IO.DirectoryInfo(_path);

                foreach (DirectoryInfo dinfo in InstanSeqDir.GetDirectories())
                {
                    this.FunctionListBox.Items.Add(dinfo.Name);

                }
            }
            else
            {
                MessageBox.Show("Path " + _path + " does not exist.", "Path Not Found");
            }
        }

        private void FunctionListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            int count = 1;
            countPass = 0; countFailed = 0;
            CountTerm = 0; CountError = 0;
            totExecTime = 0;
            countTotal = 0;
            ListViewItem lvitem;
            selectedFunc = new DirectoryInfo(InstanSeqDir.FullName + "\\" + FunctionListBox.SelectedItem);
            testReports = new List<FileInfo>();
            listView1.Items.Clear();
            String[] reportInfoItems = new string[5];
            FileInfo[] seqFiles;
            FileInfo[] reportFiles;

            if (FunctionListBox.SelectedIndex > 0)
            {
                FillContextualMenu(FunctionListBox.SelectedItem.ToString(), FunctionListBox.ContextMenuStrip);
            }

            foreach (FileInfo finfoTest in selectedFunc.GetFiles("*.seq"))
            {
                string[] tesfilename = finfoTest.Name.ToString().Split('.');
                reportFiles = selectedFunc.GetFiles(tesfilename[0] + "*_Report_*.html");

                if (reportFiles.Length > 0)
                {
                    foreach (FileInfo finfo in selectedFunc.GetFiles(tesfilename[0] + "*_Report_*.html"))
                    {
                        testReports.Add(finfo);
                        lvitem = new ListViewItem(count.ToString());
                        lvitem.SubItems.Add(finfoTest.Name);
                        lvitem.SubItems.Add(finfo.Name);
                        lvitem.Tag = finfo.FullName + "$" + finfoTest.FullName;
                        countTotal = countTotal + 1;
                        reportInfoItems = GetReportInfo(finfo.FullName);
                        lvitem.SubItems.Add(reportInfoItems[0]);
                        lvitem.SubItems.Add(reportInfoItems[1]);
                        lvitem.SubItems.Add(reportInfoItems[2]);
                        lvitem.SubItems.Add(reportInfoItems[3]);
                        lvitem.SubItems.Add(reportInfoItems[4]);

                        this.listView1.Items.Add(lvitem);
                        if (reportInfoItems[0] == "Failed")
                        {
                            lvitem.BackColor = System.Drawing.Color.LightPink;
                        }
                        if (reportInfoItems[0] == "Passed")
                        {
                            lvitem.BackColor = System.Drawing.Color.LightGreen;
                        }
                        count = count + 1;
                    }
                }

                // If no report available then simply add test name with Gray Background
                if (reportFiles.Length == 0)
                {
                    lvitem = new ListViewItem(count.ToString());
                    lvitem.SubItems.Add(finfoTest.Name);
                    lvitem.SubItems.Add("Test Not Executed");
                    this.listView1.Items.Add(lvitem);
                    lvitem.BackColor = System.Drawing.Color.LightYellow;
                    count = count + 1;
                }




            }
            seqFiles = selectedFunc.GetFiles("*.seq");
            this.lbl_NoOfTest.Text = Convert.ToString(seqFiles.GetLength(0));
            this.lbl_Executed.Text = countTotal.ToString();
            this.lbl_Passed.Text = countPass.ToString();
            this.lbl_Failed.Text = countFailed.ToString();
            this.lbl_Terminated.Text = CountTerm.ToString();
            this.lbl_Error.Text = CountError.ToString();
            this.lbl_TotalExecTime.Text = GetExecTime(totExecTime);


        }

        //Populate Contextual Menu with Function's Generic Sequence
        private void FillContextualMenu(string GenFunction, ContextMenuStrip ConMenu)
        {
            DirectoryInfo genSeqDir;
            try
            {
                if (Directory.Exists(genericseqpath + "\\" + GenFunction))
                {
                    genSeqDir = new DirectoryInfo(genericseqpath + "\\" + GenFunction);
                    ConMenu.Items.Clear();
                    ConMenu.Tag = GenFunction;
                    foreach (FileInfo finfo in genSeqDir.GetFiles("*.seq"))
                    {
                       ToolStripItem titem = ConMenu.Items.Add(finfo.Name);
                       titem.Tag = finfo.FullName;

                    }
                }
            }
            catch (System.Exception exc)
            {
                MessageBox.Show(exc.Message + "FillContextualMenu");
            }

        }

        private void contextMenuInstant_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            string selectedIXL = "IXL1";
            MLApp.MLApp matlab=null;
            try
            {
                string genericSeq = e.ClickedItem.Tag.ToString();
                string selectedFun = contextMenuInstant.Tag.ToString();
                selectedIXL = cmbBoxIXL.Text;

                string destFolder ="'"+ @"R:\Phase3\Testing\Test sequences\Instantiated sequences\" + selectedIXL + @"\" + selectedFun+"'";

                matlab = new MLApp.MLApp();
                //addpath(genpath('c:/matlab/myfiles'))
                // Change to the directory where the function is located 
                string commandPath = @"addpath(genpath('P:\Test environment'))";
                string genCmd= "readSequence(" + destFolder + "," +"'"+ genericSeq+ "');";
          
               
                matlab.Execute(commandPath);
                // Call the MATLAB function myfunc
                matlab.Execute(genCmd);

                // Display result 
             

             

                  matlab.Quit();
            }

            catch (System.Exception exc)
            {
                MessageBox.Show(exc.Message + "contextMenuInstant_ItemClicked"+ "\n Please contact GE Support");
            }
            finally
            { 
                if (matlab != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(matlab);
                }
            
            }

        }

        
        private void UpdateListViewPhase2(DirectoryInfo TestSequenceFolder)
        {

            int count = 1;
            countPass = 0; countFailed = 0;
            CountTerm = 0; CountError = 0;
            totExecTime = 0;
            countTotal = 0;
            ListViewItem lvitem;
            selectedFunc = TestSequenceFolder;
            testReports = new List<FileInfo>();
            listView2.Items.Clear();
            String[] reportInfoItems = new string[5];
            FileInfo[] seqFiles;
            try
            {
                foreach (FileInfo finfo in selectedFunc.GetFiles("*_Report*.html",SearchOption.AllDirectories))
                {
                    testReports.Add(finfo);
                    lvitem = new ListViewItem(count.ToString());
                    lvitem.SubItems.Add(finfo.Name);
                    lvitem.Tag = finfo.FullName;
                    countTotal = countTotal + 1;
                    reportInfoItems = GetReportInfo(finfo.FullName);
                    lvitem.SubItems.Add(reportInfoItems[0]);
                    lvitem.SubItems.Add(reportInfoItems[1]);
                    lvitem.SubItems.Add(reportInfoItems[2]);
                    lvitem.SubItems.Add(reportInfoItems[3]);
                    lvitem.SubItems.Add(reportInfoItems[4]);

                    this.listView2.Items.Add(lvitem);
                    if (reportInfoItems[0] == "Failed")
                    {
                        lvitem.BackColor = System.Drawing.Color.LightPink;
                    }
                    if (reportInfoItems[0] == "Passed")
                    {
                        lvitem.BackColor = System.Drawing.Color.LightGreen;
                    }
                    count = count + 1;
                }
                seqFiles = selectedFunc.GetFiles("*.seq");
                this.lblNoOfTestPh2.Text = Convert.ToString(seqFiles.GetLength(0));
                this.lblExectedPh2.Text = countTotal.ToString();
                this.lblPassedPh2.Text = countPass.ToString();
                this.lblFailedPh2.Text = countFailed.ToString();
                this.lblTerminatedPh2.Text = CountTerm.ToString();
                this.lblErrorPh2.Text = CountError.ToString();
                this.lblTotalExecTimePh2.Text = GetExecTime(totExecTime);

            }
            catch (System.Exception exc)
            {
                MessageBox.Show(exc.Message + "UpdateListViewPhase2");
            }
        }

   

        private string GetTestReqFromSeqFile(string SeqFIlePath)
        {
            Eng = new TestStandAPI.Engine();
            int stepCount, seqCount;
            string req = string.Empty;
            TestStandAPI.SequenceFile testfile;
            testfile = (TestStandAPI.SequenceFile)Eng.GetSequenceFileEx(SeqFIlePath);
            seqCount = testfile.NumSequences;//(TestStandAPI.StepGroups.StepGroup_Main);

            ProgDialog.lblStepDetail.Text = SeqFIlePath;
            try
            {
                for (int i = 0; i < seqCount; i++)
                {
                    TestStandAPI.Sequence testseq = testfile.GetSequence(i); // Steps in MainSequence
                    stepCount = testseq.GetNumSteps(TestStandAPI.StepGroups.StepGroup_Main);
                    for (int k = 0; k < stepCount; k++)
                    {
                        TestStandAPI.Step stp = testseq.GetStep(k, TestStandAPI.StepGroups.StepGroup_Main);

                        TestStandAPI.PropertyObject links = stp.Requirements.GetPropertyObject("Links", 0);
                        int nextAvailableIndex = links.GetNumElements();
                        //MessageBox.Show(  stp.GetDescriptionEx());
                        if (nextAvailableIndex > 0)
                        {
                            int t;
                            for (t = 0; t < nextAvailableIndex; t++)
                            {
                                if (req == string.Empty)
                                {
                                    req = links.GetValStringByOffset(t, 0);
                                }
                                else
                                { req = req + ";" + links.GetValStringByOffset(t, 0); }
                            }
                        }

                    }
                }
            }
            catch (System.Exception exc)
            {
                MessageBox.Show(exc.Message + "GetTestReqFromSeqFile");
            }

            return req;
        }

        private void GetTestRequirementFromExcel()
        {
            SetConfiguration();
            this.progressBarIXL.Minimum = 0;
            Excel.Range range1, range2, range3, range4, rangeNA_TR;
            ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Visible = false;


            if (Phase == 2)
            {
                TestSpec = ExcelApp.Workbooks.Open(TestRequirementPhase2);
                TestReqSheet = (Excel.Worksheet)TestSpec.Worksheets["Test_Requirements"];
                
            }
            else
            {
                TestSpec = ExcelApp.Workbooks.Open(TestRequirementPhase3);
                TestReqSheet = (Excel.Worksheet)TestSpec.Worksheets["Test_Requirements"];
                
            }


            ProgDialog = new ProgressForm();
            ProgDialog.Show(this);
            ProgDialog.lblCurrentStep.Text = "Reading Test Requirement from " + TestSpec.FullName;
            // range = (Excel.Range)TestReqSheet.get_Range("B:B",Type.Missing);
           // range = (Excel.Range)TestReqSheet.Rows[TestReqSheet.UsedRange.Rows.Count+3, Type.Missing];
            range = (Excel.Range)TestReqSheet.Rows;
            RequirementPair = new string[TestReqSheet.UsedRange.Rows.Count+3, 3];
            RequirementPair[0, 0] = ""; RequirementPair[0, 1] = ""; RequirementPair[0, 2] = "";
            this.progressBarIXL.Maximum = range.Rows.Count;
            string TestReq, NonApplicableTR;
            for (int n = 3; n <= TestReqSheet.UsedRange.Rows.Count ; n++)
            {
                try
                {
                    progressBarIXL.Value = n + 1;
                    range1 = range.Rows.get_Item(n, Type.Missing);
                    range2 = range1.Columns.get_Item(2, Type.Missing);
                    range3 = range1.Columns.get_Item(3, Type.Missing);
                    range4 = range1.Columns.get_Item(4, Type.Missing);

                    if (ProjectName == "LTA")
                    { rangeNA_TR = range1.Columns.get_Item(9, Type.Missing); }
                    else
                    { rangeNA_TR = range1.Columns.get_Item(14, Type.Missing); }
                    
                    string IXLReq = Convert.ToString(range2.Value);
                    NonApplicableTR = Convert.ToString(rangeNA_TR.Value);
                    if (Phase == 2)
                    {
                        TestReq = Convert.ToString(range4.Value);
                    }
                    else
                    {
                        if (ProjectName == "LTA")
                        {
                            TestReq = Convert.ToString(range4.Value);
                        }
                        else
                        {
                            TestReq = Convert.ToString(range3.Value);
                        }
                        
                    }
                    string add = range2.Cells.Address;
                    Excel.Range rr = TestReqSheet.Cells[n, "B"];
                    add = Convert.ToString(rr.Value);
                    string sr = rr.Address + rr.AddressLocal;
                    string ReqName = Convert.ToString(range4.Value);
                    if (range2.Value != null)
                    {
                        RequirementPair[n - 3, 0] = IXLReq;
                        ProgDialog.lblStepDetail.Text = IXLReq;
                        RequirementPair[n - 3, 1] = TestReq + "$" + NonApplicableTR;
                        RequirementPair[n - 3, 2] = ReqName;
                    }
                    else
                    {
                        RequirementPair[n - 3, 0] = RequirementPair[n - 4, 0];
                        ProgDialog.lblStepDetail.Text = RequirementPair[n - 4, 0];
                        RequirementPair[n - 3, 1] = TestReq + "$" + NonApplicableTR;
                        RequirementPair[n - 3, 2] = ReqName;
                    }
                }
                catch (System.Exception e)
                {
                    MessageBox.Show(e.Message + "GetTestRequirementFromExcel");
                    this.Cursor = System.Windows.Forms.Cursors.Arrow;
                }

            }

            TestSpec.Close();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ExcelApp);
            ExcelApp = null;
        }

        private void btnTraceP3_Click(object sender, EventArgs e)
        {
            int IXL_Req_Count = 0;
            int TR_Req_Count = 0;
            int TR_Covered = 0;
            this.listViewTrace.Items.Clear();
            lblPrg1.Visible = true;
            lblMatchReq.Visible = true;
            progressBarIXL.Visible = true;
            progressBarTrace.Visible = true;


            progressBarTrace.Minimum = 0;
            progressBarTrace.Value = 0;
            progressBarIXL.Minimum = 0; progressBarIXL.Value = 0;
            Phase = 3;
            try
            {
                FileInfo[] testSeqInfos = GenSeqDirInfo.GetFiles("*.seq", SearchOption.AllDirectories);
                GetTestRequirementFromExcel();

                testReqList = new string[testSeqInfos.Length, 3];
                int i = 0;
                ProgDialog.lblCurrentStep.Text = "Reading Requirement From Test Sequences";
                foreach (FileInfo seq in testSeqInfos)
                {
                    testReqList[i, 0] = seq.Name;
                    testReqList[i, 1] = seq.FullName;
                    testReqList[i, 2] = GetTestReqFromSeqFile(seq.FullName);
                    i++;
                }
                ListViewItem lvit;

                progressBarTrace.Maximum = RequirementPair.GetLength(0);
                for (int k = 1; k < RequirementPair.GetLength(0); k++)
                {


                    progressBarTrace.Value = k + 1;

                    if (RequirementPair[k, 0] != null & RequirementPair[k, 0] != string.Empty & RequirementPair[k, 1] != null)
                    {
                        string TestReq="", NonApplicableTR="";
                        // SPlit the test requirment and Non Applicable TR Reason
                        if (RequirementPair[k, 1].Split('$').Length > 1)
                        {
                            TestReq = RequirementPair[k, 1].Split('$')[0];
                            NonApplicableTR = RequirementPair[k, 1].Split('$')[1];
                        }
                        
                        string testSequenceName = GetTraceability(TestReq.Trim());

                        {
                            
                            lvit = new ListViewItem();
                            lvit.SubItems.Add(RequirementPair[k, 0]);
                            lvit.SubItems.Add(TestReq);
                            lvit.SubItems.Add(testSequenceName);
                            lvit.SubItems.Add(NonApplicableTR);
                            if (testSequenceName != string.Empty)
                            {
                                TR_Covered = TR_Covered + 1;
                            }
                            //  lvit.SubItems.Add("");
                            this.listViewTrace.Items.Add(lvit);
                            IXL_Req_Count = IXL_Req_Count + 1;
                            TR_Req_Count = TR_Req_Count + 1;
                        }
                    }
                    else if (RequirementPair[k, 0] != null & RequirementPair[k, 0] != string.Empty & RequirementPair[k, 1] == null)
                    {
                        lvit = new ListViewItem();
                        lvit.SubItems.Add(RequirementPair[k, 0]);
                        lvit.SubItems.Add("");
                        lvit.SubItems.Add("");
                        if (k > 0 && (RequirementPair[k, 0] != RequirementPair[k - 1, 0]))
                        {
                            this.listViewTrace.Items.Add(lvit);
                            IXL_Req_Count = IXL_Req_Count + 1;
                        }
                    }

                }
                GetChapterName();

            }

            catch (System.Exception exc)
            {
                MessageBox.Show(exc.Message + exc.Data + "\nbtnTraceP3_Click");
            }

            //Enable buttons
            btnExport.Enabled = true;


            lbl_Count_IXL_Req.Text = GetUniqueCount().ToString();
            lbl_Count_TR_Req.Text = TR_Req_Count.ToString();
            lbl_TR_Covered.Text = TR_Covered.ToString();

            //Hide Progress Bars
            lblPrg1.Visible = false;
            lblMatchReq.Visible = false;
            progressBarIXL.Visible = false;
            progressBarTrace.Visible = false;
            this.Cursor = System.Windows.Forms.Cursors.Arrow;

            //Close Progress Dialog
            if (ProgDialog != null)
            {
                ProgDialog.Close();
                ProgDialog.Dispose();

            }

        }


        private string GetTraceability(string IXLReq)
        {
            string traceString = string.Empty;
            try
            {
                for (int k = 0; k < testReqList.GetLength(0); k++)
                {
                    if (testReqList[k, 2].ToString().Contains(IXLReq))
                    {
                        if (traceString == string.Empty)
                        { traceString = testReqList[k, 0]; }
                        else
                        {
                            traceString = traceString + " , " + testReqList[k, 0];
                        }
                    }

                }
            }
            catch (System.Exception exc)
            {
                MessageBox.Show(exc.Message + "\nGetTraceability");
            }
            return traceString;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            Excel.Workbook TestReport=null;
            Excel.Application ExcelReportApp=null;
            Excel.Worksheet TestReqSheet=null;
            Excel.Worksheet CompactReqSheet = null;
            Excel.Range exRange = null;
            System.Windows.Forms.SaveFileDialog dlg;
           // System.IO.StreamWriter wr;
            try
            {          
                dlg = new SaveFileDialog();
                dlg.CreatePrompt = true;

                ExcelReportApp = new Microsoft.Office.Interop.Excel.Application();
                ExcelReportApp.Visible = false;
                TestReport = ExcelReportApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
                TestReqSheet = (Excel.Worksheet)TestReport.Worksheets.Add();
                CompactReqSheet = (Excel.Worksheet)TestReport.Worksheets.Add();
               
                TestReqSheet.Name = "Full";
                CompactReqSheet.Name = "Compact";
                dlg.Filter = "Excel files(*.xlsx)|*.xlsx";
                dlg.DefaultExt = ".xlsx";
                dlg.ShowDialog();
              
                if (listViewTrace.Items.Count > 0)
                {
                    
                    exRange = TestReqSheet.get_Range("A1","D1");
                    exRange.Font.Size = 12;
                    exRange.Font.Color = System.Drawing.Color.DarkBlue;                   

                    TestReqSheet.Cells[1, 1] = "Req Chapter Name";                  
                    TestReqSheet.Cells[1, 2] = "IXL Requirement";
                    TestReqSheet.Cells[1, 3] = "Test Requirement";
                    TestReqSheet.Cells[1, 4] = "Test Sequence Name";


                    CompactReqSheet.Cells[1, 1] = "IXL Requirement";
                    CompactReqSheet.Cells[1, 2] = "Test Sequence Name";
                    CompactReqSheet.get_Range("A1", "D1").Font.Color = System.Drawing.Color.DarkBlue;
                    for (int k = 0; k < listViewTrace.Items.Count ; k++)                  
                    {
                   
                       // Excel Row and Column BaseNumber starts with "1" therefore K+2 and n=1 is used
                        for (int n = 0; n <= 3; n++)
                        {
                            if (n < 3)
                            {
                                TestReqSheet.Cells[k + 2, n + 1] = listViewTrace.Items[k].SubItems[n].Text;
                            }
                            // Merge TR File Name and TR Non Applicability
                            if (n == 3)
                            {
                                TestReqSheet.Cells[k + 2, n + 1] = listViewTrace.Items[k].SubItems[n].Text + listViewTrace.Items[k].SubItems[n+1].Text;  
                            }
                        }


                    }

                  
                    // Add Compact Sheet 
                    string ixlReq = listViewTrace.Items[0].SubItems[1].Text;
                    string testSeq = listViewTrace.Items[0].SubItems[3].Text;         
                   
                    int rowCount = 2;
                    int ixlReqCount = 0;
                    for (int k = 1; k <listViewTrace.Items.Count; k++)
                    {
                        
                        if (ixlReq == listViewTrace.Items[k].SubItems[1].Text)
                        {
                            if (listViewTrace.Items[k].SubItems[3].Text != string.Empty)
                            {
                                if (testSeq != string.Empty)
                                {
                                    
                                        testSeq = testSeq + "," + listViewTrace.Items[k].SubItems[3].Text;
                                       
                                }
                                else
                                {
                                    testSeq = listViewTrace.Items[k].SubItems[3].Text;
                                }
                            }
                        }
                        else
                        {
                            CompactReqSheet.Cells[rowCount, 1] = ixlReq;
                            string _testSeqDupRemoved = RemoveDuplitaces(testSeq, ',');

                            //if IXL Requirement is not covered in test sequence then put the reason "NonApplicable TR" in that column
                            if (_testSeqDupRemoved == string.Empty)
                            {
                                //exclude title row r=1
                                for (int r = 2; r < listViewTrace.Items.Count; r++)
                                {
                                    if (ixlReq == listViewTrace.Items[r].SubItems[1].Text)
                                    {
                                        CompactReqSheet.Cells[rowCount, 2] = listViewTrace.Items[r].SubItems[4].Text;
                                        break;
                                    }
                                }
                            }
                            else 
                            {
                                CompactReqSheet.Cells[rowCount, 2] = _testSeqDupRemoved;
                            }
                           
                            ixlReqCount++;
                            ixlReq = listViewTrace.Items[k].SubItems[1].Text;
                            testSeq = listViewTrace.Items[k].SubItems[3].Text;
                            rowCount = rowCount + 1;
                        }

                    }

                }
                else
                {
                    MessageBox.Show(" No Traceablity Item, Please get the traceablity matrix first");

                }
                TestReport.SaveAs(dlg.FileName);
                TestReport.Close();
                MessageBox.Show("Report Exported Successfully at :" + dlg.FileName);
               
                             
            }
            catch (System.Exception exc)
            {
                MessageBox.Show("Report is not exported");
                MessageBox.Show(exc.Message + "\nbtnExport_Click");
            }
            finally
            {               
                ExcelReportApp.Quit();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ExcelReportApp);
                ExcelReportApp = null;
            }

        }

        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {

            try
            {
                if (this.tabControl2.SelectedTab.Name == "tabPage3")
                {
                    lblPrg1.Visible = false;
                    lblMatchReq.Visible = false;
                    progressBarIXL.Visible = false;
                    progressBarTrace.Visible = false;

                }
            }
            catch (System.Exception exc)
            {
                MessageBox.Show(exc.Message + "tabControl2_SelectedIndexChanged");
            }
        }

        private void openResultToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                ListViewItem lvitem=null;
                if (tabControl2.SelectedTab == tabPage2)
                {
                   lvitem = listView2.SelectedItems[0];
                }
                else
                { 
                    lvitem = listView1.SelectedItems[0];
                }

                string[] fnames = Convert.ToString(lvitem.Tag).Split('$');
                if (fnames[0] != string.Empty)
                {
                    System.Diagnostics.Process.Start("iexplore", fnames[0]);
                }
            }
            catch (System.Exception exc)
            { MessageBox.Show("openResultToolStripMenuItem_Click \n" + exc.Message); }
        }

        private void openSequenceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                ListViewItem lvitem = listView1.SelectedItems[0];
                string[] fnames = Convert.ToString(lvitem.Tag).Split('$');
                MessageBox.Show(fnames[1]);

            }
            catch (System.Exception exc)
            { MessageBox.Show("openSequenceToolStripMenuItem_Click \n" + exc.Message); }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.SetAutoSizeMode(System.Windows.Forms.AutoSizeMode.GrowAndShrink);
            txtBoxPathPh2.Text = Phase2PrinciplePath;

        }

        private int GetUniqueCount()
        {
            int count = listViewTrace.Items.Count;
            try
            {
                for (int k = 0; k < listViewTrace.Items.Count; k++)
                {
                    string refItem = listViewTrace.Items[k].Text;
                    for (int n = k + 1; n < listViewTrace.Items.Count; n++)
                    {
                        if (refItem == listViewTrace.Items[n].Text)
                        {
                            count = count - 1;
                            break;

                        }
                    }
                }
            }
            catch (System.Exception exc)
            {
                MessageBox.Show(exc.Message + "GetUniqueCount");
            }

            return count;
        }

        private void btnTraceP2_Click(object sender, EventArgs e)
        {

            int IXL_Req_Count = 0;
            int TR_Req_Count = 0;
            int TR_Covered = 0;
            this.listViewTrace.Items.Clear();
            lblPrg1.Visible = true;
            lblMatchReq.Visible = true;
            progressBarIXL.Visible = true;
            progressBarTrace.Visible = true;
            DirectoryInfo SetOfPrinciplesTest;

            progressBarTrace.Minimum = 0;
            progressBarTrace.Value = 0;
            progressBarIXL.Minimum = 0; progressBarIXL.Value = 0;
            Phase = 2;
            try
            {

                if (!(System.IO.Directory.Exists(Phase2PrinciplePath)))
                {
                    MessageBox.Show(" The folder:  " + Phase2PrinciplePath + "  does not exist. Please Map the folder first and try again");
                    return;
                }
                SetOfPrinciplesTest = new DirectoryInfo(Phase2PrinciplePath);


                FileInfo[] testSeqInfos = SetOfPrinciplesTest.GetFiles("*.seq", SearchOption.AllDirectories);

                //Read Requirments from Excel File
                GetTestRequirementFromExcel();
                ProgDialog.lblCurrentStep.Text = " Reading Test Sequences";
                testReqList = new string[testSeqInfos.Length, 3];
                int i = 0;
                foreach (FileInfo seq in testSeqInfos)
                {
                    testReqList[i, 0] = seq.Name;
                    testReqList[i, 1] = seq.FullName;
                    testReqList[i, 2] = GetTestReqFromSeqFile(seq.FullName);
                    i++;
                    ProgDialog.lblStepDetail.Text = seq.FullName;
                }
                ListViewItem lvit;

                progressBarTrace.Maximum = RequirementPair.GetLength(0);
                for (int k = 1; k < RequirementPair.GetLength(0); k++)
                {

                    ProgDialog.lblCurrentStep.Text = "Matching Requirements";
                    progressBarTrace.Value = k + 1;

                    if (RequirementPair[k, 0] != null & RequirementPair[k, 0] != string.Empty & RequirementPair[k, 1] != null)
                    {
                         string TestReq = "", NonApplicableTR = "";

                        // SPlit the test requirment and Non Applicable TR Reason
                        if (RequirementPair[k, 1].Split('$').Length > 1)
                        {
                            TestReq = RequirementPair[k, 1].Split('$')[0];
                            NonApplicableTR = RequirementPair[k, 1].Split('$')[1];
                        }

                        string tr = GetTraceability(TestReq);

                        {
                            lvit = new ListViewItem();
                            lvit.SubItems.Add(RequirementPair[k, 0]);
                            ProgDialog.lblStepDetail.Text = RequirementPair[k, 0];
                            lvit.SubItems.Add(TestReq);
                            lvit.SubItems.Add(tr);
                            lvit.SubItems.Add(NonApplicableTR);
                            if (tr != string.Empty)
                            {
                                TR_Covered = TR_Covered + 1;
                            }
                            this.listViewTrace.Items.Add(lvit);
                            IXL_Req_Count = IXL_Req_Count + 1;
                            TR_Req_Count = TR_Req_Count + 1;
                        }
                    }
                    else if (RequirementPair[k, 0] != null & RequirementPair[k, 0] != string.Empty & RequirementPair[k, 1] == null)
                    {
                        lvit = new ListViewItem();
                        lvit.SubItems.Add(RequirementPair[k, 0]);
                        ProgDialog.lblStepDetail.Text = RequirementPair[k, 0];
                        lvit.SubItems.Add("");
                        lvit.SubItems.Add("");
                        if (k > 0 && (RequirementPair[k, 0] != RequirementPair[k - 1, 0]))
                        {
                            this.listViewTrace.Items.Add(lvit);
                            IXL_Req_Count = IXL_Req_Count + 1;
                        }
                    }

                }
                // Fill Chapter Names
                GetChapterName();

            }

            catch (System.Exception exc)
            {
                MessageBox.Show(exc.Message + exc.Data + "\nbtnTraceP3_Click");
            }
            btnExport.Enabled = true;

            lbl_Count_IXL_Req.Text = GetUniqueCount().ToString();
            lbl_Count_TR_Req.Text = TR_Req_Count.ToString();
            lbl_TR_Covered.Text = TR_Covered.ToString();
            lblPrg1.Visible = false;
            lblMatchReq.Visible = false;
            progressBarIXL.Visible = false;
            progressBarTrace.Visible = false;
            this.Cursor = System.Windows.Forms.Cursors.Arrow;
            if (ProgDialog != null)
            {
                ProgDialog.Close();
                ProgDialog.Dispose();

            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            ResizeLayout();
        }

        private void ResizeLayout()
        {
            try
            {
                this.tabControl2.Height = this.Height - 100;
                this.tabControl2.Width = this.Width - 10;
                listViewTrace.Height = tabControl2.Height - 100;
                listViewTrace.Height = tabControl2.Height - 100;
                listView1.Width = tabControl2.Width - listView1.Location.X - 10;


                listViewTrace.Width = tabControl2.Width - listViewTrace.Location.X - 10;
                listView2.Width = tabControl2.Width - listView2.Location.X - 10;

                //Column Resizing

            }
            catch (SystemException exc)
            {
                MessageBox.Show(exc.Message + exc.Data + "ResizeLayout");
            }

        }

        private void treeView1_AfterSelect_1(object sender, TreeViewEventArgs e)
        {
            string testSeqPath = Convert.ToString(treeView1.SelectedNode.Tag);
            if (testSeqPath != string.Empty)
            {
                UpdateListViewPhase2(new DirectoryInfo(testSeqPath));
            }
        }
        
        private string GetChNoOfReq(string[,] ChapterList, string[] DocText, string RequirementText)
        {
            string chName = "";
            for (int n = 0; n < ChapterList.GetLength(0); n++)
            {
                if (n < ChapterList.GetLength(0) - 2)
                {
                    int startInd = Convert.ToInt32(ChapterList[n, 1]);
                    int EndInd = Convert.ToInt32(ChapterList[n + 1, 1]);

                    for (int i = startInd; i <= EndInd; i++)
                    {
                        if (DocText[i].ToString().Contains(RequirementText))
                        {

                            return ChapterList[n, 0];

                        }

                    }
                }
            }

            return chName;

        }

        private void GetChapterName()
        {

            Word.Application wordApp;
            Word.Document wDoc = null;
            wordApp = new Word.Application();
            SetConfiguration();
            string _specFilePath = "";
            string _textFilePath = "";
            if (Phase == 2)
            {
                _specFilePath = ASpecFilePath;
                // _textFilePath = ASTextFilePath;
            }
            else
            {
                _specFilePath = RSpecFilePath;
                //  _textFilePath = RSTextFilePath;
            }
            ProgDialog.lblCurrentStep.Text = "Reading Chapter Names From " + _specFilePath;

            try
            {
                object referenceType = Word.WdReferenceType.wdRefTypeNumberedItem;
                if (_specFilePath != string.Empty)
                {

                    wordApp.Visible = false;
                    wDoc = wordApp.Documents.Open(_specFilePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    wDoc.Activate();
                    Array chapterList = (Array)(object)wDoc.GetCrossReferenceItems(ref referenceType);

                    // Get the content of spec in text form
                    specText = wDoc.Content.Text.Split('\r');

                    wDoc.SaveAs2(System.IO.Directory.GetParent(Application.ExecutablePath).FullName + "\\" + wDoc.Name + ".txt", Word.WdSaveFormat.wdFormatText);
                    _textFilePath = wDoc.FullName;
                    wDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wDoc);
                    wDoc = null;
                    wordApp.Quit(Word.WdSaveOptions.wdDoNotSaveChanges, Type.Missing, Type.Missing);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                    wordApp = null;
                    specText = System.IO.File.ReadAllLines(_textFilePath);
                    chapterIndex = new string[chapterList.Length, 2];
                    int txlineHit = 0;

                    // Get the line number for each chapter
                    for (int k = 1; k <= chapterList.Length; k++)
                    {
                        for (int txtLine = txlineHit; txtLine < specText.Length; txtLine++)
                        {
                            string normalized1 = System.Text.RegularExpressions.Regex.Replace(specText[txtLine], @"\s", "");
                            string normalized2 = System.Text.RegularExpressions.Regex.Replace(chapterList.GetValue(k).ToString(), @"\s", "");
                            if (string.Equals(normalized1, normalized2))
                            {
                                // chapterIndex is Zero Based Array
                                chapterIndex[k - 1, 0] = chapterList.GetValue(k).ToString();
                                chapterIndex[k - 1, 1] = txtLine.ToString();

                            }
                        }

                    }
                    //Fill List View
                    foreach (ListViewItem lvm in listViewTrace.Items)
                    {
                        if (lvm.SubItems.Count > 0)
                        {
                            ProgDialog.lblStepDetail.Text = lvm.SubItems[1].Text;
                            lvm.Text = (GetChNoOfReq(chapterIndex, specText, lvm.SubItems[1].Text));
                        }
                    }



                }
            }
            catch (SystemException exc)
            {
                MessageBox.Show(exc.Message + "\n" + exc.Data + "\n" + "GetChapterName");
            }
            finally
            {
                // Release all Interop objects.
                if (wDoc != null)
                {
                    wDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wDoc);
                }
                if (wordApp != null)
                {
                    wordApp.Quit(Word.WdSaveOptions.wdDoNotSaveChanges, Type.Missing, Type.Missing);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                }

                GC.Collect();
            }
        }

        private void SetConfiguration()
        {
            try
            {
                ProjectName = System.Configuration.ConfigurationManager.AppSettings["Project"];
                RSpecFilePath = System.Configuration.ConfigurationManager.AppSettings["RSpecFilePath"];
                ASpecFilePath = System.Configuration.ConfigurationManager.AppSettings["ASpecFilePath"];
                ASTextFilePath = System.Configuration.ConfigurationManager.AppSettings["ASTextFilePath"];
                RSTextFilePath = System.Configuration.ConfigurationManager.AppSettings["RSTextFilePath"];
                BaselineFolderPath = System.Configuration.ConfigurationManager.AppSettings["BaselineFolderPath"];
                
                TestRequirementPhase2 = System.Configuration.ConfigurationManager.AppSettings["TestRequirementPhase2"];
                TestRequirementPhase3 = System.Configuration.ConfigurationManager.AppSettings["TestRequirementPhase3"];

                Phase2PrinciplePath = System.Configuration.ConfigurationManager.AppSettings["Ph2SetofPrinciple"];
                genericseqpath = System.Configuration.ConfigurationManager.AppSettings["GenericSequencesPath"];
                InstSeqPath = System.Configuration.ConfigurationManager.AppSettings["InstantiatedSeqPath"];
                

            }
            catch (SystemException exc)
            {
                MessageBox.Show(exc.Message + exc.Data + "SetConfiguration");
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            FormFilePath fr = new FormFilePath();
            fr.Show();

        }

        private void btnPrincipleList2_Click(object sender, EventArgs e)
        {
            try
            {
                if (System.IO.Directory.Exists(Phase2PrinciplePath))
                {
                    this.treeView1.HideSelection = false;
                    fillTreeView(Phase2PrinciplePath);

                    treeView1.TopNode.Expand();


                }
                else
                {
                    MessageBox.Show("Directory " + Phase2PrinciplePath + "  Not found. Please Select it");
                }
            }
            catch (SystemException exc)
            {
                MessageBox.Show(exc.Message + exc.Data + "btnPrincipleList2_Click");
            }

        }
        
        private void CreateDriveR()
        {
            try
            {
                SetConfiguration();
                ExecuteCommand("subst /D R:");
                ExecuteCommand("subst R: " + '"' + BaselineFolderPath + '"');
            }

            catch (SystemException exc)
            {
                MessageBox.Show(exc.Message + exc.Data + "CreateDriveR");
            }

        }


        private void ExecuteCommand(string command)
        {
            try
            {
                System.Diagnostics.Process prc = new System.Diagnostics.Process();
                prc.StartInfo.FileName = "cmd";
                prc.StartInfo.Arguments = "/C " + command;
                prc.StartInfo.CreateNoWindow = false;
                prc.StartInfo.UseShellExecute = false;
                prc.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                prc.Start();
                prc.WaitForExit(1000);
            }

            catch (SystemException exc)
            {
                MessageBox.Show(exc.Message + exc.Data + "ExecuteCommand");
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                ExecuteCommand("subst /D R:");
            }
            catch (SystemException exc)
            {
                MessageBox.Show(exc.Message + exc.Data + "Form1_FormClosing");
            }

        }

        private void btnSettings_Click(object sender, EventArgs e)
        {
            FormFilePath frmPath = new FormFilePath();
            frmPath.Tag = this;
            frmPath.ShowDialog();

        }

        private void btnChangeParam_Click(object sender, EventArgs e)
        {
            TestStandAPI.Engine TSEng;
            TSEng = new TestStandAPI.Engine();
            string IXLSelected = cmbBoxIXL.Text;

            TSEng.Globals.SetValString("veristand_model_name", 0, "veristand_" + IXLSelected);
            TSEng.Globals.SetValString("veristand_model_path", 0, "Targets/Controller/Simulation Models/Models/veristand_" + IXLSelected);
            TSEng.Globals.SetValString("controlTable", 0, @"S:\TAG_IXLTestLine_BL3_1_1\TestLine_"+ IXLSelected + @"\ControlTables\CT_TestLine_" + IXLSelected + "_modified.xls");
            TSEng.Globals.SetValString("IXL_Name", 0, IXLSelected);              
           
            
        }

       

        private void tabControl2_Selected(object sender, TabControlEventArgs e)
        {
            try
            {
                DirectoryInfo InstanceFolder = new DirectoryInfo(InstSeqPath);
                cmbBoxIXL.Items.Clear();
                if (ProjectName == "Generic")
                {
                    foreach (DirectoryInfo dinfo in InstanceFolder.GetDirectories("IXL*"))
                    {
                        cmbBoxIXL.Items.Add(dinfo.Name);

                    }
                }

                else if (ProjectName == "LTA")
                {
                    cmbBoxIXL.Items.Add("IXL1");
                
                }
            }
            catch (SystemException exc)
            {
                MessageBox.Show(exc.Message + exc.Data + "tabControl2_Selected");
            }
        }

        private void cmbBoxIXL_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string _path;
                FunctionListBox.Items.Clear();
                if (ProjectName == "Generic")
                {
                    _path = InstSeqPath + "\\" + cmbBoxIXL.Text;
                }
                else
                {
                    _path = InstSeqPath;
                }

                UpdateList(_path);
                textBox2.Text = _path;
            }
            catch (SystemException exc)
            {
                MessageBox.Show(exc.Message + exc.Data + "cmbBoxIXL_SelectedIndexChanged");
            }
        }

        private string RemoveDuplitaces(string InputString, char delimiter)
        {
           
            string refTxt=string.Empty;
            try
            {
                string[] splitTxt = InputString.Split(delimiter);
                if (splitTxt.Length > 0)
                {

                  //  refTxt = splitTxt[0];
                    for (int i = 0; i < splitTxt.Length; i++)
                    {
                        string temp = splitTxt[i];
                        for (int k = i+1; k < splitTxt.Length; k++)
                        {
                            if (temp == splitTxt[k])
                            {
                                splitTxt[k] = string.Empty;
                                
                            }
                        }                     
                    }
                    for (int i = 0; i < splitTxt.Length; i++)
                    {
                        if (splitTxt[i] != string.Empty)
                        {
                            if (refTxt == string.Empty)
                            {
                                refTxt =  splitTxt[i];
                            }
                            else
                            {
                                refTxt = refTxt + "," + splitTxt[i];
                            }
                        }
                    }

                }

            }
            catch (SystemException exc)
            {
                MessageBox.Show(exc.Message + exc.Data + "RemoveDuplitaces");
                
            }

            
            return refTxt;
        }
        public string ASFilePath
        {
            get { return this.ASpecFilePath; }
            set {this.ASpecFilePath = value; }
        
        }

        public string RSFilePath
        {
            get { return this.RSpecFilePath; }
            set { this.RSpecFilePath = value; }

        }
        public string TestRequirmentPhase2Path
        {
            get { return this.TestRequirementPhase2; }
            set { this.TestRequirementPhase2 = value; }

        }
        public string TestRequirmentPhase3Path
        {
            get { return this.TestRequirementPhase3; }
            set { this.TestRequirementPhase3 = value; }

        }
      

        public string GenricSequencePath
        {
            get { return this.genericseqpath; }
            set { this.genericseqpath = value; }

        }
        public string Phase2_Principle_Path
        {
            get { return this.Phase2PrinciplePath; }
            set { this.Phase2PrinciplePath = value; }

        }

        public string InstantiatedSequencePath
        {
            get { return this.InstSeqPath; }
            set { this.InstSeqPath = value; }

        }



    }


}

