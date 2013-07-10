using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Diagnostics;
using System.Reflection;

namespace SemtechAssistant.Resources
{
    public partial class FormTripReport : Form
    {
        public FormTripReport()
        {
            InitializeComponent();

            InitializeBackgoundWorker();
        }

        private SemtechAssistant mainForm = null;
        public FormTripReport(Form callingForm)
        {
            mainForm = callingForm as SemtechAssistant; 

            InitializeComponent();

            InitializeBackgoundWorker();
        }

        private void buttonCreate_Click(object sender, EventArgs e)
        {
            // Set animated process indicator
            //pictureBox1.Visible = true;
            this.pictureBox1.Image = Properties.Resources.Animation;
            this.pictureBox1.Refresh();

            backgroundWorker1.RunWorkerAsync();
        }

        private void findAndReplace(Word.Application wApp, object searchObj, object replaceObj)
        {
            if (wApp == null || searchObj == null || replaceObj == null)
            {
                throw new System.ArgumentException("Parameter cannot be null", "original");
            }

            if (((string)searchObj).Length == 0 || ((string)replaceObj).Length == 0)
            {
                throw new System.Exception("String length can not be zero");
            }

            object missing = Missing.Value;
            object replaceOne = Word.WdReplace.wdReplaceOne;

            wApp.Selection.Find.ClearFormatting();
            wApp.Selection.Find.Replacement.ClearFormatting();
            wApp.Selection.Find.Execute(ref searchObj, ref missing, ref missing, ref missing, ref missing,
        ref missing, ref missing, ref missing, ref missing, ref replaceObj,
        ref replaceOne, ref missing, ref missing, ref missing, ref missing);
        }

        /* Here we search a pattern like <xxx> */
        private List<string> getSearchString(Word.Document wDoc, string pattern1, string pattern2)
        {
            if (wDoc == null)
            {
                throw new System.ArgumentNullException("wApp", "Parameter can not be null");
            }

            if (pattern1 == null)
            {
                throw new System.ArgumentNullException("pattern1", "Parameter can not be null");
            }

            if (pattern2 == null)
            {
                throw new System.ArgumentNullException("pattern2", "Parameter can not be null");
            }

            if (pattern1.Length == 0 || pattern2.Length == 0)
            {
                throw new System.ArgumentException("parameter can not be empty", "pattern");
            }

            List<string> searchStrList = new List<string>();
            object missing = Missing.Value;

            foreach (Word.Paragraph para in wDoc.Paragraphs)
            {
                string rangStr = para.Range.Text;
                bool found = true;

                while (found)
                {
                    int startPos = rangStr.IndexOf(pattern1);
                    int endPos = rangStr.IndexOf(pattern2);
                    if (startPos > 0 && endPos > 0)
                    {
                        searchStrList.Add(rangStr.Substring(startPos, endPos - startPos + 1));
                        rangStr = rangStr.Substring(endPos + 1);
                    }
                    else
                    {
                        found = false;
                    }
                }
            }

            foreach (Word.Table tbl in wDoc.Tables)
            {
                string rangStr = tbl.Range.Text;
                bool found = true;
                
                while (found)
                {
                    int startPos = rangStr.IndexOf(pattern1);
                    int endPos = rangStr.IndexOf(pattern2);
                    if (startPos > 0 && endPos > 0)
                    {
                        searchStrList.Add(rangStr.Substring(startPos, endPos - startPos + 1));
                        rangStr = rangStr.Substring(endPos+1);
                    }
                    else
                    {
                        found = false;
                    }
                }
            }

            return searchStrList;
        }

        // Set up the BackgroundWorker object by 
        // attaching event handlers. 
        private void InitializeBackgoundWorker()
        {
            backgroundWorker1.DoWork +=
                new DoWorkEventHandler(backgroundWorker1_DoWork);
            backgroundWorker1.RunWorkerCompleted +=
                new RunWorkerCompletedEventHandler(
            backgroundWorker1_RunWorkerCompleted);
            backgroundWorker1.ProgressChanged +=
                new ProgressChangedEventHandler(
            backgroundWorker1_ProgressChanged);
        }

        // This event handler is where the actual,
        // potentially time-consuming work is done.
        private void backgroundWorker1_DoWork(object sender,
            DoWorkEventArgs e)
        {
            // Get the BackgroundWorker that raised this event.
            BackgroundWorker worker = sender as BackgroundWorker;

            foreach (Process proc in Process.GetProcessesByName("winword"))
            {
                proc.Kill();
            }

            //  copy letter format to temp.doc
            object temporaryFile = Path.GetTempFileName();
            File.Copy(Environment.CurrentDirectory + "\\SemteckTripReportTemplate.doc", (string)temporaryFile, true);
            //  create missing object
            object missing = Missing.Value;
            //  create Word application object
            Word.Application wordApp = new Word.ApplicationClass();
            //  create Word document object
            Word.Document aDoc = null;
            //  create & define filename object with temp.doc
            //  if temp.doc available
            if (File.Exists((string)temporaryFile))
            {
                object readOnly = false;
                object isVisible = false;
                //  make visible Word application
                wordApp.Visible = false;
                //  open Word document named temp.doc
                aDoc = wordApp.Documents.Open(ref temporaryFile, ref missing,
                                                ref readOnly, ref missing, ref missing, ref missing,
                                                ref missing, ref missing, ref missing, ref missing,
                                                ref missing, ref isVisible, ref missing, ref missing,
                                                ref missing, ref missing);
                aDoc.Activate();
                //  To find text using a Selection object
                // Create a datatable to store search string and replace string
                // column is for search string
                // row is for replace string
                DataTable dt = new DataTable();
                //findAndReplace(wordApp, "<Date>", textBoxDate.Text);
                List<string> searchStr = getSearchString(aDoc, "<", ">");
                if (searchStr.Capacity == 0)
                {
                    throw new Exception("Template file is currupted!");
                }
                foreach (string str in searchStr)
                {
                    dt.Columns.Add(str);
                }
                DataRow newRow = dt.NewRow();
                newRow["<Customer>"] = textBoxCustomer.Text;
                newRow["<Date>"] = textBoxDate.Text;
                newRow["<Attendee>"] = textBoxAttendee.Text;
                dt.Rows.Add(newRow);

                if (mainForm.myAccess == null)
                {
                    throw new ArgumentNullException();
                }
                DataSet newSet = mainForm.myAccess.GetDBRecord("Customer", textBoxCustomer.Text);
                if (newSet == null)
                {
                    MessageBox.Show("No such customer in database!");
                }
                if (newSet.Tables.Count > 1 || newSet.Tables[0].Rows.Count > 1)
                {
                    MessageBox.Show("There are more than one record (same customer) in database! Please check your database.");
                }

                //  save temp.doc after modified
                object savedFilename = Environment.CurrentDirectory + "\\SemteckTripReport.doc";
                aDoc.SaveAs(ref savedFilename, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                wordApp.Documents.Close(ref missing, ref missing, ref missing);
            }
            else
            {
                MessageBox.Show("File does not exist.", "No File", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            foreach (Process proc in Process.GetProcessesByName("winword"))
            {
                proc.Kill();
            }

            // Assign the result of the computation
            // to the Result property of the DoWorkEventArgs
            // object. This is will be available to the 
            // RunWorkerCompleted eventhandler.
            e.Result = 1;
        }

        // This event handler deals with the results of the
        // background operation.
        private void backgroundWorker1_RunWorkerCompleted(
            object sender, RunWorkerCompletedEventArgs e)
        {
            this.pictureBox1.Image = null;
            //this.pictureBox1.Refresh();

            if (e.Cancelled)
            {
                this.pictureBox1.Image = Properties.Resources.
            }
            // First, handle the case where an exception was thrown.
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message);
            }
            
            
        }

        // This event handler updates the progress bar.
        private void backgroundWorker1_ProgressChanged(object sender,
            ProgressChangedEventArgs e)
        {
            // To do
        }

        private void FormTripReport_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.mainForm.Close_newFormTripReport();
        }
    }
}