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
        }

        private void buttonCreate_Click(object sender, EventArgs e)
        {
            //  create trip report
            try
            {
                //  Just to kill WINWORD.EXE if it is running
                try
                {
                    foreach (Process proc in Process.GetProcessesByName("winword"))
                    {
                        proc.Kill();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                //  copy letter format to temp.doc
                File.Copy(Environment.CurrentDirectory + "\\SemteckTripReportTemplate.doc", "c:\\temp.doc", true);
                //  create missing object
                object missing = Missing.Value;
                //  create Word application object
                Word.Application wordApp = new Word.ApplicationClass();
                //  create Word document object
                Word.Document aDoc = null;
                //  create & define filename object with temp.doc
                object filename = "c:\\temp.doc";
                //  if temp.doc available
                if (File.Exists((string)filename))
                {
                    object readOnly = false;
                    object isVisible = false;
                    //  make visible Word application
                    wordApp.Visible = false;
                    //  open Word document named temp.doc
                    aDoc = wordApp.Documents.Open(ref filename, ref missing,
                                                    ref readOnly, ref missing, ref missing, ref missing,
                                                    ref missing, ref missing, ref missing, ref missing,
                                                    ref missing, ref isVisible, ref missing, ref missing,
                                                    ref missing, ref missing);
                    aDoc.Activate();
                    //  To find text using a Selection object
                    findAndReplace(wordApp, "<Date>", textBoxDate.Text);
                    //  save temp.doc after modified
                    //aDoc.Save();
                    object savedFilename = Environment.CurrentDirectory + "\\SemteckTripReport.doc";
                    aDoc.SaveAs(ref savedFilename, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
                    wordApp.Documents.Close(ref missing, ref missing, ref missing);
                }
                else
                    MessageBox.Show("File does not exist.", "No File", MessageBoxButtons.OK, MessageBoxIcon.Information);

                try
                {
                    foreach (Process proc in Process.GetProcessesByName("winword"))
                    {
                        proc.Kill();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Error in process.", "Internal Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
    }
}