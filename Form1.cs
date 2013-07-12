using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using maliyuAccess2003Dll;
using SemtechAssistant.Resources;

namespace SemtechAssistant
{
    public partial class SemtechAssistant : Form
    {
        #region private Variable
        public OleDbConnection dbOleConn = null;
        public maliyuAccess myAccess = null;
        public string searchString = null;
        public LinkedList<DataGridView> dgvList = null;
        private FormTripReport newFormTripReport = null;
        protected DataSet searchResult = null;
        #endregion

        public SemtechAssistant()
        {
            InitializeComponent();

            InitializeBackgoundWorker();
        }

        private void buttonLoadDB_Click(object sender, EventArgs e)
        {
            string dbPathName = null;
            // Displays an OpenFileDialog so the user can select a Cursor.
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Microsoft Access 2003|*.mdb";
            openFileDialog1.Title = "Select an Access 2003 Database";

            // Show the Dialog.
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                dbPathName = openFileDialog1.FileName;
                //MessageBox.Show(openFileDialog1.FileName);
                // Connect to the data source.
                if (dbOleConn != null)
                {
                    dbOleConn.Close();
                    dbOleConn = null;
                }
                dbOleConn = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + dbPathName + ";");
                //dbOleConn.Open();
                toolStripStatusLabel1.Text = openFileDialog1.SafeFileName;
                myAccess = new maliyuAccess(dbOleConn);

                this.buttonSearch.Enabled = true;
                this.buttonTripReport.Enabled = true;
            }
        }

        private void buttonSearch_Click(object sender, EventArgs e)
        {
            if (myAccess == null)
            {
                MessageBox.Show("Please select one database before search!");
                return;
            }

            this.pictureBox1.Visible = true;
            this.pictureBox1.Image = Properties.Resources.Animation;

            backgroundWorker1.RunWorkerAsync();
            //DataSet searchResult = myAccess.QueryWholeDB(searchString);
        }

        private void textBoxSearch_TextChanged(object sender, EventArgs e)
        {
            searchString = ((TextBox)sender).Text;
        }

        private void buttonTripReport_Click(object sender, EventArgs e)
        {
            if (newFormTripReport == null)
            {
                newFormTripReport = new FormTripReport(this);
                newFormTripReport.Show();
            }
        }

        public void Close_newFormTripReport()
        {
            newFormTripReport.Dispose();
            newFormTripReport = null;
        }

        private void SemtechAssistant_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (dbOleConn != null)
            {
                dbOleConn.Close();
                dbOleConn.Dispose();
                dbOleConn = null;
            }           
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
            //backgroundWorker1.ProgressChanged +=
            //    new ProgressChangedEventHandler(
            //backgroundWorker1_ProgressChanged);
        }

        // This event handler is where the actual,
        // potentially time-consuming work is done.
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            searchResult = myAccess.QueryWholeDB(searchString);
        }

        // This event handler deals with the results of the
        // background operation.
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.pictureBox1.Image = null;
            this.pictureBox1.Visible = false;

            if (searchResult.Tables.Count <= 0)
            {
                MessageBox.Show("Nothing found!");
            }
            else
            {
                if (dgvList == null)
                {
                    dgvList = new LinkedList<DataGridView>();
                }
                else
                {
                    foreach (DataGridView dgv in dgvList)
                    {
                        dgv.DataSource = null;
                        this.Controls.Remove(dgv);
                    }
                    dgvList.Clear();
                    //this.Refresh();
                }

                Form disForm = new Form();
                disForm.Text = "Search Result";
                disForm.WindowState = FormWindowState.Maximized;

                foreach (DataTable tbl in searchResult.Tables)
                {
                    DataGridView newDGV = new DataGridView();
                    //newDGV.AutoSize = true;
                    newDGV.Anchor = AnchorStyles.Left;
                    newDGV.Dock = DockStyle.Top;
                    newDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                    newDGV.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                    newDGV.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                    newDGV.CellBorderStyle = DataGridViewCellBorderStyle.Single;
                    //newDGV.Dock = DockStyle.Left;

                    newDGV.DataSource = tbl;
                    newDGV.Name = tbl.TableName;

                    newDGV.ColumnHeadersDefaultCellStyle.BackColor = Color.Navy;
                    newDGV.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                    newDGV.ColumnHeadersDefaultCellStyle.Font = new Font(newDGV.Font, FontStyle.Bold);
                    newDGV.BorderStyle = BorderStyle.Fixed3D;

                    disForm.Controls.Add(newDGV);
                    //this.Invalidate();
                    dgvList.AddLast(newDGV);
                }
                //this.Refresh();

                disForm.Show();
            }
        }
    }
}