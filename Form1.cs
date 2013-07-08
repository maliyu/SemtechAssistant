using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using maliyuAccess2003Dll;

namespace SemtechAssistant
{
    public partial class SemtechAssistant : Form
    {
        #region private Variable
        private OleDbConnection dbOleConn = null;
        maliyuAccess myAccess = null;
        string searchString = null;
        LinkedList<DataGridView> dgvList = null;
        #endregion

        public SemtechAssistant()
        {
            InitializeComponent();
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
                dbOleConn.Open();
                toolStripStatusLabel1.Text = openFileDialog1.SafeFileName;
            }
        }

        private void buttonSearch_Click(object sender, EventArgs e)
        {
            if (myAccess == null)
            {
                myAccess = new maliyuAccess(dbOleConn);
            }
            
            DataSet searchResult = myAccess.QueryWholeDB(searchString);

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

                foreach (DataTable tbl in searchResult.Tables)
                {
                    DataGridView newDGV = new DataGridView();
                    //newDGV.AutoSize = true;
                    newDGV.Anchor = AnchorStyles.Left;
                    newDGV.Dock = DockStyle.Bottom;
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

                    this.Controls.Add(newDGV);
                    //this.Invalidate();
                    dgvList.AddLast(newDGV);
                }
                //this.Refresh();
            }
        }

        private void textBoxSearch_TextChanged(object sender, EventArgs e)
        {
            searchString = ((TextBox)sender).Text;
        }

    }
}