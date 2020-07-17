using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using FolderPermission.Utilities;
using System.IO;

namespace FolderPermission
{
    public partial class Form1 : Form
    {
        private static runStatus status = runStatus.Idle;
        private static DataTable dt = new DataTable();
        enum runStatus
        {
            Idle,
            Started,
            Loading,
            Loaded,
            SetPermission,
            PermissionGranted,
            Auditing,
            Audited
        }
        enum auditResult
        {
            Pass = 'O',
            Fail = 'X'
        }
        public Form1()
        {
            InitializeComponent();
            Config.GetConfigurationValue();
            labelStatus.Text = runStatus.Idle.ToString();
            txtOutput.Text = Path.Combine(Path.GetDirectoryName(txtInputFile.Text), "Permission.xlsx");
            dataGridView.ScrollBars = ScrollBars.Both;
        }

        private async void checkBtn_Click(object sender, EventArgs e)
        {
            updateStatus(runStatus.Started);

            dt = ExcelData.readData(txtInputFile.Text);
            dt.Rows.Clear();
            dataGridView.DataSource = dt;
            dataGridView.Refresh();
            dataGridView.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridView.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

            updateStatus(runStatus.Loading);

            await Controller.Action.checkAsync(dt);
            
            dataGridView.Refresh();
            updateStatus(runStatus.Loaded);
            exportBtn.Enabled = true;
        }
        private void loadBtn_Click(object sender, EventArgs e)
        {
            updateStatus(runStatus.Started);

            dt = ExcelData.readData(txtInputFile.Text);
            dataGridView.DataSource = dt;
            dataGridView.Refresh();
            dataGridView.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridView.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

            updateStatus(runStatus.Loaded);
            exportBtn.Enabled = true;
            setBtn.Enabled = true;
        }
        private async void setBtn_Click(object sender, EventArgs e)
        {
            updateStatus(runStatus.SetPermission);

            status = runStatus.SetPermission;

            await Controller.Action.setPermissionAsync(dt);
            dataGridView.Refresh();

            updateStatus(runStatus.PermissionGranted);

            status = runStatus.Idle;
            MessageBox.Show("Done", "Status");
        }
        private async void auditBtn_Click(object sender, EventArgs e)
        {
            updateStatus(runStatus.Started);

            dt = ExcelData.readData(txtInputFile.Text);
            status = runStatus.Auditing;
            updateStatus(runStatus.Auditing);

            dataGridView.DataSource = dt;
            dataGridView.Refresh();
            dataGridView.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;
            dataGridView.Columns[4].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;

            await Controller.Action.auditAsync(dt, auditResult.Pass.ToString(), auditResult.Fail.ToString());

            dataGridView.Refresh();
            status = runStatus.Idle;
            updateStatus(runStatus.Audited);
            exportBtn.Enabled = true;
        }
        private void exportBtn_Click(object sender, EventArgs e)
        {
            if (txtOutput.Text != "")
            {
                Controller.Action.exportToExcel(txtOutput.Text, dt);
                MessageBox.Show("Export Done", "Status");
            } else
            {
                MessageBox.Show("No Output File", "Error");
            }
            
        }
        private void BrowseBtn_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel File (*.XLS;*.XLSX;*.XLSM;*.XLM)|*.XLS;*.XLSX;*.XLSM;*.XLM|" +
                        "All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;
                openFileDialog.Multiselect = false;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    foreach (String file in openFileDialog.FileNames)
                    {
                        txtInputFile.Text = file;
                        txtOutput.Text = Path.Combine(Path.GetDirectoryName(file), "Permission.xlsx");
                    }
                }
            }
        }
        private void dataGridView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            // If the column is the Artist column, check the
            // value.
            if (dataGridView.Columns[e.ColumnIndex].Name.Contains(Config.fileServer))
            {
                if (e.Value != null)
                {
                    
                    if (status == runStatus.SetPermission)
                    {
                        e.CellStyle.BackColor = Color.LightGreen;
                    } else if (status == runStatus.Auditing)
                    {
                        string stringValue = (string)e.Value;
                        if (stringValue.Contains("[" + auditResult.Pass.ToString() + "]"))
                        {
                            e.CellStyle.BackColor = Color.LightGreen;
                        } else if (stringValue.Contains("[" + auditResult.Fail.ToString() + "]"))
                        {
                            e.CellStyle.BackColor = Color.Red;
                        }
                    }
                }
            }
        }
        private void updateStatus(runStatus message)
        {
            labelStatus.Text = message.ToString();
            labelStatus.Refresh();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (txtSearch.Text != "")
            {
                string rowFilter = string.Format("[{0}] LIKE '*{1}*' OR [{2}] LIKE '*{3}*' OR [{4}] LIKE '*{5}*'",
                    "Name", txtSearch.Text,
                    "User ID", txtSearch.Text,
                    "UserAccount", txtSearch.Text
                    );
                (dataGridView.DataSource as DataTable).DefaultView.RowFilter = rowFilter;
            } else
            {
                (dataGridView.DataSource as DataTable).DefaultView.RowFilter = "";
            }
            
        }

        
    }
}
