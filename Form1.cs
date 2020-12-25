using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using FolderPermission.Utilities;
using System.IO;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using System.Collections.Generic;

namespace FolderPermission
{
    public partial class Form1 : Form
    {
        private static runStatus status = runStatus.Idle;
        private static DataTable dt = new DataTable();
        private static string initFilePath = "";
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
            if (dataGridView.Columns[e.ColumnIndex].Name.ToUpper().Contains(Config.fileServer.ToUpper()))
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

        private void exitToolStripMenuItem_Click(object send, EventArgs evt)
        {
            Application.Exit();
        }

        private void generateInitialFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string caption = "Generate Initial file";
            string text = "Please specify the path";
            Form prompt = new Form()
            {
                Width = 500,
                Height = 150,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = caption,
                StartPosition = FormStartPosition.CenterScreen
            };
            Label textLabel = new Label() { Left = 50, Top = 20, Width = 400, Text = text };
            TextBox textBox = new TextBox() { Left = 50, Top = 50, Width = 400 };
            Label depthLabel = new Label() { Left = 50, Top = 80, Width = 50, Text = "Depth" };
            TextBox depthBox = new TextBox() { Left = 100, Top = 80, Width =50, Text= "1" };
            Button confirmation = new Button() { Text = "Ok", Left = 350, Width = 100, Top = 70, DialogResult = DialogResult.OK };

            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(textLabel);
            prompt.Controls.Add(depthLabel);
            prompt.Controls.Add(depthBox);
            prompt.AcceptButton = confirmation;

            initFilePath = prompt.ShowDialog() == DialogResult.OK ? textBox.Text : "";
            int number;
            bool success = Int32.TryParse(depthBox.Text, out number);
            if (success)
            {
                promptClick(number);
            } else
            {
                MessageBox.Show("Depth Must be Number", "OSDIT");
            }
            
        }
        private void promptClick(int depth = 1, int ignore=0)
        {
            if (initFilePath != "")
            {
                List<string> folders = new List<string>();
                Permission.searchDir(initFilePath, folders, depth);
                DataTable initdt = new DataTable();
                initdt.Columns.Add("Name", typeof(string));
                initdt.Columns.Add("Role", typeof(string));
                initdt.Columns.Add("E-mail", typeof(string));
                initdt.Columns.Add("User ID", typeof(string));
                initdt.Columns.Add("UserAccount", typeof(string));
                initdt.Columns.Add("Project", typeof(string));
                foreach (string folder in folders)
                {
                    initdt.Columns.Add(folder, typeof(string));
                }
                initdt.Rows.Add();
                Controller.Action.exportToExcel(txtOutput.Text, initdt);
                MessageBox.Show("File has been generated", "Status");
            }
            
        }
        private void aboutUsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Next-ATM PTTDIGITAL", "OSDIT");
        }
    }
}
