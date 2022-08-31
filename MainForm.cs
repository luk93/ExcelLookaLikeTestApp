using OfficeOpenXml;
using System.Data;
using System.Reflection;

namespace ExcelLookaLikeTestApp
{
    public partial class MainForm : Form
    {
        DataTable dataTable = new();
        FileInfo file = null;
        List<string> rowHeaders = new List<string>();
        bool autoUpdate = true;

        public static int row_PopUp = 0;
        public static int col_PopUp = 0;
        public static string wsName_PopUp = "";
        public MainForm()
        {
            InitializeComponent();
        }

        private void SomeMenuToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void MainForm_Load(object sender, EventArgs e)
        {

        }

        private async void ImportFromExcelToolStripMenuItem_ClickAsync(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            OpenFileDialog openFileDialog1 = new()
            {
                InitialDirectory = @"c:\Users\localadm\Desktop",
                Title = "Select some  excel file",
                CheckFileExists = true,
                CheckPathExists = true,
                DefaultExt = "xlsx",
                Filter = "Excel file (*.xlsx)|*.xlsx",
                RestoreDirectory = true,
                ReadOnlyChecked = true,
                ShowReadOnly = true,
            };
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                file = new FileInfo(openFileDialog1.FileName);
                PopUpForm popup = new PopUpForm();
                if (popup.ShowDialog() == DialogResult.OK)
                {
                    if (file.Exists && !IsFileLocked(file.FullName))
                    {
                        dataTable = await LoadFromExcel(file,row_PopUp,col_PopUp,wsName_PopUp);
                        if (tabControl.SelectedIndex == 0)
                        {
                            await FillGridWithData(dataGridView, dataTable, rowHeaders);
                            FillRowHeaderNames(dataGridView, rowHeaders);
                        }
                        else if (tabControl.SelectedIndex == 1)
                        {
                            await CopyDataTableToListView(dataTable, listView);
                        }

                        toolStripStatusLabel1.Text = "Got data from: " + file.FullName + " successully! ";
                    }
                    else toolStripStatusLabel1.Text = "File not exist or is being used!";
                }
                else
                {
                    toolStripStatusLabel1.Text = "Column and Row not defined!";
                }
            }
        }
        private static async Task<DataTable> LoadFromExcel(FileInfo file, int row, int col, string wsName)
        {
            int defaultCol = col;
            int defaultRow = row;
            DataTable output = new();
            int colAmount = 0;
            using var exPackage = new ExcelPackage(file);
            await exPackage.LoadAsync(file);
            var ws = exPackage.Workbook.Worksheets[wsName];
            //Get Columns
            while (!string.IsNullOrWhiteSpace(ws.Cells[row, col].Value?.ToString()))
            {
                output.Columns.Add(ws.Cells[row, col].Value.ToString());
                col++;
                colAmount++;
            }
            col = defaultCol;
            row = defaultRow + 1;
            //Get Rows
            while (!string.IsNullOrWhiteSpace(ws.Cells[row, col].Value?.ToString()))
            {
                output.Rows.Add(ws.Cells[row, col].Value.ToString());
                for (int i = 0; i < colAmount; i++)
                {
                    if (!string.IsNullOrWhiteSpace(ws.Cells[row, col + i].Value?.ToString()))
                    {
                        output.Rows[row - (defaultRow + 1)][i] = ws.Cells[row, col + i].Value.ToString();
                    }
                    else
                    {
                        output.Rows[row - (defaultRow + 1)][i] = "";
                    }

                }
                row++;
            }
            return output;
        }
        private void FilterDataToolStripMenuItem_Click(object sender, EventArgs e)
        {
            for (int i = 1; i < dataTable.Rows.Count; i++)
            {
                dataTable.Rows[1][i] = "some data" + i.ToString();
            }
        }

        private async void Timer1_Tick(object sender, EventArgs e)
        {
            if(file != null && !IsFileLocked(file.FullName) && autoUpdate && row_PopUp != 0)
            {
                try
                {
                    dataTable = await LoadFromExcel(file, row_PopUp, col_PopUp, wsName_PopUp);
                    if (tabControl.SelectedIndex == 0)
                    {
                        await FillGridWithData(dataGridView, dataTable, rowHeaders);
                        FillRowHeaderNames(dataGridView, rowHeaders);
                    }
                    else if (tabControl.SelectedIndex == 1)
                    {
                        CopyDataTableToListView(dataTable, listView);
                    }
                    toolStripStatusLabel1.Text = "Data refreshed from: " + file.FullName + " successully! Done:"+ System.DateTime.Now.ToString();
                }
                catch (Exception ex)
                {
                    toolStripStatusLabel1.Text = "Error while trying to refresh data from excel: " + ex.Message;
                }
            }
        }
        public static bool IsFileLocked(string filePath)
        {
            try
            {
                var stream = File.OpenRead(filePath);
                return false;
            }
            catch (IOException)
            {
                return true;
            }
        }
        private void autoUpdateButton_Click_1(object sender, EventArgs e)
        {
            if (autoUpdate)
            {
                autoUpdate = false;
                autoUpdateButton.BackColor = Color.Red;
                autoUpdateLabel.Text = "Auto Update OFF";

            }
            else
            {
                autoUpdate = true;
                autoUpdateButton.BackColor = Color.Green;
                autoUpdateLabel.Text = "Auto Update ON";
            }
        }
        private static async Task FillGridWithData(DataGridView dataGridView, DataTable dataTable, List<string> headers)
        {
            dataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            
            dataGridView.DataSource = dataTable;
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                if (row.IsNewRow) continue;
                headers.Add(row.Cells[0].Value.ToString());
            }
            dataTable.Columns.RemoveAt(0);
        }
        private void FillRowHeaderNames(DataGridView dataGridView, List<string> headers)
        {
            int i = 0;
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                row.HeaderCell.Value = headers[i];
                i++;
            }
            dataGridView.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            dataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
        }
        private static async Task CopyDataTableToListView(DataTable data, ListView listView)
        {
            listView.BeginUpdate();
            if (listView.Columns.Count != data.Columns.Count)
            {
                listView.Columns.Clear();
                //else columns
                foreach (DataColumn column in data.Columns)
                {
                    listView.Columns.Add(column.ColumnName); 
                }
            }
            // clear rows
            listView.Items.Clear();
            //// load rows
            bool isAlternateItem = true;
            foreach (DataRow row in data.Rows)
            {
                ListViewItem item = new ListViewItem(row[0].ToString());
                if (isAlternateItem)
                {
                    item.BackColor = Color.WhiteSmoke;
                    isAlternateItem = false;
                }
                else
                {
                    item.BackColor = Color.AliceBlue;
                    isAlternateItem = true;
                }
                for (int i = 1; i < data.Columns.Count; i++)
                {
                    item.SubItems.Add(row[i].ToString());
                }
                listView.Items.Add(item);
            }
            listView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
            listView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
            listView.EndUpdate();
        }
    }
}