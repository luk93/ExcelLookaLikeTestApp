using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelLookaLikeTestApp
{
    public partial class PopUpForm : Form
    {
        public PopUpForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
                MainForm.row_PopUp = int.Parse(rowTextBox.Text);
                MainForm.col_PopUp = int.Parse(columnTextBox.Text);
                MainForm.wsName_PopUp = groupNameTextBox.Text;
                PopUpForm.ActiveForm.Close();
        }

        private void columnTextBox_TextChanged(object sender, EventArgs e)
        {
            if(string.IsNullOrEmpty(columnTextBox.Text)) acceptButton.Enabled = false;
        }

        private void rowTextBox_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(rowTextBox.Text)) acceptButton.Enabled = false;
        }

        private void groupNameTextBox_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(groupNameTextBox.Text)) acceptButton.Enabled = false;
        }
    }
}
