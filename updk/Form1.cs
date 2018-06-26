using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace updk
{
    public partial class connectionStringForm : Form
    {
        public connectionStringForm()
        {
            InitializeComponent();
           // DataBaseNameTextBox.Text = @"Data Source=.;Initial Catalog=UPDK5;Integrated Security=True";
            AuthModeComboBox.SelectedIndex = 0;
            ServerNameTextBox.Text = @".\sqlexpress";
            DataBaseNameTextBox.Text = "UPDK5";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string connectionString = String.Format(@"
                Data Source={0};
                Initial Catalog={1};
                Integrated Security={2}", ServerNameTextBox.Text, DataBaseNameTextBox.Text, AuthModeComboBox.SelectedIndex == 0 ? "True" :
                    String.Format("False;User Id = {0}; Password = {1};", LoginTextBox.Text, PasswordTextBox.Text));
            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();
                    connection.Close();
                }
            }
            catch (System.Data.SqlClient.SqlException)
            {
                MessageBox.Show("Нет подключения к базе данных", "Ошибка");
                return;
            }

            MainForm mf = new MainForm(connectionString);//, beginDate.Value, endDate.Value);
           // MainForm mf = new MainForm(@"Data Source=.\SQLEXPRESS;Initial Catalog=UPDK5;Integrated Security=True");
            mf.ShowDialog();
        }

        private void connectionStringForm_Load(object sender, EventArgs e)
        {

        }

        private void AuthModeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (AuthModeComboBox.SelectedIndex == 0)
            {
                LoginTextBox.Enabled = false;
                PasswordTextBox.Enabled = false;
            }
            else
            {
                LoginTextBox.Enabled = true;
                PasswordTextBox.Enabled = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            HelpForm helpForm = new HelpForm();
            helpForm.ShowDialog();
        }
    }
}
