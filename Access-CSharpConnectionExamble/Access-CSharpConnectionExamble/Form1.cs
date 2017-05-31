using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Access_CSharpConnectionExamble {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
        }
        string AccessConnectionString = "";
        string tableName = "";

        private void btnOpen_Click(object sender, EventArgs e) {
            openFileDialog1.ShowDialog();
            txtOpenFilePath.Text = openFileDialog1.FileName;
            AccessConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;"+
                                       "Data Source='"+openFileDialog1.FileName+"';Persist Security Info=True";
            txtConnectionString.Text = AccessConnectionString;
        }

        public void ConnectToAccessDB() {
            OleDbConnection AccessConnection = new OleDbConnection();
            AccessConnection.ConnectionString = AccessConnectionString;
            try {
                AccessConnection.Open();
                OleDbCommand command = new OleDbCommand();
                command.CommandText = "SELECT * FROM "+tableName+";";
                command.CommandType = CommandType.Text;
                command.Connection = AccessConnection;

                OleDbDataReader dataDB;
                dataDB = command.ExecuteReader();


                DataTable dbTable = new DataTable();
                dbTable.Load(dataDB);
                dgvDataBase.DataSource = dbTable;
                dgvDataBase.Refresh();

                AccessConnection.Close();
            }
            catch (Exception ex) {
                MessageBox.Show("Connection Error!\n\n" + ex.Message);
            }
        }

        private void btnConnect_Click(object sender, EventArgs e) {
            if (AccessConnectionString == "") {
                MessageBox.Show("Choose an Access DataBase!");
            }
            else if (txtTableName.Text == "") {
                MessageBox.Show("Fill the Table Name field!");
            }
            else {
                tableName = txtTableName.Text;
                ConnectToAccessDB();
            }
        }


    }

   

}
