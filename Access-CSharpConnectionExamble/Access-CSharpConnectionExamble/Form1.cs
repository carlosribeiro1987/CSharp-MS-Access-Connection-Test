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
        
        string AccessConnectionString = ""; //The connection string will be stored here
        string tableName = ""; //The name of the table populated in DataGridView

        private void btnOpen_Click(object sender, EventArgs e) {
            openFileDialog1.ShowDialog();
            txtOpenFilePath.Text = openFileDialog1.FileName;
            AccessConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;"+
                                       "Data Source='"+openFileDialog1.FileName+"';Persist Security Info=True";
            txtConnectionString.Text = AccessConnectionString;
        }

        public void ConnectToAccessDB() {
            OleDbConnection AccessConnection = new OleDbConnection(); //Create a new connection
            AccessConnection.ConnectionString = AccessConnectionString; //Specifies the connection string
            try {
                AccessConnection.Open(); //Open the connection
                OleDbCommand command = new OleDbCommand(); 
                command.CommandText = "SELECT * FROM "+tableName+";"; //Select all fields from the table spacified in "Table Name" textBox
                command.CommandType = CommandType.Text; //Sets command type to text
                command.Connection = AccessConnection; 

                OleDbDataReader dataDB; //Create a data reader to receive the DB data
                dataDB = command.ExecuteReader(); //Execute the command


                DataTable dbTable = new DataTable(); //Create a DataTable to receive data
                dbTable.Load(dataDB); // Load data return by command
                dgvDataBase.DataSource = dbTable; //Populates datagridView with the fields of DB table
                dgvDataBase.Refresh(); //Refresh DataGridView

                AccessConnection.Close(); //Close connection
            }
            catch (Exception ex) {
                MessageBox.Show("Connection Error!\n\n" + ex.Message); //Returns MessageBox error 
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
