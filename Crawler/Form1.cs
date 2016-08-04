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

namespace Crawler
{
    public partial class Form1 : Form
    {
        string connectionString = "Server=tcp:zjding.database.windows.net,1433;Initial Catalog=Costco;Persist Security Info=False;User ID=zjding;Password=G4indigo;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;";

        SqlCommand cmdSaleTax;
        SqlDataAdapter daSaleTax;
        SqlCommandBuilder cmbSaleTax;
        DataSet dsSaleTax;
        DataTable dtSaleTax;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string sqlString = @"SELECT ReportID, FromDate, ToDate, NumberOfTransactions, StateSaleTax, CitySaleTax, CountySaleTax, StateTaxSubmitted, CityTaxSubmitted, CountyTaxSubmitted 
                                 FROM SaleTax 
                                 Order by FromDate DESC";

            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            cmdSaleTax = new SqlCommand(sqlString, connection);
            daSaleTax = new SqlDataAdapter(cmdSaleTax);
            cmbSaleTax = new SqlCommandBuilder(daSaleTax);
            dsSaleTax = new DataSet();
            daSaleTax.Fill(dsSaleTax, "tbSaleTax");
            dtSaleTax = dsSaleTax.Tables["tbSaleTax"];
            connection.Close();
        }
    }
}
