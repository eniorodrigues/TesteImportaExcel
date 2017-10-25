using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data.Common;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Data.Sql;

namespace TesteImportaExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public static string path = @"C:\Users\enrodrigues\Desktop\4 - Custos - Materiais\DBFC 02-17.xlsb";
        public static string excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;";
        string[] files;
        string conexao;
        string baseDeDados;

        private void button1_Click(object sender, EventArgs e)
        {

            label2.Text = "Importando...";

            Excel.Application app = new Excel.Application();
            app.Workbooks.Add("");

            foreach (string element in files)
            {
                System.Console.WriteLine(element);
 
            app.Workbooks.Add(@"C:\Users\enrodrigues\Desktop\4 - Custos - Materiais\" + element);
 
            for (int i = 2; i <= app.Workbooks.Count; i++)
            {
                    progressBar1.Maximum = app.Workbooks.Count;
                    progressBar1.Value = i;

                    for (int j = 1; j <= app.Workbooks[i].Worksheets.Count; j++)
                {
                        
                     label2.Text = "Importando arquivo " + element + " da sheet " + app.Workbooks[i].Worksheets[j].name;
                    string myString = (app.Workbooks[i].Worksheets[j].Cells[1, 1]).Value2.ToString();
                    
                    using (OleDbConnection connection =
                           new OleDbConnection(excelConnectionString))
                    {

                        OleDbCommand command = new OleDbCommand
                                ("Select Periodo, Ano, Material, [Descricao do Material], Planta, [Custo Médio Total (MAT+MOD+DGF)], [Estoque Final], '" +
                                app.Workbooks[i].Worksheets[j].name+"', '" + element + "'  FROM ["+ app.Workbooks[i].Worksheets[j].name+"$]", connection);

                                 connection.Open();

                        //Create DbDataReader to Data Worksheet
                    using (DbDataReader dr = command.ExecuteReader())
                        {
                                // SQL Server Connection String
                                string sqlConnectionString = "Data Source=" + conexao + ";Initial Catalog=" + baseDeDados + ";Integrated Security=True";


                                // Bulk Copy to SQL Server
                                using (SqlBulkCopy bulkCopy =
                                       new SqlBulkCopy(sqlConnectionString))
                            {
                                bulkCopy.DestinationTableName = "Custos_Totais";
                                bulkCopy.WriteToServer(dr);
                               
                            }
                        }

                    }

                 

                }

                }

             
            }
            MessageBox.Show("Importação realizada com sucesso");
        }
       
        private void button2_Click(object sender, EventArgs e)
        {

            Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "C:\\Users\\enrodrigues\\Desktop\\4 - Custos - Materiais";
            openFileDialog1.Filter = "Csv files (*csv.*)|*csv.*|Excel files (*.xls*)|*.xls*";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;
            openFileDialog1.Multiselect = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((myStream = openFileDialog1.OpenFile()) != null)
                    {
                        using (myStream)
                        {
                            files = (openFileDialog1.SafeFileNames);
                            foreach (string file in files)
                                listBox1.Items.Add(file);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            comboBoxBase.Items.Clear();
            conexao = comboBoxConexao.Text;

            using (var con = new SqlConnection("Data Source="+ conexao+ "; Integrated Security=True;"))
            {
                con.Open();
                DataTable databases = con.GetSchema("Databases");

                foreach (DataRow database in databases.Rows)
                {
                    String databaseName = database.Field<String>("database_name");
                     comboBoxBase.Items.Add(databaseName);
                }
            }
        }

        private void txtConexao_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string myServer = Environment.MachineName;

            DataTable servers = SqlDataSourceEnumerator.Instance.GetDataSources();
            for (int i = 0; i < servers.Rows.Count; i++)
            {
                if (myServer == servers.Rows[i]["ServerName"].ToString())
                {
                    if ((servers.Rows[i]["InstanceName"] as string) != null)
                        comboBoxConexao.Items.Add(servers.Rows[i]["ServerName"] + "\\" + servers.Rows[i]["InstanceName"]);
                    else
                        comboBoxConexao.Items.Add(servers.Rows[i]["ServerName"]);
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            baseDeDados = comboBoxBase.SelectedItem.ToString();
        }
    }
}
