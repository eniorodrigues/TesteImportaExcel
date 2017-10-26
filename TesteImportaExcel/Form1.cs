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
    public partial class testeBtn : Form
    {
        public testeBtn()
        {
            InitializeComponent();
        }

        public static string path = @"C:\Users\enrodrigues\Desktop\4 - Custos - Materiais\DBFC 02-17.xlsb";
        public static string excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;";
        string[] files;
        string conexao;
        string baseDeDados;
        string tabela;
        string caminho;

        List<string> colunas = new List<string>();

        private void button1_Click(object sender, EventArgs e)
        {

            label2.Text = "Importando...";

            Excel.Application app = new Excel.Application();
            app.Workbooks.Add("");

            foreach (string element in files)
            {

            app.Workbooks.Add(@caminho);
 
            for (int i = 2; i <= app.Workbooks.Count; i++)
            {
                    progressBar1.Maximum = app.Workbooks.Count;
                    progressBar1.Value = i;

                    for (int j = 1; j <= app.Workbooks[i].Worksheets.Count; j++)
                {
                        
                     label2.Text = "Importando arquivo " + element + " da sheet " + app.Workbooks[i].Worksheets[j].name;

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
                                bulkCopy.DestinationTableName = tabela;
                                bulkCopy.WriteToServer(dr);
                               
                            }
                        }
                            connection.Close();
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
                            caminho = (openFileDialog1.FileName);
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

        }

        private void txtConexao_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
      
        }
     
        private void button4_Click(object sender, EventArgs e)
        {
            baseDeDados = comboBoxBase.SelectedItem.ToString();
            tabela = textBox3.Text;

            using (var con = new SqlConnection("Data Source=" + conexao + ";Initial Catalog=" + baseDeDados + ";Integrated Security=True"))
            {
                con.Open();
                DataTable databases = con.GetSchema("Databases");
                string sql = "create table " + tabela + " (";

                StringBuilder campos = new StringBuilder();

                for (int i = 0; i < colunas.Count; i++)
                {
                MessageBox.Show(colunas[i].ToString());

                    campos.Append("[" + colunas[i].ToString() + "] varchar (255), ");

                }
                campos.Append("[Sheet] varchar (255), [Arquivo] varchar(255) )");
                sql = sql + campos.ToString();
                MessageBox.Show(sql);
            

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.ExecuteNonQuery();
                MessageBox.Show("Tabela " + tabela + " adicionada com sucesso!");
                con.Close();
            }

        }

        private void comboBoxConexao_SelectedIndexChanged(object sender, EventArgs e)
        {
             comboBoxBase.Items.Clear();
             conexao = comboBoxConexao.Text;

            using (var con = new SqlConnection("Data Source=" + conexao + "; Integrated Security=True;"))
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


        private void button3_Click_1(object sender, EventArgs e)
        {
            label2.Text = caminho;

            Excel.Application app = new Excel.Application();
            app.Workbooks.Add("");
           
                    app.Workbooks.Add(@caminho);

                         for (int k = 1; k <= app.Workbooks[2].Worksheets[1].UsedRange.Columns.Count; k++)
                        {
                        
                            string coluna = app.Workbooks[2].Worksheets[1].Cells[1, k].Value2.ToString();
                            colunas.Add(coluna);
                            listBox2.Items.Add(coluna);
                       
                        }

        }

        private void comboBoxConexao_Enter(object sender, EventArgs e)
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

        private void comboBoxBase_SelectedIndexChanged(object sender, EventArgs e)
        {
            baseDeDados = comboBoxBase.SelectedItem.ToString();
        }
    }
}
