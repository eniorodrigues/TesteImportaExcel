﻿using System;
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

        public static string path;
        public static string excelConnectionString;
        string[] files;
        string conexao;
        string baseDeDados;
        string tabela;
        string caminho;
        string directoryPath;
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        List<string> filesAdionado = new List<string>();
        List<string> colunas = new List<string>();
        int v = 0;
        private void button1_Click(object sender, EventArgs e)
        {

            if (filesAdionado.Count == 0)
            {
                MessageBox.Show("Adicione arquivos");
            }
            else

            label2.Text = "Importando...";

           try
           {
                foreach (string element in filesAdionado)
                {
                    
  
                    MyApp = new Excel.Application();
                    MyBook = MyApp.Workbooks.Open(directoryPath + "\\" + element);

                    excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + directoryPath + "\\" + element + ";Extended Properties=Excel 12.0;";
                    Excel.Application app = new Excel.Application();
                    MyApp.Workbooks.Add("");
                    MyApp.Workbooks.Add(@directoryPath + "\\" + element);
            

                for (int i = 2; i <= MyApp.Workbooks.Count; i++)
                    {
                        progressBar1.Maximum = MyApp.Workbooks.Item[i].Worksheets.Count;
                        MessageBox.Show(MyApp.Workbooks.Item[i].Worksheets.Count.ToString());
                      
                        for (int j = 1; j <= MyApp.Workbooks[i].Worksheets.Count; j++)
                        {
                            if (MyApp.Workbooks[i].Worksheets[j].name != "Input" && MyApp.Workbooks[i].Worksheets[j].name != "Combined")
                            {
                                progressBar1.Value = j ;
                                label2.Text = "Importando arquivo " + element + " da sheet " + MyApp.Workbooks[i].Worksheets[j].name;
 
                                using (OleDbConnection connection =
                                new OleDbConnection(excelConnectionString))
                                {
                                connection.Open();

                                    StringBuilder comandoExcel = new StringBuilder();
                                    for (int h = 0; h < colunas.Count; h++)
                                    {
                                    comandoExcel.Append("[" + Convert.ToString(colunas[h]).Replace(".","#") + "], ");
                                    }

                                        string sheet = MyApp.Workbooks[i].Worksheets[j].name;
                                        string arquivo = element;
                                        string campos = Convert.ToString(comandoExcel);
                                
                                    OleDbCommand command = new OleDbCommand
                                        ("Select " + campos + " '" + sheet + "', ' "+ arquivo +"' FROM [" + MyApp.Workbooks[i].Worksheets[j].name + "$]", connection);

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
                                    //   connection.Close();
                                }

                            }

                        }


                    }
                    if (filesAdionado.Count == 0)
                    {
                        MessageBox.Show("Sem arquivo para adicionar");
                    }
                    else
                    {
                        MessageBox.Show("Importação realizada com sucesso");
                    }



                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro em importar arquivo " + ex.Message);
            }
            filesAdionado.Clear();
            listBox1.Items.Clear();
        }




    public void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
    {
        var backgroundWorker = sender as BackgroundWorker;
        for (int j = 0; j < 100000; j++)
        {
          //  Calculate(j);
            backgroundWorker.ReportProgress((j * 100) / 100000);
        }
    }

        public void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
    {
        progressBar1.Value = e.ProgressPercentage;
    }

    private void button2_Click(object sender, EventArgs e)
        {
            Stream myStream = null;
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "C:\\";
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


                            caminho = openFileDialog1.FileName;
                            directoryPath = Path.GetDirectoryName(openFileDialog1.FileName);
                            files = (openFileDialog1.SafeFileNames);

                            foreach (string file in files)
                            {
                                filesAdionado.Add(file);
                                listBox1.Items.Add(file);
                                carregaLinhas();
                            }

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
            if(textBox3.Text == "")
            {
                MessageBox.Show("Digite o nome da base");

            }
            else
            {

            try
                {
                    baseDeDados = comboBoxBase.SelectedItem.ToString();
                    tabela = textBox3.Text;

                    using (var con = new SqlConnection("Data Source=" + conexao + " ;Initial Catalog=" + baseDeDados + " ;Integrated Security=True"))
                    {
                        con.Open();
                        DataTable databases = con.GetSchema("Databases");
                        string sql = "create table " + tabela + " (";

                        StringBuilder campos = new StringBuilder();

                        for (int i = 0; i < colunas.Count; i++)
                        {
                            campos.Append("[" + Convert.ToString(colunas[i]) + "] varchar (255), ");
                            //     MessageBox.Show(campos.ToString());
                        }
                        campos.Append("[Sheet] varchar (255), [Arquivo] varchar(255) )");
                        sql = sql + campos.ToString();

                        SqlCommand cmd = new SqlCommand(sql, con);
                        cmd.ExecuteNonQuery();
                        MessageBox.Show("Tabela " + tabela + " adicionada com sucesso!");
                        con.Close();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro em importar arquivo " + ex.Message);
                }

            }

        }

        private void comboBoxConexao_SelectedIndexChanged(object sender, EventArgs e)
        {
           
            conexao = comboBoxConexao.Text;

            using (var con = new SqlConnection("Data Source=" + conexao + "; Integrated Security=True;"))
            {
                con.Open();
                DataTable databases = con.GetSchema("Databases");
                comboBoxBase.Items.Clear();
                foreach (DataRow database in databases.Rows)
                {
                     
                    String databaseName = database.Field<String>("database_name");
                    comboBoxBase.Items.Add(databaseName);
                }
            }
        }

        public void carregaLinhas()
        {

            label2.Text = caminho;

            MyApp = new Excel.Application();
            MyBook = MyApp.Workbooks.Open(caminho);

            excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + caminho + ";Extended Properties=Excel 12.0;";
            Excel.Application app = new Excel.Application();
            MyApp.Workbooks.Add("");
            MyApp.Workbooks.Add(caminho);

            if (MyApp.Workbooks[2].Worksheets[1].name != "Input" && MyApp.Workbooks[2].Worksheets[1].name != "Combined")
            {

                for (int k = 1; k <= MyApp.Workbooks[2].Worksheets[1].UsedRange.Columns.Count; k++)
                {

                    if (Convert.ToString((MyApp.Workbooks[2].Worksheets[1].Cells[1, k].Value2)) != null && Convert.ToString(MyApp.Workbooks[2].Worksheets[1].Cells[1, k].Value2.ToString()) != "")
                    {
                        string coluna = Convert.ToString(MyApp.Workbooks[2].Worksheets[1].Cells[1, k].Value2);
                        colunas.Add(coluna);
                        listBox2.Items.Add(coluna);
                    }


                }
            }
            else
            {
                for (int k = 1; k <= MyApp.Workbooks[2].Worksheets[2].UsedRange.Columns.Count; k++)
                {

                    string coluna = Convert.ToString(MyApp.Workbooks[2].Worksheets[2].Cells[1, k].Value2.ToString());
                    colunas.Add(coluna);
                    listBox2.Items.Add(coluna);

                }
            }
        }

        private void button3_Click_1(object sender, EventArgs e)
        {

           

        }

        private void comboBoxConexao_Enter(object sender, EventArgs e)
        {
            comboBoxConexao.Items.Clear();
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
