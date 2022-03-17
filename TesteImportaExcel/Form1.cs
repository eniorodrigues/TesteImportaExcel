using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data.Common;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Data.Sql;
using Microsoft.VisualBasic.FileIO;
using System.Runtime.InteropServices;
using Microsoft.Win32;

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
        public string[] files;
        public string conexao;
        public string baseDeDados;
        public string tabela;
        public string caminho;
        public string directoryPath;
        private static Excel.Application MyApp = null;
        public List<string> filesAdionado = new List<string>();
        public List<string> colunas = new List<string>();
        public List<string> colunasCreate = new List<string>();
        public string tipoArquivo;
        DataTable csvData = new DataTable();
        public DataTable arquivoCSV;
        string ext;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        Stream myStream = null;
        //string[] campos;
        bool adicionado = false;

        private void button1_Click(object sender, EventArgs e)
        {

            if (filesAdionado.Count == 0)
            {
                MessageBox.Show("Adicione arquivos");
            }
            else
                label2.Text = "Importando...";
            if (ext != ".csv")
            {
                try
                {
                    foreach (string element in filesAdionado)
                    {

                        MyApp = new Excel.Application();
                        object misValue = System.Reflection.Missing.Value;
                        excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + directoryPath + "\\" + element + ";Extended Properties=Excel 12.0;";
                        MyApp.Workbooks.Add("");
                        MyApp.Workbooks.Add(@directoryPath + "\\" + element);

                        xlWorkBook = MyApp.Workbooks.Add(misValue);
                        xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                        xlWorkSheet.EnableSelection = Microsoft.Office.Interop.Excel.XlEnableSelection.xlNoSelection;

                        for (int i = 2; i <= MyApp.Workbooks.Count; i++)
                        {
                            for (int j = 1; j <= MyApp.Workbooks[i].Worksheets.Count; j++)
                            {

                                string emptyString = Convert.ToString((MyApp.Workbooks[i].Worksheets[j].Cells[1, 1]).Value2);

                                if (emptyString != null)
                                {

                                    if (MyApp.Workbooks[i].Worksheets[j].name != "Input" && MyApp.Workbooks[i].Worksheets[j].name != "Combined")
                                    {


                                        label2.Text = "Importando arquivo " + element + " da sheet " + MyApp.Workbooks[i].Worksheets[j].name;

                                        using (OleDbConnection connection =
                                        new OleDbConnection(excelConnectionString))
                                        {
                                            connection.Open();

                                            StringBuilder comandoExcel = new StringBuilder();
                                            for (int h = 0; h < colunas.Count; h++)
                                            {
                                                comandoExcel.Append("[" + Convert.ToString(colunas[h]).Replace(".", "#") + "], ");
                                            }

                                            string sheet = MyApp.Workbooks[i].Worksheets[j].name;
                                            string arquivo = element;
                                            string campos = Convert.ToString(comandoExcel);

                                            OleDbCommand command = new OleDbCommand
                                                ("Select " + campos + " '" + sheet + "', ' " + arquivo + "' FROM [" + MyApp.Workbooks[i].Worksheets[j].name + "$]", connection);
                                            //MessageBox.Show(command.CommandText);
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
                                        }


                                    }
                                }

                            }


                        }

                        if (filesAdionado.Count == 0)
                        {
                            MessageBox.Show("Sem arquivo para adicionar");
                        }
                    }

                    try
                    {
                        MyApp.Quit();
                        MyApp.Workbooks.Close();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }


                    Marshal.ReleaseComObject(MyApp.Workbooks);
                    Marshal.ReleaseComObject(MyApp);
                    MessageBox.Show("Importação realizada com sucesso");

                }

                catch (Exception ex)
                {
                    MessageBox.Show("Erro em importar arquivo " + ex.Message);
                }
                filesAdionado.Clear();
                listBox1.Items.Clear();

                try
                {
                    MyApp.Quit();
                    MyApp.Workbooks.Close();
                    Marshal.ReleaseComObject(MyApp.Workbooks);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                Marshal.ReleaseComObject(MyApp);

            }

            //////////////////////////CSVCSVCSVCSV
            else
            {

                foreach (string element in filesAdionado)
                {
                    using (StreamReader sr = new StreamReader(@directoryPath + "\\" + element))
                    {
                        DataTable dataTable = new DataTable();

                        string linha = "";

                        linha = sr.ReadLine();
                        string[] cab = linha.Split(new String[] { "\",\"", ",\"", ";" }, StringSplitOptions.RemoveEmptyEntries);

                        for (int i = 0; i < cab.Length; i++)
                        {
                            dataTable.Columns.Add(cab[i].Trim('"'));
                        }
                        dataTable.Columns.Add("Sheet");
                        dataTable.Columns.Add("Arquivo");

                        while (!sr.EndOfStream)
                        {
                            linha = sr.ReadLine();
                            try
                            {
                                string[] campos = linha.Split(new String[] { "\",\"", ",\"", ";" }, StringSplitOptions.RemoveEmptyEntries);
                                DataRow row = dataTable.NewRow();
                                int i = 0;

                                for ( i = 0; i < campos.Length; i++)
                                {
                                    row[i] = campos[i].Trim('"');
                                }
                                row[i] = element;
                                row[i + 1] = element;
                                dataTable.Rows.Add(row);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show(ex.Message);
                            }

                        }


                        using (SqlConnection cn = new SqlConnection("Data Source=" + conexao +";Initial Catalog=" + baseDeDados +";Integrated Security=True"))
                        {
                            cn.Open();
                            using (SqlBulkCopy copy = new SqlBulkCopy(cn))
                            {

                                copy.DestinationTableName = tabela;
                                copy.WriteToServer(dataTable);
                            }
                            cn.Close();
                        }

                    }

                }
                filesAdionado.Clear();
                listBox1.Items.Clear();
            }
        }



        public void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            var backgroundWorker = sender as BackgroundWorker;
            for (int j = 0; j < 100000; j++)
            {
                backgroundWorker.ReportProgress((j * 100) / 100000);
            }

        }

        public void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
        }

        private void button2_Click(object sender, EventArgs e)
        {

            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "C:\\";
            openFileDialog1.Filter = "Csv files (*.csv*)|*.csv*|Excel files (*.xls*)|*.xls*";
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

                            ext = Path.GetExtension(openFileDialog1.FileName);

                            if (ext == ".csv")
                            {
                                // objeto para leitura de arquivo texto
                                StreamReader sr = new StreamReader(openFileDialog1.FileName);
                                string linha = "";
                                int lin = 0;

                                // lê a linha de títulos (nomes dos campos)
                                // enquanto não for fim do arquivo

                                //cabeçalho
                                linha = sr.ReadLine();
                                string[] cab = linha.Split(new String[] { "\",\"", ",\"", ";" }, StringSplitOptions.RemoveEmptyEntries);
                                //
                                for (int i = 0; i < cab.Length; i++)
                                {
                                //    dataGridView1.Columns.Add(cab[i].Trim('"'), cab[i].Trim('"'));
                                }

                                while (!sr.EndOfStream)
                                {
                                    // lê a linha atual do arquivo e avança para a próxima
                                    linha = sr.ReadLine();
                                    // quebra a linha no caractere ";" e retorna um array contendo as partes
                                    string[] campos = linha.Split(new String[] { "\",\"", ",\"", ";" }, StringSplitOptions.RemoveEmptyEntries);
                                    // mostra no grid
                                  //  dataGridView1.RowCount++;
                                    for (int i = 0; i < campos.Length; i++)
                                    {
                                   //     dataGridView1.Rows[lin].Cells[i].Value = campos[i].Trim('"');
                                    }
                                    lin++;
                                }
                                sr.Close();
                                sr.Dispose();
                            }


                            caminho = openFileDialog1.FileName;
                            directoryPath = Path.GetDirectoryName(openFileDialog1.FileName);
                            files = (openFileDialog1.SafeFileNames);

                            foreach (string file in files)
                            {
                                filesAdionado.Add(file);
                                listBox1.Items.Add(file);

                            }

                            carregaLinhas();

                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
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
            if (textBox3.Text == "")
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
                            campos.Append("[" + Convert.ToString(colunasCreate[i]) + "] varchar (255), ");
                        }
                        campos.Append("[Sheet] varchar (255), [Arquivo] varchar(255) )");
                        sql = sql + campos.ToString();
                        //MessageBox.Show(sql);
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
            if (adicionado == false)
            {
                label2.Text = caminho;
                MyApp = new Excel.Application();
                excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + caminho + ";Extended Properties=Excel 12.0;";
                MyApp.Workbooks.Add("");
                MyApp.Workbooks.Add(caminho);

                if (ext == ".csv")
                {
                    // objeto para leitura de arquivo texto
                    StreamReader sr = new StreamReader(myStream);
                    string linha = "";
                    //   int lin = 0;

                    // lê a linha de títulos (nomes dos campos)
                    // enquanto não for fim do arquivo

                    //cabeçalho
                    linha = sr.ReadLine();
                    string[] cab = linha.Split(new String[] { "\",\"", ",\"", ",", ";" }, StringSplitOptions.RemoveEmptyEntries);

                    for (int i = 0; i < cab.Length; i++)
                    {
                        string coluna = cab[i].Trim('"');
                        colunas.Add(coluna);
                        colunasCreate.Add(coluna.Trim());
                        listBox2.Items.Add(coluna);
                    }

                }
                else
                {
                    if (MyApp.Workbooks[2].Worksheets[1].name != "Input" && MyApp.Workbooks[2].Worksheets[1].name != "Combined")
                    {
                        for (int k = 1; k <= MyApp.Workbooks[2].Worksheets[1].UsedRange.Columns.Count; k++)
                        {
                            if (Convert.ToString((MyApp.Workbooks[2].Worksheets[1].Cells[1, k].Value2)) != null && Convert.ToString(MyApp.Workbooks[2].Worksheets[1].Cells[1, k].Value2.ToString()) != "")
                            {
                                string coluna = Convert.ToString(MyApp.Workbooks[2].Worksheets[1].Cells[1, k].Value2);
                                colunas.Add(coluna);
                                colunasCreate.Add(coluna.Trim());
                                listBox2.Items.Add(coluna.Trim());
                            }
                        }
                    }
                    else
                    {
                        for (int k = 1; k <= MyApp.Workbooks[2].Worksheets[2].UsedRange.Columns.Count; k++)
                        {
                            string coluna = Convert.ToString(MyApp.Workbooks[2].Worksheets[2].Cells[1, k].Value2.ToString());
                            colunas.Add(coluna);
                            colunasCreate.Add(coluna.Trim());
                            listBox2.Items.Add(coluna.Trim());
                        }
                    }
                }
                //MyApp.Quit();
                //MyApp.Workbooks.Close();
                Marshal.ReleaseComObject(MyApp.Workbooks);
                Marshal.ReleaseComObject(MyApp);
                adicionado = true;
            }

        }

        private void button3_Click_1(object sender, EventArgs e)
        {



        }

        private void comboBoxConexao_Enter(object sender, EventArgs e)
        {
            //comboBoxConexao.Items.Clear();
            //string myServer = Environment.MachineName;

            //DataTable servers = SqlDataSourceEnumerator.Instance.GetDataSources();
            //for (int i = 0; i < servers.Rows.Count; i++)
            //{
            //    if (myServer == servers.Rows[i]["ServerName"].ToString())
            //    {
            //        if ((servers.Rows[i]["InstanceName"] as string) != null)
            //            comboBoxConexao.Items.Add(servers.Rows[i]["ServerName"] + "\\" + servers.Rows[i]["InstanceName"]);
            //        else
            //            comboBoxConexao.Items.Add(servers.Rows[i]["ServerName"]);
            //    }
            //}
            comboBoxConexao.Items.Clear();
            string myServer = Environment.MachineName;

            comboBoxConexao.Items.Add(myServer);
            string ServerName = Environment.MachineName;
            RegistryView registryView = Environment.Is64BitOperatingSystem ? RegistryView.Registry64 : RegistryView.Registry32;
            using (RegistryKey hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, registryView))
            {
                RegistryKey instanceKey = hklm.OpenSubKey(@"SOFTWARE\Microsoft\Microsoft SQL Server\Instance Names\SQL", false);
                if (instanceKey != null)
                {
                    foreach (var instanceName in instanceKey.GetValueNames())
                    {
                        comboBoxConexao.Items.Add(ServerName + "\\" + instanceName);
                    }
                }
            }
        }

        private void comboBoxBase_SelectedIndexChanged(object sender, EventArgs e)
        {
            baseDeDados = comboBoxBase.SelectedItem.ToString();

            using (var con = new SqlConnection("Data Source=" + conexao + ";Initial Catalog=" + baseDeDados + "; Integrated Security=True;"))
            {
                con.Open();
                DataTable t = con.GetSchema("Tables");

                foreach (DataRow ta in t.Rows)
                {

                    String datatableName = ta.Field<String>("table_name");

                }
            }
        }

        private void comboBoxTabela_SelectedIndexChanged(object sender, EventArgs e)
        {

            List<string> listacolumnas = new List<string>();
            using (SqlConnection connection = new SqlConnection("Data Source=" + conexao + " ;Initial Catalog=" + baseDeDados + " ;Integrated Security=True"))
            using (SqlCommand command = connection.CreateCommand())
            {
                command.CommandText = "select c.name from sys.columns c inner join sys.tables t on t.object_id = c.object_id where t.name = '" + tabela + "'";
                connection.Open();
                using (var reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        listacolumnas.Add(reader.GetString(0));
                        if (adicionado == false)
                        {
                            listBox2.Items.Add(reader.GetString(0));
                           
                        }

                    }
                    adicionado = true;
                }
            }
        }

        private void buttonLimpar_Click(object sender, EventArgs e)
        {
            listBox2.Items.Clear();
            adicionado = false;
        }
              
     }
}
      
