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
        string usuario;
        string senha;

        private void button1_Click(object sender, EventArgs e)
        {

            label2.Text = "Importando sheet";

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
                        
                     // Excel.Worksheet ws = app.Workbooks[i].Worksheets[j];
                     label2.Text = "Importando arquivo " + element + " da sheet " + app.Workbooks[i].Worksheets[j].name;
                    string myString = (app.Workbooks[i].Worksheets[j].Cells[1, 1]).Value2.ToString();
                    
                     //ws.Copy(app.Workbooks[1].Worksheets[1]);

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
                            string sqlConnectionString = "Data Source=BRCAENRODRIGUES\\msSQLEXPRESS;Initial Catalog=TESTE_2016;Integrated Security=True";

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

        }
    }
}
