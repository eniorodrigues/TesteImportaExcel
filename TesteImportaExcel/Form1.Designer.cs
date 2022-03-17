namespace TesteImportaExcel
{
    partial class testeBtn
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnImportarSQL = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.comboBoxBase = new System.Windows.Forms.ComboBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.comboBoxConexao = new System.Windows.Forms.ComboBox();
            this.button4 = new System.Windows.Forms.Button();
            this.listBox2 = new System.Windows.Forms.ListBox();
            this.label3 = new System.Windows.Forms.Label();
            this.buttonLimpar = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnImportarSQL
            // 
            this.btnImportarSQL.Location = new System.Drawing.Point(48, 592);
            this.btnImportarSQL.Margin = new System.Windows.Forms.Padding(10, 9, 10, 9);
            this.btnImportarSQL.Name = "btnImportarSQL";
            this.btnImportarSQL.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.btnImportarSQL.Size = new System.Drawing.Size(504, 65);
            this.btnImportarSQL.TabIndex = 0;
            this.btnImportarSQL.Text = "Importar para SQL";
            this.btnImportarSQL.UseVisualStyleBackColor = true;
            this.btnImportarSQL.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(38, 34);
            this.button2.Margin = new System.Windows.Forms.Padding(10, 9, 10, 9);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(440, 65);
            this.button2.TabIndex = 2;
            this.button2.Text = "Selecionar arquivos";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 37;
            this.listBox1.Location = new System.Drawing.Point(38, 194);
            this.listBox1.Margin = new System.Windows.Forms.Padding(10, 9, 10, 9);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(495, 374);
            this.listBox1.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(38, 148);
            this.label1.Margin = new System.Windows.Forms.Padding(10, 0, 10, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(151, 37);
            this.label1.TabIndex = 4;
            this.label1.Text = "Arquivos:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(48, 666);
            this.label2.Margin = new System.Windows.Forms.Padding(10, 0, 10, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(84, 37);
            this.label2.TabIndex = 6;
            this.label2.Text = "label";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(573, 148);
            this.label5.Margin = new System.Windows.Forms.Padding(10, 0, 10, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(235, 37);
            this.label5.TabIndex = 10;
            this.label5.Text = "Base de Dados";
            // 
            // comboBoxBase
            // 
            this.comboBoxBase.FormattingEnabled = true;
            this.comboBoxBase.Location = new System.Drawing.Point(583, 194);
            this.comboBoxBase.Margin = new System.Windows.Forms.Padding(10, 9, 10, 9);
            this.comboBoxBase.Name = "comboBoxBase";
            this.comboBoxBase.Size = new System.Drawing.Size(615, 45);
            this.comboBoxBase.TabIndex = 11;
            this.comboBoxBase.SelectedIndexChanged += new System.EventHandler(this.comboBoxBase_SelectedIndexChanged);
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(583, 518);
            this.textBox3.Margin = new System.Windows.Forms.Padding(10, 9, 10, 9);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(615, 44);
            this.textBox3.TabIndex = 12;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(573, 472);
            this.label6.Margin = new System.Windows.Forms.Padding(10, 0, 10, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(208, 37);
            this.label6.TabIndex = 13;
            this.label6.Text = "Nova Tabela:";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(573, 26);
            this.label7.Margin = new System.Windows.Forms.Padding(10, 0, 10, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(153, 37);
            this.label7.TabIndex = 15;
            this.label7.Text = "Conexão:";
            // 
            // comboBoxConexao
            // 
            this.comboBoxConexao.FormattingEnabled = true;
            this.comboBoxConexao.Location = new System.Drawing.Point(583, 71);
            this.comboBoxConexao.Margin = new System.Windows.Forms.Padding(10, 9, 10, 9);
            this.comboBoxConexao.Name = "comboBoxConexao";
            this.comboBoxConexao.Size = new System.Drawing.Size(615, 45);
            this.comboBoxConexao.TabIndex = 17;
            this.comboBoxConexao.Text = "Selecione a conexão";
            this.comboBoxConexao.SelectedIndexChanged += new System.EventHandler(this.comboBoxConexao_SelectedIndexChanged);
            this.comboBoxConexao.Enter += new System.EventHandler(this.comboBoxConexao_Enter);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(583, 592);
            this.button4.Margin = new System.Windows.Forms.Padding(10, 9, 10, 9);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(272, 65);
            this.button4.TabIndex = 18;
            this.button4.Text = "Criar Tabela";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // listBox2
            // 
            this.listBox2.FormattingEnabled = true;
            this.listBox2.ItemHeight = 37;
            this.listBox2.Location = new System.Drawing.Point(1254, 71);
            this.listBox2.Margin = new System.Windows.Forms.Padding(10, 9, 10, 9);
            this.listBox2.Name = "listBox2";
            this.listBox2.Size = new System.Drawing.Size(596, 485);
            this.listBox2.TabIndex = 20;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(1244, 26);
            this.label3.Margin = new System.Windows.Forms.Padding(10, 0, 10, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(135, 37);
            this.label3.TabIndex = 21;
            this.label3.Text = "Colunas";
            // 
            // buttonLimpar
            // 
            this.buttonLimpar.Location = new System.Drawing.Point(1254, 589);
            this.buttonLimpar.Margin = new System.Windows.Forms.Padding(10, 9, 10, 9);
            this.buttonLimpar.Name = "buttonLimpar";
            this.buttonLimpar.Size = new System.Drawing.Size(238, 65);
            this.buttonLimpar.TabIndex = 24;
            this.buttonLimpar.Text = "Limpar Colunas";
            this.buttonLimpar.UseVisualStyleBackColor = true;
            this.buttonLimpar.Click += new System.EventHandler(this.buttonLimpar_Click);
            // 
            // testeBtn
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(19F, 37F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1897, 734);
            this.Controls.Add(this.buttonLimpar);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.listBox2);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.comboBoxConexao);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.comboBoxBase);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.btnImportarSQL);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(10, 9, 10, 9);
            this.Name = "testeBtn";
            this.Text = "Importa Excel";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnImportarSQL;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox comboBoxBase;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox comboBoxConexao;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.ListBox listBox2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button buttonLimpar;
    }
}

