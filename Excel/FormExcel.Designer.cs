namespace Excel
{
    partial class FormExcel
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormExcel));
            this.labelHeader = new System.Windows.Forms.Label();
            this.btnSair = new System.Windows.Forms.Button();
            this.btnProcessar = new System.Windows.Forms.Button();
            this.pgbProcesso = new System.Windows.Forms.ProgressBar();
            this.label1 = new System.Windows.Forms.Label();
            this.btnCaminhoArq = new System.Windows.Forms.Button();
            this.lblCaminho = new System.Windows.Forms.Label();
            this.ofdCaminho = new System.Windows.Forms.OpenFileDialog();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // labelHeader
            // 
            this.labelHeader.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.labelHeader.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            resources.ApplyResources(this.labelHeader, "labelHeader");
            this.labelHeader.Name = "labelHeader";
            // 
            // btnSair
            // 
            resources.ApplyResources(this.btnSair, "btnSair");
            this.btnSair.Name = "btnSair";
            this.btnSair.UseVisualStyleBackColor = true;
            this.btnSair.Click += new System.EventHandler(this.btnSair_Click);
            // 
            // btnProcessar
            // 
            resources.ApplyResources(this.btnProcessar, "btnProcessar");
            this.btnProcessar.Name = "btnProcessar";
            this.btnProcessar.UseVisualStyleBackColor = true;
            this.btnProcessar.Click += new System.EventHandler(this.btnProcessar_Click);
            // 
            // pgbProcesso
            // 
            resources.ApplyResources(this.pgbProcesso, "pgbProcesso");
            this.pgbProcesso.Name = "pgbProcesso";
            // 
            // label1
            // 
            resources.ApplyResources(this.label1, "label1");
            this.label1.Name = "label1";
            // 
            // btnCaminhoArq
            // 
            resources.ApplyResources(this.btnCaminhoArq, "btnCaminhoArq");
            this.btnCaminhoArq.BackgroundImage = global::Excel.Properties.Resources.pasta;
            this.btnCaminhoArq.Name = "btnCaminhoArq";
            this.btnCaminhoArq.UseVisualStyleBackColor = true;
            this.btnCaminhoArq.Click += new System.EventHandler(this.btnCaminhoArq_Click);
            // 
            // lblCaminho
            // 
            resources.ApplyResources(this.lblCaminho, "lblCaminho");
            this.lblCaminho.ForeColor = System.Drawing.Color.Red;
            this.lblCaminho.Name = "lblCaminho";
            // 
            // ofdCaminho
            // 
            this.ofdCaminho.FileName = "openFileDialog1";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lblCaminho);
            resources.ApplyResources(this.groupBox1, "groupBox1");
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.TabStop = false;
            // 
            // FormExcel
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnCaminhoArq);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.pgbProcesso);
            this.Controls.Add(this.btnProcessar);
            this.Controls.Add(this.btnSair);
            this.Controls.Add(this.labelHeader);
            this.KeyPreview = true;
            this.MaximizeBox = false;
            this.Name = "FormExcel";
            this.Load += new System.EventHandler(this.FormExcel_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelHeader;
        private System.Windows.Forms.Button btnSair;
        private System.Windows.Forms.Button btnProcessar;
        private System.Windows.Forms.ProgressBar pgbProcesso;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnCaminhoArq;
        private System.Windows.Forms.Label lblCaminho;
        private System.Windows.Forms.OpenFileDialog ofdCaminho;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}

