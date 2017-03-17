namespace ExcelReader
{
    partial class Configuration
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
            this.label3 = new System.Windows.Forms.Label();
            this.startrow = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.colcount = new System.Windows.Forms.TextBox();
            this.OK = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(71, 80);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(66, 13);
            this.label3.TabIndex = 10;
            this.label3.Text = "Starting row:";
            // 
            // startrow
            // 
            this.startrow.Location = new System.Drawing.Point(171, 80);
            this.startrow.Name = "startrow";
            this.startrow.Size = new System.Drawing.Size(76, 20);
            this.startrow.TabIndex = 9;
            this.startrow.Text = "1";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(36, 38);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(101, 13);
            this.label2.TabIndex = 8;
            this.label2.Text = "Number of columns:";
            // 
            // colcount
            // 
            this.colcount.Location = new System.Drawing.Point(171, 38);
            this.colcount.Name = "colcount";
            this.colcount.Size = new System.Drawing.Size(76, 20);
            this.colcount.TabIndex = 7;
            this.colcount.Text = "10";
            // 
            // OK
            // 
            this.OK.Location = new System.Drawing.Point(172, 203);
            this.OK.Name = "OK";
            this.OK.Size = new System.Drawing.Size(75, 23);
            this.OK.TabIndex = 11;
            this.OK.Text = "OK";
            this.OK.UseVisualStyleBackColor = true;
            this.OK.Click += new System.EventHandler(this.OK_Click);
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(26, 128);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(246, 70);
            this.label1.TabIndex = 12;
            this.label1.Text = "Use these fields, when re-populating an existing Excel template, in order to re-u" +
    "se cell values. \r\n\r\nTypically, the \'Number of columns\' should be set to a large " +
    "value, and \'Starting row\' at 1.";
            this.label1.UseWaitCursor = true;
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.Cursor = System.Windows.Forms.Cursors.WaitCursor;
            this.label4.ForeColor = System.Drawing.Color.Red;
            this.label4.Location = new System.Drawing.Point(26, 110);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(45, 18);
            this.label4.TabIndex = 13;
            this.label4.Text = "NOTE:";
            this.label4.UseWaitCursor = true;
            // 
            // Configuration
            // 
            this.AcceptButton = this.OK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.OK);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.startrow);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.colcount);
            this.KeyPreview = true;
            this.Name = "Configuration";
            this.Text = "Configuration";
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Configuration_KeyDown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox startrow;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox colcount;
        private System.Windows.Forms.Button OK;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
    }
}