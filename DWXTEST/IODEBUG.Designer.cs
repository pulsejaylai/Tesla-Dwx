namespace DWXTEST
{
    partial class IODEBUG
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
            this.comdata = new System.Windows.Forms.ComboBox();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // comdata
            // 
            this.comdata.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comdata.FormattingEnabled = true;
            this.comdata.Location = new System.Drawing.Point(133, 29);
            this.comdata.Name = "comdata";
            this.comdata.Size = new System.Drawing.Size(236, 20);
            this.comdata.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(210, 65);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(74, 20);
            this.button1.TabIndex = 1;
            this.button1.Text = "OK";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // IODEBUG
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(521, 97);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.comdata);
            this.Name = "IODEBUG";
            this.Text = "IODEBUG";
            this.Load += new System.EventHandler(this.IODEBUG_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ComboBox comdata;
        private System.Windows.Forms.Button button1;
    }
}