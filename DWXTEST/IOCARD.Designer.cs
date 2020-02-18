namespace DWXTEST
{
    partial class IOCARD
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
            this.comtype = new System.Windows.Forms.ComboBox();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // comtype
            // 
            this.comtype.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comtype.FormattingEnabled = true;
            this.comtype.Location = new System.Drawing.Point(116, 22);
            this.comtype.Name = "comtype";
            this.comtype.Size = new System.Drawing.Size(208, 20);
            this.comtype.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(173, 74);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(76, 22);
            this.button1.TabIndex = 1;
            this.button1.Text = "OK";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // IOCARD
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(483, 98);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.comtype);
            this.Name = "IOCARD";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "IOCARD";
            this.Load += new System.EventHandler(this.IOCARD_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ComboBox comtype;
        private System.Windows.Forms.Button button1;
    }
}