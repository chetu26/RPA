namespace EditWordFiles
{
    partial class WordFileModificationForm
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
            this.browsePdfBtn = new System.Windows.Forms.Button();
            this.pdfTextBox = new System.Windows.Forms.TextBox();
            this.button3 = new System.Windows.Forms.Button();
            this.docTextBox = new System.Windows.Forms.TextBox();
            this.modifyBtn = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // browsePdfBtn
            // 
            this.browsePdfBtn.Location = new System.Drawing.Point(35, 33);
            this.browsePdfBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.browsePdfBtn.Name = "browsePdfBtn";
            this.browsePdfBtn.Size = new System.Drawing.Size(165, 33);
            this.browsePdfBtn.TabIndex = 0;
            this.browsePdfBtn.Text = "Select PDF Folder";
            this.browsePdfBtn.UseVisualStyleBackColor = true;
            this.browsePdfBtn.Click += new System.EventHandler(this.BrowsePdfBtn_Click);
            // 
            // pdfTextBox
            // 
            this.pdfTextBox.Location = new System.Drawing.Point(232, 33);
            this.pdfTextBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.pdfTextBox.Multiline = true;
            this.pdfTextBox.Name = "pdfTextBox";
            this.pdfTextBox.Size = new System.Drawing.Size(439, 32);
            this.pdfTextBox.TabIndex = 1;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(35, 106);
            this.button3.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(165, 33);
            this.button3.TabIndex = 3;
            this.button3.Text = "Select Template File";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.BrowseDocBtn_Click);
            // 
            // docTextBox
            // 
            this.docTextBox.Location = new System.Drawing.Point(232, 106);
            this.docTextBox.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.docTextBox.Multiline = true;
            this.docTextBox.Name = "docTextBox";
            this.docTextBox.Size = new System.Drawing.Size(439, 32);
            this.docTextBox.TabIndex = 4;
            // 
            // modifyBtn
            // 
            this.modifyBtn.Location = new System.Drawing.Point(507, 171);
            this.modifyBtn.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.modifyBtn.Name = "modifyBtn";
            this.modifyBtn.Size = new System.Drawing.Size(165, 33);
            this.modifyBtn.TabIndex = 5;
            this.modifyBtn.Text = "Transform";
            this.modifyBtn.UseVisualStyleBackColor = true;
            this.modifyBtn.Click += new System.EventHandler(this.ModifyBtn_Click);
            // 
            // WordFileModificationForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(764, 266);
            this.Controls.Add(this.modifyBtn);
            this.Controls.Add(this.docTextBox);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.pdfTextBox);
            this.Controls.Add(this.browsePdfBtn);
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "WordFileModificationForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "PDF TO Word Processor";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button browsePdfBtn;
        private System.Windows.Forms.TextBox pdfTextBox;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.TextBox docTextBox;
        private System.Windows.Forms.Button modifyBtn;
    }
}

