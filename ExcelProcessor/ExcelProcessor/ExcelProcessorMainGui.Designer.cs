namespace ExcelProcessor
{
	partial class ExcelProcessorMainGui
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
            this.ChooseFirstFolderButton = new System.Windows.Forms.Button();
            this.ProcessFilesButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.StatusLabelText = new System.Windows.Forms.Label();
            this.FolderChosenPath = new System.Windows.Forms.Label();
            this.ChoosenPathLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // ChooseFirstFolderButton
            // 
            this.ChooseFirstFolderButton.Location = new System.Drawing.Point(12, 226);
            this.ChooseFirstFolderButton.Name = "ChooseFirstFolderButton";
            this.ChooseFirstFolderButton.Size = new System.Drawing.Size(107, 23);
            this.ChooseFirstFolderButton.TabIndex = 0;
            this.ChooseFirstFolderButton.Text = "Choose folder";
            this.ChooseFirstFolderButton.UseVisualStyleBackColor = true;
            this.ChooseFirstFolderButton.Click += new System.EventHandler(this.ChooseFirstFolderButton_Click);
            // 
            // ProcessFilesButton
            // 
            this.ProcessFilesButton.Location = new System.Drawing.Point(405, 226);
            this.ProcessFilesButton.Name = "ProcessFilesButton";
            this.ProcessFilesButton.Size = new System.Drawing.Size(107, 23);
            this.ProcessFilesButton.TabIndex = 2;
            this.ProcessFilesButton.Text = "Process files";
            this.ProcessFilesButton.UseVisualStyleBackColor = true;
            this.ProcessFilesButton.Click += new System.EventHandler(this.ProcessFilesButton_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 162);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(40, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Status:";
            // 
            // StatusLabelText
            // 
            this.StatusLabelText.AutoSize = true;
            this.StatusLabelText.Location = new System.Drawing.Point(13, 187);
            this.StatusLabelText.Name = "StatusLabelText";
            this.StatusLabelText.Size = new System.Drawing.Size(0, 13);
            this.StatusLabelText.TabIndex = 4;
            // 
            // FolderChosenPath
            // 
            this.FolderChosenPath.AutoSize = true;
            this.FolderChosenPath.Location = new System.Drawing.Point(13, 9);
            this.FolderChosenPath.Name = "FolderChosenPath";
            this.FolderChosenPath.Size = new System.Drawing.Size(63, 13);
            this.FolderChosenPath.TabIndex = 5;
            this.FolderChosenPath.Text = "Folder path:";
            // 
            // ChoosenPathLabel
            // 
            this.ChoosenPathLabel.AutoSize = true;
            this.ChoosenPathLabel.Location = new System.Drawing.Point(13, 33);
            this.ChoosenPathLabel.Name = "ChoosenPathLabel";
            this.ChoosenPathLabel.Size = new System.Drawing.Size(0, 13);
            this.ChoosenPathLabel.TabIndex = 6;
            this.ChoosenPathLabel.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // ExcelProcessorMainGui
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(524, 261);
            this.Controls.Add(this.ChoosenPathLabel);
            this.Controls.Add(this.FolderChosenPath);
            this.Controls.Add(this.StatusLabelText);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ProcessFilesButton);
            this.Controls.Add(this.ChooseFirstFolderButton);
            this.Name = "ExcelProcessorMainGui";
            this.Text = "Excel processor";
            this.ResumeLayout(false);
            this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Button ChooseFirstFolderButton;
		private System.Windows.Forms.Button ProcessFilesButton;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label StatusLabelText;
		private System.Windows.Forms.Label FolderChosenPath;
        private System.Windows.Forms.Label ChoosenPathLabel;
    }
}

