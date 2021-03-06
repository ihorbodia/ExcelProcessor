﻿using ExcelProcessor.Logic;
using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelProcessor
{
    
	public partial class ExcelProcessorMainGui : Form
	{
        string chosenPath = string.Empty;
		public ExcelProcessorMainGui()
		{
			InitializeComponent();
			StatusLabelText.Text = "Choose folder";
            this.Text = "Excel processor: v1.1";
            
        }

        private void ChooseFirstFolderButton_Click(object sender, System.EventArgs e)
        {
            chosenPath = Common.FilesHelper.SelectFolder();
            if (chosenPath != string.Empty)
            {
                StatusLabelText.Text = "Start process";
                ChoosenPathLabel.Text = chosenPath;
            }
        }

        private void ProcessFilesButton_Click(object sender, System.EventArgs e)
        {
            StatusLabelText.Text = "Processing";
            try
            {
                new Task(() =>
                {
                    Thread t = new Thread(RunProgram);
                    t.Start();
                    t.Join();
                    StatusLabelText.BeginInvoke((MethodInvoker)delegate () { StatusLabelText.Text = "Finish"; });
                    Console.WriteLine("Finish");
                }).Start();
            }
            catch (Exception)
            {
                StatusLabelText.Text = "Something wrong";
            }
            
        }
        private void RunProgram()
        {
            FilesProcessor fp;
            fp = new FilesProcessor(chosenPath);
            fp.InitCountryDocFiles();
            fp.Run();
            fp.SaveCountryDocFiles();
        }
    }
}
