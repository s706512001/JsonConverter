﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace JsonConverter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            this.filePathLabel.Text = Message.INFO_INPUT_FILE;

            EventDispatcher.instance.UpdateInformation += Instance_UpdateInformation;
            EventDispatcher.instance.UpdateInformationWithFilePath += Instance_UpdateInformationWithFilePath;
            EventDispatcher.instance.ShowMessageBox += Instance_ShowMessageBox;
            EventDispatcher.instance.CommandLineExecuteStart += Instance_CommandLineExecuteStart;
            EventDispatcher.instance.CommandLineExecuteEnd += Instance_CommandLineExecuteEnd;
            this.filePathLabel.DragEnter += Form1_DragEnter;
            this.filePathLabel.DragDrop += Form1_DragDrop;

            Main.instance.CheckCommandLineInput();
        }

        private void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
        }

        private void Form1_DragDrop(object sender, DragEventArgs e)
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files.Length > 0)
                OnGetFilePath(files[0]);
        }

        private void borwseBtn_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = Directory.GetCurrentDirectory();
            openFileDialog1.Filter = "json files|*.json|csv files|*.csv|xlsx files|*.xlsx";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                OnGetFilePath(openFileDialog1.FileName);
        }

        private async void convertBtn_Click(object sender, EventArgs e)
        {
            this.convertBtn.Enabled = false;

            await Main.instance.StartConvertAsync(filePathLabel.Text, (FileType)convertToCombo.SelectedItem);

            this.convertBtn.Enabled = true;
        }

        public void OnGetFilePath(string filePath)
        {
            var fileType = Main.instance.ForPathFileType(filePath);

            UpdateInfoLabel(Message.INFO_SELECT_A_CONVERT_TO_TYPE);
            RefreshConvertToCombo(fileType);

            filePathLabel.Text = (FileType.none != fileType) ? filePath : Message.INFO_INPUT_FILE;
            convertBtn.Enabled = (FileType.none != fileType);
        }

        private void RefreshConvertToCombo(FileType fileType)
        {
            var convertToList = new List<FileType>();
            switch (fileType)
            {
                case FileType.json:
                    convertToTitle.Text = Message.INFO_CONVERT_TO_TITLE_JSON;
                    convertToList.Add(FileType.csv);
                    convertToList.Add(FileType.xlsx);
                    break;
                case FileType.csv:
                    convertToTitle.Text = Message.INFO_CONVERT_TO_TITLE_CSV;
                    convertToList.Add(FileType.json);
                    break;
                case FileType.xlsx:
                    convertToTitle.Text = Message.INFO_CONVERT_TO_TITLE_XLSX;
                    convertToList.Add(FileType.json);
                    break;
                default:
                    break;
            }
            convertToCombo.DataSource = convertToList.ToArray();
            convertToCombo.Enabled = true;
        }

        private void UpdateInfoLabel(string info)
        {
            infoLabel.Text = info;
        }

        #region Dispatch Event
        private void Instance_UpdateInformation(object sender, params object[] args)
        {
            var info = (string)args[0];
            UpdateInfoLabel(info);
        }

        private void Instance_UpdateInformationWithFilePath(object sender, params object[] args)
        {
            var filePath = (string)args[0];
            var info = (string)args[1];

            this.filePathLabel.Text = filePath;
            this.infoLabel.Text = info;
            this.convertBtn.Enabled = false;
        }

        private void Instance_ShowMessageBox(object sender, params object[] args)
        {
            var message = (string)args[0];

            MessageBox.Show(message);
        }

        private void Instance_CommandLineExecuteStart(object sender, params object[] args)
            => this.convertBtn.Enabled = false;

        private void Instance_CommandLineExecuteEnd(object sender, params object[] args)
            => this.convertBtn.Enabled = true;
        #endregion Dispatch Event
    }
}
