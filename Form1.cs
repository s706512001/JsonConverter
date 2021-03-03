using System;
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
            this.filePathLabel.DragEnter += Form1_DragEnter;
            this.filePathLabel.DragDrop += Form1_DragDrop;
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
            openFileDialog1.Filter = "json files|*.json|csv files|*.csv";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                OnGetFilePath(openFileDialog1.FileName);
        }

        private void convertBtn_Click(object sender, EventArgs e)
        {
            Main.instance.StartConvert(filePathLabel.Text);
        }

        public void OnGetFilePath(string filePath)
        {
            var fileType = Main.instance.ForFileType(filePath);
            switch (fileType)
            {
                case FileType.Json:
                    UpdateInfoLabel(Message.INFO_JSON_TO_CSV);
                    break;
                case FileType.Csv:
                    UpdateInfoLabel(Message.INFO_CSV_TO_JSON);
                    break;
                default:
                    UpdateInfoLabel(Message.ERROR_INVALID_FILE_TYPE);
                    break;
            }

            filePathLabel.Text = (FileType.None != fileType) ? filePath : Message.INFO_INPUT_FILE;
            convertBtn.Enabled = (FileType.None != fileType);
        }

        private void UpdateInfoLabel(string info)
        {
            infoLabel.Text = info;
        }

        private void Instance_UpdateInformation(object sender, params object[] args)
        {
            var info = (string)args[0];
            UpdateInfoLabel(info);
        }
    }
}
