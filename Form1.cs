using System;
using System.IO;
using System.Windows.Forms;

namespace JsonConverter
{
    public partial class Form1 : Form
    {
        public string INITIAL_INFOMATION { private set; get; }

        public Form1()
        {
            InitializeComponent();

            INITIAL_INFOMATION = this.filePathLabel.Text;

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
            var info = string.Empty;
            switch (Main.instance.ForFileType(filePath))
            {
                case FileType.Json:
                    info = "Json -> Csv";
                    break;
                case FileType.Csv:
                    info = "Csv -> Json";
                    break;
            }

            var enabled = !string.IsNullOrEmpty(info);
            filePathLabel.Text = enabled ? filePath : INITIAL_INFOMATION;
            convertBtn.Enabled = enabled;
            UpdateInfoLabel(enabled ? info : Message.ERROR_INVALID_FILE_TYPE);
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
