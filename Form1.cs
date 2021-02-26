using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JsonConverter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            EventDispatcher.instance.UpdateInformation += Instance_UpdateInformation;
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
            Main.instance.StartConvert(openFileDialog1.FileName);
        }

        public void OnGetFilePath(string filePath)
        {
            var extension = Path.GetExtension(filePath);
            var info = string.Empty;
            if (extension == ".json")
                info = "Json -> Csv";
            else if (extension == ".csv")
                info = "Csv -> Json";

            var enabled = !string.IsNullOrEmpty(info);
            filePathText.Text = filePath;
            convertBtn.Enabled = enabled;
            UpdateInfoLabel(enabled ? info : "Error File");
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
