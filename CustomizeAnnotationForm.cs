using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Newtonsoft.Json;

namespace 课件帮PPT助手
{
    public partial class CustomizeAnnotationForm : Form
    {
        public event Action<string, string, string> AnnotationSaved;

        public CustomizeAnnotationForm()
        {
            InitializeComponent();
            this.Load += CustomizeAnnotationForm_Load;
            this.saveButton.Click += SaveButton_Click;
            this.cancelButton.Click += CancelButton_Click;
        }

        private void CustomizeAnnotationForm_Load(object sender, EventArgs e)
        {
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            string symbol = symbolTextBox.Text;
            string name = nameTextBox.Text;
            string position = bottomRadioButton.Checked ? "底部" :
                               startEndRadioButton.Checked ? "开头和末尾" : "末尾";

            if (!string.IsNullOrEmpty(symbol) && !string.IsNullOrEmpty(name))
            {
                SaveCustomSymbol(symbol); // Save the custom symbol
                AnnotationSaved?.Invoke(symbol, name, position);
                this.Close();
            }
            else
            {
                MessageBox.Show("请填写符号和名称。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void SaveCustomSymbol(string symbol)
        {
            string filePath = "custom_symbols.json";
            List<string> symbols = new List<string>();

            if (File.Exists(filePath))
            {
                string json = File.ReadAllText(filePath);
                symbols = JsonConvert.DeserializeObject<List<string>>(json);
            }

            if (!symbols.Contains(symbol))
            {
                symbols.Add(symbol);
            }

            string outputJson = JsonConvert.SerializeObject(symbols);
            File.WriteAllText(filePath, outputJson);
        }
    }
}
