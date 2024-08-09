using System.Collections.Generic;
using System.Windows;

namespace 课件帮PPT助手
{
    public partial class MemoryEditorWindow : Window
    {
        public Dictionary<string, Dictionary<string, string>> UpdatedFeedbackDict { get; private set; }

        public MemoryEditorWindow(Dictionary<string, Dictionary<string, string>> feedbackDict)
        {
            InitializeComponent(); // 初始化组件，加载XAML文件中的UI

            // 复制一份用户反馈字典，以便进行编辑
            UpdatedFeedbackDict = new Dictionary<string, Dictionary<string, string>>(feedbackDict);

            // 将记忆的字词和拼音添加到ListBox中
            PopulateListBox();
        }

        private void PopulateListBox()
        {
            listBox.Items.Clear();
            foreach (var key in UpdatedFeedbackDict.Keys)
            {
                foreach (var subKey in UpdatedFeedbackDict[key].Keys)
                {
                    string itemText = $"{key} ({subKey}) -> {UpdatedFeedbackDict[key][subKey]}";
                    listBox.Items.Add(itemText);
                }
            }
        }

        private void EditButton_Click(object sender, RoutedEventArgs e)
        {
            if (listBox.SelectedItem != null)
            {
                string selectedItem = listBox.SelectedItem.ToString();
                string[] parts = selectedItem.Split(new[] { " -> " }, System.StringSplitOptions.None);

                if (parts.Length == 2)
                {
                    string[] wordAndChar = parts[0].Split(new[] { " (" }, System.StringSplitOptions.None);
                    string key = wordAndChar[0];
                    string subKey = wordAndChar[1].TrimEnd(')');

                    string newPinyin = Microsoft.VisualBasic.Interaction.InputBox("编辑拼音", "编辑拼音", UpdatedFeedbackDict[key][subKey]);
                    if (!string.IsNullOrEmpty(newPinyin))
                    {
                        UpdatedFeedbackDict[key][subKey] = newPinyin;
                        PopulateListBox();
                    }
                }
            }
        }

        private void AddButton_Click(object sender, RoutedEventArgs e)
        {
            string newWord = Microsoft.VisualBasic.Interaction.InputBox("输入新的字词", "新增字词", "");
            if (!string.IsNullOrEmpty(newWord))
            {
                string newChar = Microsoft.VisualBasic.Interaction.InputBox("输入汉字", "新增汉字", "");
                if (!string.IsNullOrEmpty(newChar))
                {
                    string newPinyin = Microsoft.VisualBasic.Interaction.InputBox("输入拼音", "新增拼音", "");
                    if (!string.IsNullOrEmpty(newPinyin))
                    {
                        if (!UpdatedFeedbackDict.ContainsKey(newWord))
                        {
                            UpdatedFeedbackDict[newWord] = new Dictionary<string, string>();
                        }
                        UpdatedFeedbackDict[newWord][newChar] = newPinyin;
                        PopulateListBox();
                    }
                }
            }
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            if (listBox.SelectedItem != null)
            {
                string selectedItem = listBox.SelectedItem.ToString();
                string[] parts = selectedItem.Split(new[] { " -> " }, System.StringSplitOptions.None);

                if (parts.Length == 2)
                {
                    string[] wordAndChar = parts[0].Split(new[] { " (" }, System.StringSplitOptions.None);
                    string key = wordAndChar[0];
                    string subKey = wordAndChar[1].TrimEnd(')');

                    if (UpdatedFeedbackDict.ContainsKey(key) && UpdatedFeedbackDict[key].ContainsKey(subKey))
                    {
                        UpdatedFeedbackDict[key].Remove(subKey);
                        if (UpdatedFeedbackDict[key].Count == 0)
                        {
                            UpdatedFeedbackDict.Remove(key);
                        }
                        PopulateListBox();
                    }
                }
            }
        }

        private void ClearButton_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("确定要清空所有数据吗？此操作无法撤销。", "清空数据", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
            {
                UpdatedFeedbackDict.Clear();
                PopulateListBox();
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            // 将用户编辑的内容保存到 UpdatedFeedbackDict 中
            this.DialogResult = true; // 设置DialogResult为true，以便在调用ShowDialog时关闭窗口
            this.Close(); // 关闭窗口
        }
    }
}
