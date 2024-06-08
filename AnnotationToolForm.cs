using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Newtonsoft.Json;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace 课件帮PPT助手
{
    public partial class AnnotationToolForm : Form
    {
        public event Action<string, Color, bool, bool, Color, Color> AnnotationApplied;

        private const string CustomAnnotationsFile = "custom_annotations.json";
        private const string CustomAnnotationPrefix = "[自定义] ";

        public string SelectedAnnotationType { get; private set; }
        public Color AnnotationColor { get; private set; }
        public bool IsBold { get; private set; }
        public bool IsItalic { get; private set; }
        public bool IsHighlight { get; private set; }
        public Color HighlightColor { get; private set; }
        public Color TextColor { get; private set; }

        public AnnotationToolForm()
        {
            InitializeComponent();
            SetDefaultValues();
            this.TopMost = true;
            this.FormBorderStyle = FormBorderStyle.FixedToolWindow;

            confirmButton.BackColor = Color.FromArgb(47, 85, 151);
            confirmButton.ForeColor = Color.White;
            clearButton.BackColor = Color.FromArgb(47, 85, 151);
            clearButton.ForeColor = Color.White;

            InitializeContextMenu();
            LoadCustomAnnotations();
        }

        private void SaveCustomAnnotations()
        {
            var customAnnotations = new List<string>();
            foreach (var item in annotationTypeComboBox.Items)
            {
                if (item.ToString().StartsWith(CustomAnnotationPrefix))
                {
                    customAnnotations.Add(item.ToString());
                }
            }

            var json = JsonConvert.SerializeObject(customAnnotations);
            File.WriteAllText(CustomAnnotationsFile, json);
        }

        private void LoadCustomAnnotations()
        {
            if (File.Exists(CustomAnnotationsFile))
            {
                var json = File.ReadAllText(CustomAnnotationsFile);
                var customAnnotations = JsonConvert.DeserializeObject<List<string>>(json);
                foreach (var annotation in customAnnotations)
                {
                    annotationTypeComboBox.Items.Add(annotation);
                }
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            SaveCustomAnnotations();
        }

        private void InitializeContextMenu()
        {
            contextMenuStrip = new ContextMenuStrip();
            customizeAnnotationMenuItem = new ToolStripMenuItem("自定义标注");
            customizeAnnotationMenuItem.Click += CustomizeAnnotationMenuItem_Click;
            deleteAnnotationMenuItem = new ToolStripMenuItem("删除标注");
            deleteAnnotationMenuItem.Click += DeleteAnnotationMenuItem_Click;
            contextMenuStrip.Items.Add(customizeAnnotationMenuItem);
            contextMenuStrip.Items.Add(deleteAnnotationMenuItem);

            annotationTypeComboBox.ContextMenuStrip = contextMenuStrip;
            annotationTypeComboBox.DrawMode = DrawMode.OwnerDrawFixed;
            annotationTypeComboBox.DrawItem += AnnotationTypeComboBox_DrawItem;
            annotationTypeComboBox.MouseDown += AnnotationTypeComboBox_MouseDown;
        }

        private void AnnotationTypeComboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0)
                return;

            e.DrawBackground();
            e.Graphics.DrawString(annotationTypeComboBox.Items[e.Index].ToString(), e.Font, Brushes.Black, e.Bounds);
            e.DrawFocusRectangle();
        }

        private void AnnotationTypeComboBox_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                int index = annotationTypeComboBox.SelectedIndex;
                contextMenuStrip.Show(Cursor.Position);
            }
        }

        private void DeleteAnnotationMenuItem_Click(object sender, EventArgs e)
        {
            if (annotationTypeComboBox.SelectedIndex >= 0)
            {
                string selectedItem = annotationTypeComboBox.SelectedItem.ToString();
                if (selectedItem.StartsWith(CustomAnnotationPrefix))
                {
                    annotationTypeComboBox.Items.RemoveAt(annotationTypeComboBox.SelectedIndex);
                    SaveCustomAnnotations();
                    deleteCustomAnnotationButton.Enabled = false;
                }
                else
                {
                    MessageBox.Show("无法删除默认标注。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void CustomizeAnnotationMenuItem_Click(object sender, EventArgs e)
        {
            this.Hide();

            CustomizeAnnotationForm customizeForm = new CustomizeAnnotationForm();
            customizeForm.AnnotationSaved += (symbol, name, position) =>
            {
                annotationTypeComboBox.Items.Add($"{CustomAnnotationPrefix}{name} ({symbol}) - {position}");
            };

            customizeForm.FormClosed += (s, args) =>
            {
                this.Show();
            };

            customizeForm.ShowDialog();
        }

        public void SetDefaultValues()
        {
            annotationTypeComboBox.SelectedItem = "横线";
            annotationColorButton.BackColor = Color.Red;
            boldCheckBox.Checked = true;
            italicCheckBox.Checked = false;
            highlightCheckBox.Checked = false;
            highlightColorButton.BackColor = SystemColors.Control;
            textColorButton.BackColor = Color.Red;
        }

        private void AnnotationTypeComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedItem = annotationTypeComboBox.SelectedItem.ToString();
            deleteCustomAnnotationButton.Enabled = selectedItem.StartsWith(CustomAnnotationPrefix);
        }

        private void AnnotationColorButton_Click(object sender, EventArgs e)
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    AnnotationColor = colorDialog.Color;
                    annotationColorButton.BackColor = AnnotationColor;
                }
            }
        }

        private void HighlightColorButton_Click(object sender, EventArgs e)
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    HighlightColor = colorDialog.Color;
                    highlightColorButton.BackColor = HighlightColor;
                }
            }
        }

        private void TextColorButton_Click(object sender, EventArgs e)
        {
            using (ColorDialog colorDialog = new ColorDialog())
            {
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    TextColor = colorDialog.Color;
                    textColorButton.BackColor = TextColor;
                }
            }
        }

        private void HighlightCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            highlightColorButton.Enabled = highlightCheckBox.Checked;
            if (!highlightCheckBox.Checked)
            {
                highlightColorButton.BackColor = SystemColors.Control;
            }
        }

        private void ConfirmButton_Click(object sender, EventArgs e)
        {
            SelectedAnnotationType = annotationTypeComboBox.SelectedItem.ToString();
            AnnotationColor = annotationColorButton.BackColor;
            IsBold = boldCheckBox.Checked;
            IsItalic = italicCheckBox.Checked;
            IsHighlight = highlightCheckBox.Checked;
            HighlightColor = highlightCheckBox.Checked ? highlightColorButton.BackColor : Color.Empty;
            TextColor = textColorButton.BackColor;
            AnnotationApplied?.Invoke(SelectedAnnotationType, AnnotationColor, IsBold, IsItalic, HighlightColor, TextColor);
        }

        private void ClearButton_Click(object sender, EventArgs e)
        {
            PowerPoint.Application app = Globals.ThisAddIn.Application;
            PowerPoint.Selection sel = app.ActiveWindow.Selection;

            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
            {
                PowerPoint.TextRange textRange = sel.TextRange;

                // Clear text properties only for the selected text
                textRange.Font.Bold = Office.MsoTriState.msoFalse;
                textRange.Font.Italic = Office.MsoTriState.msoFalse;
                textRange.Font.Color.RGB = ColorTranslator.ToOle(Color.Black);

                // Remove annotations only for the selected text
                string text = textRange.Text;
                text = RemoveAnnotations(text);
                textRange.Text = text;

                // Remove shapes behind the selected text range if they overlap
                PowerPoint.Slide slide = (PowerPoint.Slide)app.ActiveWindow.View.Slide;
                List<PowerPoint.Shape> shapesToDelete = new List<PowerPoint.Shape>();

                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Name.StartsWith("Annotation_") && IsShapeOverlappingTextRange(shape, textRange))
                    {
                        shapesToDelete.Add(shape);
                    }
                }

                foreach (var shape in shapesToDelete)
                {
                    shape.Delete();
                }
            }
        }

        private bool IsShapeOverlappingTextRange(PowerPoint.Shape shape, PowerPoint.TextRange textRange)
        {
            // Check if the shape overlaps with the text range bounds
            float textLeft = textRange.BoundLeft;
            float textTop = textRange.BoundTop;
            float textWidth = textRange.BoundWidth;
            float textHeight = textRange.BoundHeight;

            float shapeLeft = shape.Left;
            float shapeTop = shape.Top;
            float shapeWidth = shape.Width;
            float shapeHeight = shape.Height;

            return !(shapeLeft + shapeWidth < textLeft ||
                     shapeLeft > textLeft + textWidth ||
                     shapeTop + shapeHeight < textTop ||
                     shapeTop > textTop + textHeight);
        }

        private string RemoveAnnotations(string text)
        {
            // Define both default and custom symbols to remove
            string[] defaultSymbols = new string[] { "{", "}", "※", "/", "//", "[", "]", "*", "(", ")", "▲", "○", "●" };

            // Load custom symbols from file
            string filePath = "custom_symbols.json";
            List<string> customSymbols = new List<string>();

            if (File.Exists(filePath))
            {
                string json = File.ReadAllText(filePath);
                customSymbols = JsonConvert.DeserializeObject<List<string>>(json);
            }

            // Combine default and custom symbols
            HashSet<string> symbolsToRemove = new HashSet<string>(defaultSymbols);
            foreach (string symbol in customSymbols)
            {
                symbolsToRemove.Add(symbol);
            }

            // Remove all symbols from text
            foreach (string symbol in symbolsToRemove)
            {
                text = text.Replace(symbol, "");
            }

            return text;
        }

        private void DeleteCustomAnnotationButton_Click(object sender, EventArgs e)
        {
            if (annotationTypeComboBox.SelectedIndex >= 0)
            {
                string selectedItem = annotationTypeComboBox.SelectedItem.ToString();
                if (selectedItem.StartsWith(CustomAnnotationPrefix))
                {
                    annotationTypeComboBox.Items.RemoveAt(annotationTypeComboBox.SelectedIndex);
                    SaveCustomAnnotations();
                    deleteCustomAnnotationButton.Enabled = false;
                }
            }
        }
    }
}
