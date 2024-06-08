using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace 课件帮PPT助手
{
    public partial class AnimationForm : Form
    {
        private List<PowerPoint.Shape> selectedShapes;
        private Dictionary<string, NumericUpDown> durationControls;

        public AnimationForm()
        {
            InitializeComponent();
            durationControls = new Dictionary<string, NumericUpDown>();
            textBox.KeyDown += TextBox_KeyDown;
            selectAllButton.Click += SelectAllButton_Click;
            animateButton.Click += AnimateButton_Click;
            adjustAnimationButton.Click += AdjustAnimationButton_Click;
            listBox.SelectedIndexChanged += ListBox_SelectedIndexChanged;
            upButton.Click += (s, ev) => AdjustAnimationDirection(listBox, PowerPoint.MsoAnimDirection.msoAnimDirectionBottom);
            downButton.Click += (s, ev) => AdjustAnimationDirection(listBox, PowerPoint.MsoAnimDirection.msoAnimDirectionTop);
            leftButton.Click += (s, ev) => AdjustAnimationDirection(listBox, PowerPoint.MsoAnimDirection.msoAnimDirectionRight);
            rightButton.Click += (s, ev) => AdjustAnimationDirection(listBox, PowerPoint.MsoAnimDirection.msoAnimDirectionLeft);
        }

        private void TextBox_KeyDown(object sender, KeyEventArgs ev)
        {
            if (ev.KeyCode == Keys.Enter)
            {
                PowerPoint.Application pptApplication = Globals.ThisAddIn.Application;
                PowerPoint.DocumentWindow activeWindow = pptApplication.ActiveWindow;
                PowerPoint.Selection selection = activeWindow.Selection;

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    TextBox textBox = sender as TextBox;
                    string prefix = textBox.Text;

                    if (!string.IsNullOrEmpty(prefix))
                    {
                        int counter = 1;
                        selectedShapes = new List<PowerPoint.Shape>();
                        foreach (PowerPoint.Shape shape in selection.ShapeRange)
                        {
                            shape.Name = $"{prefix}-{counter}";
                            selectedShapes.Add(shape);
                            counter++;
                        }

                        activeWindow.View.GotoSlide(activeWindow.View.Slide.SlideIndex);
                    }
                    else
                    {
                        MessageBox.Show("命名前缀不能为空。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("请选择一个或多个对象。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }

        private void SelectAllButton_Click(object sender, EventArgs ev)
        {
            if (selectedShapes == null || !selectedShapes.Any())
            {
                MessageBox.Show("请先完成第一步的批量命名。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Slide slide = pptApp.ActiveWindow.View.Slide;

            pptApp.ActiveWindow.Selection.Unselect();
            foreach (var shape in selectedShapes)
            {
                shape.Select(Office.MsoTriState.msoFalse);
            }
        }

        private void AnimateButton_Click(object sender, EventArgs ev)
        {
            if (selectedShapes == null || !selectedShapes.Any())
            {
                MessageBox.Show("请先完成第一步的批量命名。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Slide slide = pptApp.ActiveWindow.View.Slide;

            PowerPoint.TimeLine timeLine = slide.TimeLine;
            bool isFirstEffect = true;
            foreach (PowerPoint.Shape shape in selectedShapes)
            {
                PowerPoint.Effect effect = timeLine.MainSequence.AddEffect(
                    shape,
                    PowerPoint.MsoAnimEffect.msoAnimEffectWipe,
                    PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone,
                    isFirstEffect ? PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick : PowerPoint.MsoAnimTriggerType.msoAnimTriggerAfterPrevious
                );

                if (shape.Width > shape.Height)
                {
                    effect.EffectParameters.Direction = PowerPoint.MsoAnimDirection.msoAnimDirectionLeft;
                }
                else
                {
                    effect.EffectParameters.Direction = PowerPoint.MsoAnimDirection.msoAnimDirectionUp;
                }

                if (durationControls.TryGetValue(shape.Name, out NumericUpDown durationControl))
                {
                    effect.Timing.Duration = (float)durationControl.Value;
                }

                isFirstEffect = false;
            }
        }

        private void AdjustAnimationButton_Click(object sender, EventArgs ev)
        {
            if (selectedShapes != null)
            {
                listBox.Items.Clear();
                string currentPrefix = selectedShapes[0].Name.Split('-')[0];

                foreach (var shape in selectedShapes)
                {
                    if (shape.Name.StartsWith(currentPrefix))
                    {
                        listBox.Items.Add(shape.Name);
                    }
                }
            }
            adjustPanel.Visible = !adjustPanel.Visible;
        }

        private void AdjustAnimationDirection(ListBox listBox, PowerPoint.MsoAnimDirection direction)
        {
            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Slide slide = pptApp.ActiveWindow.View.Slide;

            foreach (string shapeName in listBox.SelectedItems)
            {
                var shape = slide.Shapes[shapeName];
                var effect = slide.TimeLine.MainSequence.Cast<PowerPoint.Effect>().FirstOrDefault(e => e.Shape.Name == shapeName);
                if (effect != null)
                {
                    effect.EffectParameters.Direction = direction;
                }
            }
        }

        private void DurationControl_ValueChanged(object sender, EventArgs ev)
        {
            NumericUpDown durationControl = sender as NumericUpDown;
            string shapeName = durationControl.Tag as string;

            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Slide slide = pptApp.ActiveWindow.View.Slide;

            var effect = slide.TimeLine.MainSequence.Cast<PowerPoint.Effect>().FirstOrDefault(e => e.Shape.Name == shapeName);
            if (effect != null)
            {
                effect.Timing.Duration = (float)durationControl.Value;
            }
        }

        private void MultiDurationControl_ValueChanged(object sender, EventArgs ev)
        {
            NumericUpDown multiDurationControl = sender as NumericUpDown;
            float newDuration = (float)multiDurationControl.Value;

            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            PowerPoint.Slide slide = pptApp.ActiveWindow.View.Slide;

            foreach (string shapeName in listBox.SelectedItems)
            {
                var effect = slide.TimeLine.MainSequence.Cast<PowerPoint.Effect>().FirstOrDefault(e => e.Shape.Name == shapeName);
                if (effect != null)
                {
                    effect.Timing.Duration = newDuration;
                }

                if (durationControls.TryGetValue(shapeName, out NumericUpDown durationControl))
                {
                    durationControl.Value = (decimal)newDuration;
                }
            }
        }

        private void ListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            adjustPanel.Controls.OfType<NumericUpDown>().ToList().ForEach(control => adjustPanel.Controls.Remove(control));

            PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            pptApp.ActiveWindow.Selection.Unselect();

            if (listBox.SelectedItems.Count > 1)
            {
                foreach (var selectedItem in listBox.SelectedItems)
                {
                    string shapeName = selectedItem.ToString();
                    var shape = pptApp.ActiveWindow.View.Slide.Shapes[shapeName];
                    shape.Select(Office.MsoTriState.msoFalse);
                }

                NumericUpDown multiDurationControl = new NumericUpDown
                {
                    Minimum = 0.1m,
                    Maximum = 10m,
                    DecimalPlaces = 2,
                    Increment = 0.1m,
                    Value = 0.50m
                };
                multiDurationControl.ValueChanged += MultiDurationControl_ValueChanged;
                multiDurationControl.Location = new System.Drawing.Point(270, 170);
                adjustPanel.Controls.Add(multiDurationControl);
            }
            else
            {
                foreach (var selectedItem in listBox.SelectedItems)
                {
                    string shapeName = selectedItem.ToString();
                    var shape = pptApp.ActiveWindow.View.Slide.Shapes[shapeName];
                    shape.Select(Office.MsoTriState.msoFalse);

                    if (!durationControls.TryGetValue(shapeName, out NumericUpDown durationControl))
                    {
                        durationControl = new NumericUpDown
                        {
                            Minimum = 0.1m,
                            Maximum = 10m,
                            DecimalPlaces = 2,
                            Increment = 0.1m,
                            Value = 0.50m,
                            Tag = shapeName
                        };
                        durationControl.ValueChanged += DurationControl_ValueChanged;
                        durationControls[shapeName] = durationControl;
                    }

                    durationControl.Location = new System.Drawing.Point(270, 170);
                    adjustPanel.Controls.Add(durationControl);
                }
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
