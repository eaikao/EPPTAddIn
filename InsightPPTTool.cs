using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using System.IO;

namespace InsightForPPTAddIn
{
    public partial class InsightPPTTool
    {
        private void InsightPPTTool_Load(object sender, RibbonUIEventArgs e) {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e) {
            //InsertForm insertForm = new InsertForm();
            //insertForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            //insertForm.ShowDialog();

            //获取当前ppt中所有的幻灯片
            Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;
            var activeSlide = (Slide)Globals.ThisAddIn.Application.ActiveWindow.View.Slide;
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Multiselect = true;
            dlg.Filter = "视频文件|*.avi;*.rmvb;*.rm;*.asf;*.asx;*.wmx;*.mov;*.flv;*.swf;*.divx;*.mpg;*.mpeg;*.mpe;*.wmv;*.mp4;*.mkv;*.vob";
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                foreach (string file in dlg.FileNames)
                {
                    Slide curSlide = slides.Add(activeSlide.SlideIndex + 1, PpSlideLayout.ppLayoutBlank);
                    Shape meidaShap = curSlide.Shapes.AddMediaObject2(file);
                    meidaShap.Top = meidaShap.Height / 4;
                    meidaShap.Left = meidaShap.Width / 4;
                    meidaShap.Width = meidaShap.Width / 2;
                    if (meidaShap != null && meidaShap.MediaType == PpMediaType.ppMediaTypeMovie) {
                        meidaShap.AnimationSettings.PlaySettings.LoopUntilStopped = Microsoft.Office.Core.MsoTriState.msoTrue;
                        meidaShap.AnimationSettings.PlaySettings.PlayOnEntry = Microsoft.Office.Core.MsoTriState.msoTrue;
                        var effect = curSlide.TimeLine.MainSequence[1];
                        if (effect.Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerOnPageClick) {
                            effect.Timing.TriggerType = MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                        }
                    }
                    activeSlide = curSlide;
                }

            }
        }
    }
}
