namespace InsightForPPTAddIn
{
    partial class InsightPPTTool : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public InsightPPTTool()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.m_pInsightTab = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.m_pInsightTab.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // m_pInsightTab
            // 
            this.m_pInsightTab.Groups.Add(this.group1);
            this.m_pInsightTab.Label = "多视频导入插件";
            this.m_pInsightTab.Name = "m_pInsightTab";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Label = "批量插入视频";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button1.Image = global::InsightForPPTAddIn.Properties.Resources.vedio;
            this.button1.Label = "选择视频";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // InsightPPTTool
            // 
            this.Name = "InsightPPTTool";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.m_pInsightTab);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.InsightPPTTool_Load);
            this.m_pInsightTab.ResumeLayout(false);
            this.m_pInsightTab.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab m_pInsightTab;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal InsightPPTTool InsightPPTTool
        {
            get { return this.GetRibbon<InsightPPTTool>(); }
        }
    }
}
