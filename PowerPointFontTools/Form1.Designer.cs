namespace PowerPointFontTools
{
    partial class MainApp
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.labelDropZone = new System.Windows.Forms.Label();
            this.listBoxMissingFonts = new System.Windows.Forms.ListBox();
            this.listBoxInstalledFonts = new System.Windows.Forms.ListBox();
            this.btnExportMissingList = new System.Windows.Forms.Button();
            this.btnExportInstalledFonts = new System.Windows.Forms.Button();
            this.labelMissing = new System.Windows.Forms.Label();
            this.labelInstalled = new System.Windows.Forms.Label();
            this.labelStatus = new System.Windows.Forms.Label();
            this.btnReloadFonts = new System.Windows.Forms.Button();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.labelProgress = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // labelDropZone
            // 
            this.labelDropZone.AllowDrop = true;
            this.labelDropZone.BackColor = System.Drawing.Color.LightGray;
            this.labelDropZone.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.labelDropZone.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelDropZone.Location = new System.Drawing.Point(24, 24);
            this.labelDropZone.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.labelDropZone.Name = "labelDropZone";
            this.labelDropZone.Size = new System.Drawing.Size(1358, 158);
            this.labelDropZone.TabIndex = 0;
            this.labelDropZone.Text = "拖入 .pptx 文件到此处";
            this.labelDropZone.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.labelDropZone.DragDrop += new System.Windows.Forms.DragEventHandler(this.labelDropZone_DragDrop);
            this.labelDropZone.DragEnter += new System.Windows.Forms.DragEventHandler(this.labelDropZone_DragEnter);
            // 
            // listBoxMissingFonts
            // 
            this.listBoxMissingFonts.FormattingEnabled = true;
            this.listBoxMissingFonts.ItemHeight = 24;
            this.listBoxMissingFonts.Location = new System.Drawing.Point(24, 250);
            this.listBoxMissingFonts.Margin = new System.Windows.Forms.Padding(6);
            this.listBoxMissingFonts.Name = "listBoxMissingFonts";
            this.listBoxMissingFonts.Size = new System.Drawing.Size(756, 508);
            this.listBoxMissingFonts.TabIndex = 1;
            this.listBoxMissingFonts.DoubleClick += new System.EventHandler(this.listBoxMissingFonts_DoubleClick);
            // 
            // listBoxInstalledFonts
            // 
            this.listBoxInstalledFonts.FormattingEnabled = true;
            this.listBoxInstalledFonts.ItemHeight = 24;
            this.listBoxInstalledFonts.Location = new System.Drawing.Point(816, 250);
            this.listBoxInstalledFonts.Margin = new System.Windows.Forms.Padding(6);
            this.listBoxInstalledFonts.Name = "listBoxInstalledFonts";
            this.listBoxInstalledFonts.Size = new System.Drawing.Size(756, 508);
            this.listBoxInstalledFonts.TabIndex = 2;
            // 
            // btnExportMissingList
            // 
            this.btnExportMissingList.Enabled = false;
            this.btnExportMissingList.Location = new System.Drawing.Point(24, 780);
            this.btnExportMissingList.Margin = new System.Windows.Forms.Padding(6);
            this.btnExportMissingList.Name = "btnExportMissingList";
            this.btnExportMissingList.Size = new System.Drawing.Size(760, 60);
            this.btnExportMissingList.TabIndex = 3;
            this.btnExportMissingList.Text = "导出缺失字体名单 (TXT)";
            this.btnExportMissingList.UseVisualStyleBackColor = true;
            this.btnExportMissingList.Click += new System.EventHandler(this.btnExportMissingList_Click);
            // 
            // btnExportInstalledFonts
            // 
            this.btnExportInstalledFonts.Enabled = false;
            this.btnExportInstalledFonts.Location = new System.Drawing.Point(816, 780);
            this.btnExportInstalledFonts.Margin = new System.Windows.Forms.Padding(6);
            this.btnExportInstalledFonts.Name = "btnExportInstalledFonts";
            this.btnExportInstalledFonts.Size = new System.Drawing.Size(760, 60);
            this.btnExportInstalledFonts.TabIndex = 4;
            this.btnExportInstalledFonts.Text = "导出已安装字体 (ZIP)";
            this.btnExportInstalledFonts.UseVisualStyleBackColor = true;
            this.btnExportInstalledFonts.Click += new System.EventHandler(this.btnExportInstalledFonts_Click);
            // 
            // labelMissing
            // 
            this.labelMissing.AutoSize = true;
            this.labelMissing.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelMissing.ForeColor = System.Drawing.Color.Red;
            this.labelMissing.Location = new System.Drawing.Point(24, 210);
            this.labelMissing.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.labelMissing.Name = "labelMissing";
            this.labelMissing.Size = new System.Drawing.Size(275, 29);
            this.labelMissing.TabIndex = 5;
            this.labelMissing.Text = "缺失字体 (0) - 双击搜索";
            // 
            // labelInstalled
            // 
            this.labelInstalled.AutoSize = true;
            this.labelInstalled.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelInstalled.ForeColor = System.Drawing.Color.Green;
            this.labelInstalled.Location = new System.Drawing.Point(816, 210);
            this.labelInstalled.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.labelInstalled.Name = "labelInstalled";
            this.labelInstalled.Size = new System.Drawing.Size(177, 29);
            this.labelInstalled.TabIndex = 6;
            this.labelInstalled.Text = "已安装字体 (0)";
            // 
            // labelStatus
            // 
            this.labelStatus.AutoSize = true;
            this.labelStatus.Location = new System.Drawing.Point(24, 860);
            this.labelStatus.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.labelStatus.Name = "labelStatus";
            this.labelStatus.Size = new System.Drawing.Size(58, 24);
            this.labelStatus.TabIndex = 7;
            this.labelStatus.Text = "就绪";
            // 
            // btnReloadFonts
            // 
            this.btnReloadFonts.Location = new System.Drawing.Point(1396, 24);
            this.btnReloadFonts.Margin = new System.Windows.Forms.Padding(6);
            this.btnReloadFonts.Name = "btnReloadFonts";
            this.btnReloadFonts.Size = new System.Drawing.Size(180, 160);
            this.btnReloadFonts.TabIndex = 8;
            this.btnReloadFonts.Text = "重新加载\r\n字体库";
            this.btnReloadFonts.UseVisualStyleBackColor = true;
            this.btnReloadFonts.Click += new System.EventHandler(this.btnReloadFonts_Click);
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(24, 917);
            this.progressBar.Margin = new System.Windows.Forms.Padding(6);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(1552, 30);
            this.progressBar.TabIndex = 9;
            this.progressBar.Visible = false;
            // 
            // labelProgress
            // 
            this.labelProgress.AutoSize = true;
            this.labelProgress.Location = new System.Drawing.Point(24, 890);
            this.labelProgress.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.labelProgress.Name = "labelProgress";
            this.labelProgress.Size = new System.Drawing.Size(0, 24);
            this.labelProgress.TabIndex = 10;
            this.labelProgress.Visible = false;
            // 
            // MainApp
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(1600, 961);
            this.Controls.Add(this.labelProgress);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.btnReloadFonts);
            this.Controls.Add(this.labelStatus);
            this.Controls.Add(this.labelInstalled);
            this.Controls.Add(this.labelMissing);
            this.Controls.Add(this.btnExportInstalledFonts);
            this.Controls.Add(this.btnExportMissingList);
            this.Controls.Add(this.listBoxInstalledFonts);
            this.Controls.Add(this.listBoxMissingFonts);
            this.Controls.Add(this.labelDropZone);
            this.Margin = new System.Windows.Forms.Padding(6);
            this.Name = "MainApp";
            this.Text = "PowerPoint 字体工具";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelDropZone;
        private System.Windows.Forms.ListBox listBoxMissingFonts;
        private System.Windows.Forms.ListBox listBoxInstalledFonts;
        private System.Windows.Forms.Button btnExportMissingList;
        private System.Windows.Forms.Button btnExportInstalledFonts;
        private System.Windows.Forms.Label labelMissing;
        private System.Windows.Forms.Label labelInstalled;
        private System.Windows.Forms.Label labelStatus;
        private System.Windows.Forms.Button btnReloadFonts;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label labelProgress;
    }
}

