using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PowerPointFontTools
{
    /// <summary>
    /// UI层 - 只负责界面交互
    /// </summary>
    public partial class MainApp : Form
    {
        private FontManager fontManager;
        private HashSet<string> missingFonts = new HashSet<string>();
        private HashSet<string> installedPptFonts = new HashSet<string>();
        private string currentPptxPath = "";

        public MainApp()
        {
            InitializeComponent();
            InitializeFontManager();
        }

        private void InitializeFontManager()
        {
            // 设置数据库文件路径
            string appDataFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "PowerPointFontTools");

            if (!Directory.Exists(appDataFolder))
            {
                Directory.CreateDirectory(appDataFolder);
            }

            string fontDatabasePath = Path.Combine(appDataFolder, "fonts.db");

            // 创建字体管理器
            fontManager = new FontManager(fontDatabasePath);
            fontManager.StatusChanged += FontManager_StatusChanged;
            fontManager.ProgressChanged += FontManager_ProgressChanged;

            // 加载或构建字体数据库
            if (!fontManager.LoadDatabase())
            {
                labelStatus.Text = "首次运行，正在构建字体数据库...";
                Application.DoEvents();
                fontManager.ScanSystemFonts();
                fontManager.SaveDatabase();
            }
        }

        #region 事件处理器

        private void FontManager_StatusChanged(object sender, string status)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<object, string>(FontManager_StatusChanged), sender, status);
                return;
            }

            labelStatus.Text = status;
            Application.DoEvents();
        }

        private void FontManager_ProgressChanged(object sender, FontManager.ProgressEventArgs e)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<object, FontManager.ProgressEventArgs>(FontManager_ProgressChanged), sender, e);
                return;
            }

            progressBar.Visible = true;
            labelProgress.Visible = true;
            progressBar.Maximum = e.Total;
            progressBar.Value = Math.Min(e.Current, e.Total);
            labelProgress.Text = $"{e.Message} ({e.Current}/{e.Total})";
            Application.DoEvents();
        }

        #endregion

        #region 拖放处理

        private void labelDropZone_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length == 1 && Path.GetExtension(files[0]).ToLower() == ".pptx")
                {
                    e.Effect = DragDropEffects.Copy;
                    return;
                }
            }
            e.Effect = DragDropEffects.None;
        }

        private void labelDropZone_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files.Length > 0)
            {
                currentPptxPath = files[0];
                ProcessPptxFile(currentPptxPath);
            }
        }

        #endregion

        #region PPT文件处理

        private void ProcessPptxFile(string filePath)
        {
            try
            {
                listBoxMissingFonts.Items.Clear();
                listBoxInstalledFonts.Items.Clear();
                missingFonts.Clear();
                installedPptFonts.Clear();
                btnExportMissingList.Enabled = false;
                btnExportInstalledFonts.Enabled = false;

                labelDropZone.Text = "正在分析文件: " + Path.GetFileName(filePath);
                Application.DoEvents();

                // 获取 PPTX 中使用的所有字体
                HashSet<string> usedFonts = fontManager.GetFontsFromPptx(filePath);

                // 检查每个字体是否已安装
                List<string> allFonts = usedFonts.OrderBy(f => f).ToList();
                int missingCount = 0;
                int installedCount = 0;

                foreach (string font in allFonts)
                {
                    bool isInstalled = fontManager.IsFontInstalled(font);

                    if (!isInstalled)
                    {
                        missingFonts.Add(font);
                        listBoxMissingFonts.Items.Add(font);
                        missingCount++;
                    }
                    else
                    {
                        installedPptFonts.Add(font);
                        listBoxInstalledFonts.Items.Add(font);
                        installedCount++;
                    }
                }

                // 更新标签
                labelMissing.Text = $"缺失字体 ({missingCount}) - 双击搜索";
                labelInstalled.Text = $"已安装字体 ({installedCount})";

                // 启用相应的按钮
                if (missingCount > 0)
                {
                    btnExportMissingList.Enabled = true;
                }

                if (installedCount > 0)
                {
                    btnExportInstalledFonts.Enabled = true;
                }

                labelStatus.Text = $"分析完成：共 {allFonts.Count} 个字体，缺失 {missingCount} 个，已安装 {installedCount} 个";
                labelDropZone.Text = "分析完成 - 可以拖入新的 .pptx 文件";
            }
            catch (Exception ex)
            {
                MessageBox.Show("处理文件时出错: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                labelDropZone.Text = "拖入 .pptx 文件到此处";
            }
        }

        #endregion

        #region 导出功能

        private void btnExportMissingList_Click(object sender, EventArgs e)
        {
            if (missingFonts.Count == 0)
            {
                MessageBox.Show("没有缺失的字体需要导出。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "文本文件|*.txt";
                sfd.Title = "保存缺失字体名单";
                sfd.FileName = "missing_fonts_list.txt";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        StringBuilder sb = new StringBuilder();
                        sb.AppendLine("缺失的字体列表");
                        sb.AppendLine("================");
                        sb.AppendLine($"文件: {Path.GetFileName(currentPptxPath)}");
                        sb.AppendLine($"生成时间: {DateTime.Now}");
                        sb.AppendLine($"缺失字体数量: {missingFonts.Count}");
                        sb.AppendLine();

                        foreach (string font in missingFonts.OrderBy(f => f))
                        {
                            sb.AppendLine(font);
                        }

                        File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
                        MessageBox.Show($"缺失字体名单已导出！\n导出位置: {sfd.FileName}", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("导出名单时出错: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnExportInstalledFonts_Click(object sender, EventArgs e)
        {
            if (installedPptFonts.Count == 0)
            {
                MessageBox.Show("没有已安装的字体需要导出。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "ZIP 文件|*.zip";
                sfd.Title = "保存已安装字体文件";
                sfd.FileName = "installed_fonts.zip";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        // 显示进度条
                        progressBar.Visible = true;
                        labelProgress.Visible = true;
                        progressBar.Value = 0;

                        fontManager.ExportFontsToZip(installedPptFonts, sfd.FileName, Path.GetFileName(currentPptxPath));

                        // 隐藏进度条
                        progressBar.Visible = false;
                        labelProgress.Visible = false;

                        MessageBox.Show($"已安装字体导出完成！\n导出位置: {sfd.FileName}", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        progressBar.Visible = false;
                        labelProgress.Visible = false;
                        MessageBox.Show("导出字体时出错: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        #endregion

        #region 其他功能

        private void listBoxMissingFonts_DoubleClick(object sender, EventArgs e)
        {
            if (listBoxMissingFonts.SelectedItem != null)
            {
                string fontName = listBoxMissingFonts.SelectedItem.ToString();
                string searchUrl = $"https://www.bing.com/search?q={Uri.EscapeDataString(fontName + " font download")}";

                try
                {
                    System.Diagnostics.Process.Start(searchUrl);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("无法打开浏览器: " + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnReloadFonts_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(
                "重新构建字体数据库将重新扫描所有字体文件，这可能需要几秒钟。\n\n旧的数据库将被覆盖，是否继续？",
                "确认",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                fontManager.ScanSystemFonts();
                fontManager.SaveDatabase();

                // 如果有已加载的PPT文件，重新分析
                if (!string.IsNullOrEmpty(currentPptxPath) && File.Exists(currentPptxPath))
                {
                    ProcessPptxFile(currentPptxPath);
                }
            }
        }

        #endregion
    }
}
