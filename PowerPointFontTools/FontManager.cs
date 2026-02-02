using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace PowerPointFontTools
{
    /// <summary>
    /// 字体管理器 - 负责字体数据库和业务逻辑
    /// </summary>
    public class FontManager
    {
        private Dictionary<string, FontInfo> installedFontMap = new Dictionary<string, FontInfo>(StringComparer.OrdinalIgnoreCase);
        private string fontDatabasePath;

        public event EventHandler<string> StatusChanged;
        public event EventHandler<ProgressEventArgs> ProgressChanged;

        public FontManager(string databasePath)
        {
            fontDatabasePath = databasePath;
        }

        #region 数据类定义

        public class FontInfo
        {
            public string FamilyName { get; set; }
            public HashSet<string> AlternativeNames { get; set; }
            public List<string> FilePaths { get; set; }
            public bool IsInstalled { get; set; }

            public FontInfo()
            {
                AlternativeNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                FilePaths = new List<string>();
            }
        }

        [XmlRoot("FontDatabase")]
        public class FontDatabase
        {
            public DateTime CreatedTime { get; set; }
            public int TotalFonts { get; set; }
            public int TotalFiles { get; set; }

            [XmlArray("Fonts")]
            [XmlArrayItem("Font")]
            public List<FontInfoData> Fonts { get; set; }

            public FontDatabase()
            {
                Fonts = new List<FontInfoData>();
            }
        }

        public class FontInfoData
        {
            public string FamilyName { get; set; }

            [XmlArray("AlternativeNames")]
            [XmlArrayItem("Name")]
            public List<string> AlternativeNames { get; set; }

            [XmlArray("FilePaths")]
            [XmlArrayItem("Path")]
            public List<string> FilePaths { get; set; }

            public bool IsInstalled { get; set; }

            public FontInfoData()
            {
                AlternativeNames = new List<string>();
                FilePaths = new List<string>();
            }
        }

        public class ProgressEventArgs : EventArgs
        {
            public int Current { get; set; }
            public int Total { get; set; }
            public string Message { get; set; }
        }

        #endregion

        #region 数据库操作

        public bool LoadDatabase()
        {
            try
            {
                if (File.Exists(fontDatabasePath))
                {
                    OnStatusChanged("正在加载字体数据库...");

                    XmlSerializer serializer = new XmlSerializer(typeof(FontDatabase));
                    using (FileStream fs = new FileStream(fontDatabasePath, FileMode.Open))
                    {
                        FontDatabase database = (FontDatabase)serializer.Deserialize(fs);

                        installedFontMap.Clear();
                        foreach (FontInfoData fontData in database.Fonts)
                        {
                            FontInfo fontInfo = new FontInfo
                            {
                                FamilyName = fontData.FamilyName,
                                IsInstalled = fontData.IsInstalled
                            };

                            foreach (string name in fontData.AlternativeNames)
                            {
                                fontInfo.AlternativeNames.Add(name);
                            }

                            foreach (string path in fontData.FilePaths)
                            {
                                fontInfo.FilePaths.Add(path);
                            }

                            foreach (string altName in fontInfo.AlternativeNames)
                            {
                                if (!installedFontMap.ContainsKey(altName))
                                {
                                    installedFontMap[altName] = fontInfo;
                                }
                            }
                        }

                        OnStatusChanged($"字体数据库已加载：{database.TotalFonts} 个字体，{database.TotalFiles} 个文件 (创建于: {database.CreatedTime:yyyy-MM-dd HH:mm})");
                        return true;
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                OnStatusChanged($"加载字体数据库失败: {ex.Message}");
                return false;
            }
        }

        public void SaveDatabase()
        {
            try
            {
                OnStatusChanged("正在保存字体数据库...");

                FontDatabase database = new FontDatabase
                {
                    CreatedTime = DateTime.Now,
                    TotalFonts = installedFontMap.Values.Distinct().Count(),
                    TotalFiles = installedFontMap.Values.Distinct().Sum(f => f.FilePaths.Count)
                };

                foreach (FontInfo fontInfo in installedFontMap.Values.Distinct())
                {
                    FontInfoData fontData = new FontInfoData
                    {
                        FamilyName = fontInfo.FamilyName,
                        IsInstalled = fontInfo.IsInstalled,
                        AlternativeNames = fontInfo.AlternativeNames.ToList(),
                        FilePaths = fontInfo.FilePaths.ToList()
                    };
                    database.Fonts.Add(fontData);
                }

                XmlSerializer serializer = new XmlSerializer(typeof(FontDatabase));
                using (FileStream fs = new FileStream(fontDatabasePath, FileMode.Create))
                {
                    serializer.Serialize(fs, database);
                }

                OnStatusChanged($"字体数据库已保存：{database.TotalFonts} 个字体，{database.TotalFiles} 个文件");
            }
            catch (Exception ex)
            {
                OnStatusChanged($"保存字体数据库失败: {ex.Message}");
            }
        }

        #endregion

        #region 字体扫描

        public void ScanSystemFonts()
        {
            try
            {
                OnStatusChanged("正在扫描字体文件...");

                installedFontMap.Clear();

                List<string> fontFolders = new List<string>();
                fontFolders.Add(Environment.GetFolderPath(Environment.SpecialFolder.Fonts));

                string userFontsFolder1 = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Microsoft", "Windows", "Fonts");
                if (Directory.Exists(userFontsFolder1))
                {
                    fontFolders.Add(userFontsFolder1);
                }

                string userFontsFolder2 = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                    "AppData", "Local", "Microsoft", "Windows", "Fonts");
                if (Directory.Exists(userFontsFolder2) && !fontFolders.Contains(userFontsFolder2))
                {
                    fontFolders.Add(userFontsFolder2);
                }

                int totalFiles = 0;
                int processedFiles = 0;
                int errorFiles = 0;

                foreach (string fontsFolder in fontFolders)
                {
                    if (Directory.Exists(fontsFolder))
                    {
                        OnStatusChanged($"正在扫描: {Path.GetFileName(fontsFolder)}...");

                        string[] fontFiles = Directory.GetFiles(fontsFolder, "*.ttf")
                            .Concat(Directory.GetFiles(fontsFolder, "*.otf"))
                            .Concat(Directory.GetFiles(fontsFolder, "*.TTF"))
                            .Concat(Directory.GetFiles(fontsFolder, "*.OTF"))
                            .ToArray();

                        totalFiles += fontFiles.Length;

                        foreach (string fontFile in fontFiles)
                        {
                            try
                            {
                                processedFiles++;
                                if (processedFiles % 10 == 0)
                                {
                                    OnStatusChanged($"正在扫描字体文件... ({processedFiles}/{totalFiles})");
                                    OnProgressChanged(processedFiles, totalFiles, $"扫描字体: {Path.GetFileName(fontFile)}");
                                }

                                string[] fontNames = GetFontNamesFromFile(fontFile);

                                if (fontNames != null && fontNames.Length > 0)
                                {
                                    string primaryFontName = fontNames[0];

                                    FontInfo fontInfo = installedFontMap.Values
                                        .FirstOrDefault(f => f.FamilyName.Equals(primaryFontName, StringComparison.OrdinalIgnoreCase));

                                    if (fontInfo == null)
                                    {
                                        fontInfo = new FontInfo
                                        {
                                            FamilyName = primaryFontName,
                                            IsInstalled = true
                                        };

                                        foreach (string name in fontNames)
                                        {
                                            fontInfo.AlternativeNames.Add(name);
                                        }

                                        string fileNameWithoutExt = Path.GetFileNameWithoutExtension(fontFile);
                                        fontInfo.AlternativeNames.Add(fileNameWithoutExt);
                                        fontInfo.AlternativeNames.Add(fileNameWithoutExt.Replace(" ", ""));
                                        fontInfo.AlternativeNames.Add(fileNameWithoutExt.Replace(" ", "-"));
                                        fontInfo.AlternativeNames.Add(fileNameWithoutExt.Replace(" ", "_"));
                                        fontInfo.AlternativeNames.Add(fileNameWithoutExt.Replace("-", ""));
                                        fontInfo.AlternativeNames.Add(fileNameWithoutExt.Replace("_", ""));

                                        foreach (string altName in fontInfo.AlternativeNames)
                                        {
                                            if (!installedFontMap.ContainsKey(altName))
                                            {
                                                installedFontMap[altName] = fontInfo;
                                            }
                                        }
                                    }

                                    if (!fontInfo.FilePaths.Contains(fontFile))
                                    {
                                        fontInfo.FilePaths.Add(fontFile);
                                    }
                                }
                                else
                                {
                                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(fontFile);

                                    FontInfo fontInfo = new FontInfo
                                    {
                                        FamilyName = fileNameWithoutExt,
                                        IsInstalled = false
                                    };
                                    fontInfo.FilePaths.Add(fontFile);
                                    fontInfo.AlternativeNames.Add(fileNameWithoutExt);

                                    if (!installedFontMap.ContainsKey(fileNameWithoutExt))
                                    {
                                        installedFontMap[fileNameWithoutExt] = fontInfo;
                                    }
                                }
                            }
                            catch
                            {
                                errorFiles++;
                            }
                        }
                    }
                }

                int uniqueFonts = installedFontMap.Values.Distinct().Count();
                int fontsWithFiles = installedFontMap.Values.Distinct().Count(f => f.FilePaths.Count > 0);

                OnStatusChanged($"字体扫描完成：{uniqueFonts} 个字体，{fontsWithFiles} 个有文件，共 {totalFiles} 个文件 (处理: {processedFiles}, 错误: {errorFiles})");
            }
            catch (Exception ex)
            {
                OnStatusChanged($"加载系统字体时出错: {ex.Message}");
            }
        }

        private string[] GetFontNamesFromFile(string fontFilePath)
        {
            List<string> fontNames = new List<string>();

            try
            {
                using (PrivateFontCollection pfc = new PrivateFontCollection())
                {
                    pfc.AddFontFile(fontFilePath);

                    if (pfc.Families.Length > 0)
                    {
                        foreach (System.Drawing.FontFamily family in pfc.Families)
                        {
                            fontNames.Add(family.Name);
                        }
                    }
                }
            }
            catch
            {
                return null;
            }

            return fontNames.ToArray();
        }

        #endregion

        #region 字体检测

        public bool IsFontInstalled(string fontName)
        {
            if (string.IsNullOrWhiteSpace(fontName))
                return false;

            if (installedFontMap.ContainsKey(fontName))
            {
                FontInfo fontInfo = installedFontMap[fontName];
                return fontInfo.FilePaths.Count > 0;
            }

            string[] variants = {
                fontName,
                fontName.Replace(" ", ""),
                fontName.Replace(" ", "-"),
                fontName.Replace(" ", "_"),
                fontName.Replace("-", " "),
                fontName.Replace("_", " ")
            };

            foreach (string variant in variants)
            {
                if (installedFontMap.ContainsKey(variant))
                {
                    FontInfo fontInfo = installedFontMap[variant];
                    return fontInfo.FilePaths.Count > 0;
                }
            }

            return false;
        }

        public HashSet<string> GetFontsFromPptx(string filePath)
        {
            HashSet<string> fonts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            using (ZipArchive archive = ZipFile.OpenRead(filePath))
            {
                foreach (ZipArchiveEntry entry in archive.Entries)
                {
                    if (entry.FullName.StartsWith("ppt/slides/slide") &&
                        entry.FullName.EndsWith(".xml") &&
                        !entry.FullName.Contains("slideMaster") &&
                        !entry.FullName.Contains("slideLayout"))
                    {
                        using (Stream stream = entry.Open())
                        {
                            XDocument doc = XDocument.Load(stream);
                            ExtractFontsFromSlide(doc, fonts);
                        }
                    }
                }
            }

            return fonts;
        }

        private void ExtractFontsFromSlide(XDocument doc, HashSet<string> fonts)
        {
            var allFontElements = doc.Descendants()
                .SelectMany(e => e.Attributes())
                .Where(a => a.Name.LocalName == "typeface");

            foreach (var attr in allFontElements)
            {
                string fontName = attr.Value;
                if (!string.IsNullOrWhiteSpace(fontName) &&
                    !fontName.StartsWith("+") &&
                    !fontName.All(char.IsDigit))
                {
                    fonts.Add(fontName);
                }
            }

            var charsetFonts = doc.Descendants()
                .SelectMany(e => e.Attributes())
                .Where(a => a.Name.LocalName == "charset");

            foreach (var attr in charsetFonts)
            {
                string fontName = attr.Value;
                if (!string.IsNullOrWhiteSpace(fontName) &&
                    !fontName.StartsWith("+") &&
                    !fontName.All(char.IsDigit))
                {
                    fonts.Add(fontName);
                }
            }
        }

        #endregion

        #region 字体导出

        public void ExportFontsToZip(HashSet<string> fontNames, string zipPath, string sourcePptxName)
        {
            if (File.Exists(zipPath))
            {
                File.Delete(zipPath);
            }

            using (ZipArchive archive = ZipFile.Open(zipPath, ZipArchiveMode.Create))
            {
                int exportedCount = 0;
                int current = 0;
                int total = fontNames.Count;
                StringBuilder notFoundFonts = new StringBuilder();
                HashSet<string> addedFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (string fontName in fontNames)
                {
                    current++;
                    OnProgressChanged(current, total, $"导出字体: {fontName}");

                    bool found = false;

                    if (installedFontMap.ContainsKey(fontName))
                    {
                        FontInfo fontInfo = installedFontMap[fontName];
                        if (fontInfo.FilePaths.Count > 0)
                        {
                            foreach (string fontFilePath in fontInfo.FilePaths)
                            {
                                if (File.Exists(fontFilePath))
                                {
                                    string fileName = Path.GetFileName(fontFilePath);

                                    if (!addedFiles.Contains(fileName))
                                    {
                                        archive.CreateEntryFromFile(fontFilePath, fileName);
                                        addedFiles.Add(fileName);
                                        exportedCount++;
                                    }

                                    found = true;
                                }
                            }
                        }
                    }

                    if (!found)
                    {
                        notFoundFonts.AppendLine(fontName);
                    }
                }

                ZipArchiveEntry infoEntry = archive.CreateEntry("导出信息.txt");
                using (StreamWriter writer = new StreamWriter(infoEntry.Open()))
                {
                    writer.WriteLine("字体导出信息");
                    writer.WriteLine("==================");
                    writer.WriteLine($"来源文件: {sourcePptxName}");
                    writer.WriteLine($"导出时间: {DateTime.Now}");
                    writer.WriteLine($"成功导出: {exportedCount} 个字体文件");
                    writer.WriteLine($"请求字体数: {fontNames.Count}");
                    writer.WriteLine();

                    if (notFoundFonts.Length > 0)
                    {
                        writer.WriteLine("以下字体未找到对应的字体文件：");
                        writer.WriteLine();
                        writer.Write(notFoundFonts.ToString());
                        writer.WriteLine();
                        writer.WriteLine("这些字体可能：");
                        writer.WriteLine("1. 未正确安装在系统中");
                        writer.WriteLine("2. 使用了完全不同的文件名");
                        writer.WriteLine("3. 是系统字体（.ttc 格式）");
                        writer.WriteLine("4. 需要从其他来源获取");
                    }
                    else
                    {
                        writer.WriteLine("所有字体文件均已成功找到并导出！");
                    }
                }
            }

            OnProgressChanged(fontNames.Count, fontNames.Count, "导出完成");
        }

        #endregion

        #region 事件触发

        protected virtual void OnStatusChanged(string status)
        {
            StatusChanged?.Invoke(this, status);
        }

        protected virtual void OnProgressChanged(int current, int total, string message)
        {
            ProgressChanged?.Invoke(this, new ProgressEventArgs
            {
                Current = current,
                Total = total,
                Message = message
            });
        }

        #endregion
    }
}
