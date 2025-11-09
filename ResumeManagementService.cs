using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using 页面.Models;

// 新增：导入结果类（放在命名空间内、ResumeManagementService类外）
public class ImportResult
{
    // 所有成功导入的简历（新增+替换）
    public List<Resume> ImportedResumes { get; set; } = new List<Resume>();
    // 新增的简历数量
    public int NewCount { get; set; }
    // 替换的简历数量
    public int ReplacedCount { get; set; }
}

namespace 页面.Services
{

    // 简历管理服务
    public class ResumeManagementService
    {
        private readonly string _dataDirectory;
        private readonly string _resumeDataFile;
        private readonly ResumeParserService _parserService;
        private readonly string _directoryDataFile;

        public ResumeManagementService()
        {
            _dataDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "智能简历解析系统");
            _resumeDataFile = Path.Combine(_dataDirectory, "resumes.json");

            // 目录数据文件路径（存储在同一数据目录下）
            _directoryDataFile = Path.Combine(_dataDirectory, "directories.json");

            _parserService = new ResumeParserService();

            // 确保数据目录存在
            if (!Directory.Exists(_dataDirectory))
            {
                Directory.CreateDirectory(_dataDirectory);
            }
        }

        //获取所有目录列表
        public List<ResumeDirectory> GetAllDirectories()
        {
            try
            {
                // 如果目录文件不存在，返回空列表
                if (!File.Exists(_directoryDataFile))
                {
                    return new List<ResumeDirectory>();
                }

                // 读取并反序列化目录数据
                var json = File.ReadAllText(_directoryDataFile);
                return JsonSerializer.Deserialize<List<ResumeDirectory>>(json) ?? new List<ResumeDirectory>();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"加载目录数据失败: {ex.Message}");
                return new List<ResumeDirectory>();
            }
        }

        //创建新目录（名称不可重复）
        public bool CreateDirectory(string directoryName)
        {
            if (string.IsNullOrWhiteSpace(directoryName))
                return false;

            // 获取现有目录列表
            var directories = GetAllDirectories();

            // 检查目录名称是否已存在（不区分大小写）
            if (directories.Exists(d => d.Name.Equals(directoryName, StringComparison.OrdinalIgnoreCase)))
                return false;

            // 添加新目录
            directories.Add(new ResumeDirectory
            {
                Id = Guid.NewGuid().ToString(), 
                Name = directoryName,
                CreatedTime = DateTime.Now
            });

            try
            {
                // 序列化并保存目录数据
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };
                var json = JsonSerializer.Serialize(directories, options);
                File.WriteAllText(_directoryDataFile, json);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"创建目录失败: {ex.Message}");
                return false;
            }
        }

        // 移动简历到指定目录
        public bool MoveResumeToDirectory(string resumeId, string directoryId)
        {
            if (string.IsNullOrWhiteSpace(resumeId) || string.IsNullOrWhiteSpace(directoryId))
                return false;
            var resumes = LoadResumes();
            var targetResume = resumes.FirstOrDefault(r => r.Id == resumeId);
            if (targetResume == null)
                return false;

            // 验证目录是否存在
            var directories = GetAllDirectories();
            if (!directories.Any(d => d.Id == directoryId))
                return false; // 目录不存在

            if (targetResume.DirectoryId == directoryId)
                return true;
            
            try
            {
                targetResume.DirectoryId = directoryId;
                // 保存更新后的简历列表
                SaveResumes(resumes);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"移动简历到目录失败: {ex.Message}");
                return false;
            }
        }

        //按目录筛选简历
        public List<Resume> GetResumesByDirectory(string directoryId = null)
        {
            var allResumes = LoadResumes();
            if (string.IsNullOrWhiteSpace(directoryId))
                return allResumes; // 目录ID为空时返回所有简历

            // 筛选出属于目标目录的简历
            return allResumes.Where(r => r.DirectoryId == directoryId).ToList();
        }

        // 删除目录方法
        public bool DeleteDirectory(string directoryId)
        {
            if (string.IsNullOrWhiteSpace(directoryId))
                return false;

            var directories = GetAllDirectories();
            var directoryToRemove = directories.FirstOrDefault(d => d.Id == directoryId);
            if (directoryToRemove == null)
                return false;
            var allResumes = LoadResumes();  // 重新加载所有简历
            int resumeCountInDir = allResumes.Count(r => r.DirectoryId == directoryId);

            if (resumeCountInDir > 0)
            {
                throw new InvalidOperationException(
                    $"目录下存在 {resumeCountInDir} 份简历，无法删除。请先将这些简历移动到其他目录。"
                );
            }

            directories.Remove(directoryToRemove);

            try
            {
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };
                var json = JsonSerializer.Serialize(directories, options);
                File.WriteAllText(_directoryDataFile, json);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"删除目录失败: {ex.Message}");
                return false;
            }
        }

        // 补充根据目录名称获取目录ID的方法
        public string GetDirectoryIdByName(string directoryName)
        {
            var directories = GetAllDirectories();
            return directories.FirstOrDefault(d => d.Name == directoryName)?.Id;
        }

        // 导入简历文件
        public ImportResult ImportResumes(List<string> filePaths, bool replaceExisting = false, string targetDirectoryId = null)
        {
            var importResult = new ImportResult();
            var existingResumes = LoadResumes();
            var existingFileNames = existingResumes.Select(r => r.FileName).ToHashSet();

            foreach (var filePath in filePaths)
            {
                try
                {
                    var fileName = Path.GetFileName(filePath);
                    bool isReplace = false;

                    // 检查文件是否已存在
                    if (existingFileNames.Contains(fileName))
                    {
                        var result = MessageBox.Show(
                            $"文件 {fileName} 已存在，是否替换？",
                            "文件已存在",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Exclamation);

                        if (result == DialogResult.No)
                        {
                            continue; // 跳过，不统计
                        }
                        isReplace = true; // 标记为替换
                    }

                    // 替换逻辑
                    if (isReplace)
                    {
                        existingResumes.RemoveAll(r => r.FileName == fileName);
                        importResult.ReplacedCount++; // 关键：替换数量+1
                    }
                    else
                    {
                        importResult.NewCount++; // 关键：非替换则为新增，数量+1
                    }

                    // 解析简历
                    Resume resume = null;
                    var extension = Path.GetExtension(filePath).ToLower();

                    switch (extension)
                    {
                        case ".doc":
                            resume = _parserService.ParseWordDocument(filePath);
                            break;
                        case ".docx":
                            resume = _parserService.ParseWordDocument(filePath);
                            break;
                        case ".pdf":
                            resume = _parserService.ParsePdfDocument(filePath);
                            break;
                        default:
                            throw new NotSupportedException($"不支持的文件格式: {extension}");
                    }

                    if (resume != null)
                    {
                        resume.ImportTime = DateTime.Now;
                        // 如果指定了目录，设置简历的目录ID
                        if (!string.IsNullOrWhiteSpace(targetDirectoryId))
                        {
                            resume.DirectoryId = targetDirectoryId;
                        }
                        existingResumes.Add(resume);
                        importResult.ImportedResumes.Add(resume); // 加入结果列表
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"导入文件 {filePath} 失败: {ex.Message}");
                }
            }

            // 保存更新后的简历列表
            if (importResult.ImportedResumes.Count > 0)
            {
                SaveResumes(existingResumes);
            }

            return importResult; // 返回包含统计数据的结果
        }

        // 加载所有简历 简历列表
        public List<Resume> LoadResumes()
        {
            try
            {
                if (File.Exists(_resumeDataFile))
                {
                    var json = File.ReadAllText(_resumeDataFile);
                    return JsonSerializer.Deserialize<List<Resume>>(json) ?? new List<Resume>();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"加载简历数据失败: {ex.Message}");
            }
            return new List<Resume>();
        }


        // 保存简历列表
        private void SaveResumes(List<Resume> resumes)
        {
            try
            {
                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };
                var json = JsonSerializer.Serialize(resumes, options);
                File.WriteAllText(_resumeDataFile, json);
            }
            catch (Exception ex)
            {
                throw new Exception($"保存简历数据失败: {ex.Message}");
            }
        }





        // 删除简历
        public bool DeleteResume(string resumeId)
        {
            var resumes = LoadResumes();
            var resume = resumes.FirstOrDefault(r => r.Id == resumeId);
            
            if (resume != null)
            {
                resumes.Remove(resume);
                SaveResumes(resumes);
                return true;
            }
            
            return false;
        }

        // 获取简历详情
        public Resume GetResumeById(string resumeId)
        {
            var resumes = LoadResumes();
            return resumes.FirstOrDefault(r => r.Id == resumeId);
        }

        // 搜索简历
        public List<Resume> SearchResumes(string keyword, DateTime? startDate, DateTime? endDate, string searchField, string fileNameKeyword)
        {
            var allResumes = LoadResumes();
            var results = new List<Resume>();

            foreach (var resume in allResumes)
            {
                bool matches = false;

                // 日期筛选（保持原有逻辑）
                if (startDate.HasValue && resume.ImportTime < startDate.Value)
                    continue;
                if (endDate.HasValue && resume.ImportTime > endDate.Value.AddDays(1))
                    continue;

                // 文件名独立筛选（新增逻辑）
                if (!string.IsNullOrWhiteSpace(fileNameKeyword))
                {
                    // 文件名不匹配则直接跳过
                    if (!(resume.FileName?.Contains(fileNameKeyword, StringComparison.OrdinalIgnoreCase) ?? false))
                        continue;
                }

                // 关键词搜索（移除原"文件名"分支，与文件名筛选分离）
                if (string.IsNullOrWhiteSpace(keyword))
                {
                    matches = true;
                }
                else
                {
                    switch (searchField)
                    {
                        case "姓名":
                            matches = resume.Name?.Contains(keyword, StringComparison.OrdinalIgnoreCase) ?? false;
                            break;
                        case "手机":
                            matches = resume.Phone?.Contains(keyword) ?? false;
                            break;
                        case "邮箱":
                            matches = resume.Email?.Contains(keyword, StringComparison.OrdinalIgnoreCase) ?? false;
                            break;
                        case "全部":
                        default:
                            // "全部"仅包含姓名、手机、邮箱（不再包含文件名，文件名通过独立参数筛选）
                            matches = (resume.Name?.Contains(keyword, StringComparison.OrdinalIgnoreCase) ?? false) ||
                                     (resume.Phone?.Contains(keyword) ?? false) ||
                                     (resume.Email?.Contains(keyword, StringComparison.OrdinalIgnoreCase) ?? false);
                            break;
                    }
                }

                if (matches)
                    results.Add(resume);
            }

            return results;
        }
        // 查找重复简历
        public Dictionary<string, List<Resume>> FindDuplicateResumes(bool checkName, bool checkPhone, bool checkEmail, bool checkIdCard)
        {
            var allResumes = LoadResumes();
            var duplicates = new Dictionary<string, List<Resume>>();

            if (checkName)
            {
                var nameGroups = allResumes
                    .Where(r => !string.IsNullOrWhiteSpace(r.Name))
                    .GroupBy(r => r.Name)
                    .Where(g => g.Count() > 1);

                foreach (var group in nameGroups)
                {
                    duplicates[$"姓名: {group.Key}"] = group.ToList();
                }
            }

            if (checkPhone)
            {
                var phoneGroups = allResumes
                    .Where(r => !string.IsNullOrWhiteSpace(r.Phone))
                    .GroupBy(r => r.Phone)
                    .Where(g => g.Count() > 1);

                foreach (var group in phoneGroups)
                {
                    duplicates[$"手机号: {group.Key}"] = group.ToList();
                }
            }

            if (checkEmail)
            {
                var emailGroups = allResumes
                    .Where(r => !string.IsNullOrWhiteSpace(r.Email))
                    .GroupBy(r => r.Email)
                    .Where(g => g.Count() > 1);

                foreach (var group in emailGroups)
                {
                    duplicates[$"邮箱: {group.Key}"] = group.ToList();
                }
            }

            if (checkIdCard)
            {
                var idCardGroups = allResumes
                    .Where(r => !string.IsNullOrWhiteSpace(r.IdCard))
                    .GroupBy(r => r.IdCard)
                    .Where(g => g.Count() > 1);

                foreach (var group in idCardGroups)
                {
                    duplicates[$"身份证号: {group.Key}"] = group.ToList();
                }
            }

            return duplicates;
        }

        // 导出查重结果到文本文件
        public void ExportDuplicatesToTxt(Dictionary<string, List<Resume>> duplicates, string filePath)
        {
            var sb = new StringBuilder();
            sb.AppendLine("简历查重结果");
            sb.AppendLine($"生成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"发现 {duplicates.Count} 组重复简历");
            sb.AppendLine();

            foreach (var duplicate in duplicates)
            {
                sb.AppendLine($"{duplicate.Key}");
                sb.AppendLine($"重复数量: {duplicate.Value.Count} 份");
                sb.AppendLine();

                foreach (var resume in duplicate.Value)
                {
                    sb.AppendLine($"    姓名: {resume.Name}");
                    sb.AppendLine($"    文件名: {resume.FileName}");
                    sb.AppendLine($"    手机: {resume.Phone}");
                    sb.AppendLine($"    邮箱: {resume.Email}");
                    sb.AppendLine($"    身份证号: {resume.IdCard}");
                    sb.AppendLine($"    导入时间: {resume.ImportTime:yyyy-MM-dd HH:mm:ss}");
                    sb.AppendLine();
                }
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        // 导出查重结果到Word文档
        public void ExportDuplicatesToWord(Dictionary<string, List<Resume>> duplicates, string filePath)
        {
            using (var document = Xceed.Words.NET.DocX.Create(filePath))
            {
                document.InsertParagraph("简历查重结果");
                document.InsertParagraph($"生成时间: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                document.InsertParagraph($"发现 {duplicates.Count} 组重复简历");
                document.InsertParagraph();

                foreach (var duplicate in duplicates)
                {
                    document.InsertParagraph($"{duplicate.Key}");
                    document.InsertParagraph($"重复数量: {duplicate.Value.Count} 份");
                    document.InsertParagraph();

                    foreach (var resume in duplicate.Value)
                    {
                        document.InsertParagraph($"    姓名: {resume.Name}");
                        document.InsertParagraph($"    文件名: {resume.FileName}");
                        document.InsertParagraph($"    手机: {resume.Phone}");
                        document.InsertParagraph($"    邮箱: {resume.Email}");
                        document.InsertParagraph($"    身份证号: {resume.IdCard}");
                        document.InsertParagraph($"    导入时间: {resume.ImportTime:yyyy-MM-dd HH:mm:ss}");
                        document.InsertParagraph();
                    }

                    document.InsertParagraph();
                }

                document.Save();
            }
        }
    }
}
