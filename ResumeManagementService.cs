using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using 页面.Models;
namespace 页面.Services
{

    // 简历管理服务
    public class ResumeManagementService
    {
        private readonly string _dataDirectory;
        private readonly string _resumeDataFile;
        private readonly string _directoriesFile;
        private readonly ResumeParserService _parserService;

        public ResumeManagementService()
        {
            _dataDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "智能简历解析系统");
            _resumeDataFile = Path.Combine(_dataDirectory, "resumes.json");
            _directoriesFile = Path.Combine(_dataDirectory, "directories.json");
            _parserService = new ResumeParserService();

            // 确保数据目录存在
            if (!Directory.Exists(_dataDirectory))
            {
                Directory.CreateDirectory(_dataDirectory);
            }
        }

        // 导入简历文件
        public List<Resume> ImportResumes(List<string> filePaths, bool replaceExisting = false)
        {
            var importedResumes = new List<Resume>();
            var existingResumes = LoadResumes();
            var existingFileNames = existingResumes.Select(r => r.FileName).ToHashSet();

            foreach (var filePath in filePaths)
            {
                try
                {
                    var fileName = Path.GetFileName(filePath);
                    
                    // 检查文件是否已存在
                    if (existingFileNames.Contains(fileName) && !replaceExisting)
                    {
                        continue; // 跳过已存在的文件
                    }

                    // 删除已存在的同名简历
                    if (replaceExisting && existingFileNames.Contains(fileName))
                    {
                        existingResumes.RemoveAll(r => r.FileName == fileName);
                    }

                    // 解析简历
                    Resume resume = null;
                    var extension = Path.GetExtension(filePath).ToLower();
                    
                    switch (extension)
                    {
                        case ".doc":
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
                        existingResumes.Add(resume);
                        importedResumes.Add(resume);
                    }
                }
                catch (Exception ex)
                {
                    // 记录错误但继续处理其他文件
                    Console.WriteLine($"导入文件 {filePath} 失败: {ex.Message}");
                }
            }

            // 保存更新后的简历列表
            if (importedResumes.Count > 0)
            {
                SaveResumes(existingResumes);
            }

            return importedResumes;
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
        public List<Resume> SearchResumes(string keyword, DateTime? startDate, DateTime? endDate, string searchField)
        {
            var allResumes = LoadResumes();
            var results = new List<Resume>();

            foreach (var resume in allResumes)
            {
                bool matches = false;

                // 日期筛选
                if (startDate.HasValue && resume.ImportTime < startDate.Value)
                    continue;
                if (endDate.HasValue && resume.ImportTime > endDate.Value.AddDays(1))
                    continue;

                // 关键词搜索
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
                        case "文件名":
                            matches = resume.FileName?.Contains(keyword, StringComparison.OrdinalIgnoreCase) ?? false;
                            break;
                        case "手机":
                            matches = resume.Phone?.Contains(keyword) ?? false;
                            break;
                        case "邮箱":
                            matches = resume.Email?.Contains(keyword, StringComparison.OrdinalIgnoreCase) ?? false;
                            break;
                        case "全部":
                        default:
                            matches = (resume.Name?.Contains(keyword, StringComparison.OrdinalIgnoreCase) ?? false) ||
                                     (resume.FileName?.Contains(keyword, StringComparison.OrdinalIgnoreCase) ?? false) ||
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
