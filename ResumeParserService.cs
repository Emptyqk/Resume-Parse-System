using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Linq;
using 页面.Models;
using Xceed.Words.NET;

namespace 页面.Services
{
    // 简历解析服务
    public class ResumeParserService
    {
        // 解析Word文档

        public Resume ParseWordDocument(string filePath)
        {
            try
            {
                var resume = new Resume
                {
                    OriginalFilePath = filePath,
                    FileName = Path.GetFileName(filePath)
                };

                // 读取Word文档内容
                string content = ReadWordContent(filePath);
                
                // 解析简历内容
                ParseResumeContent(content, resume);
                
                return resume;
            }
            catch (Exception ex)
            {
                throw new Exception($"解析Word文档失败: {ex.Message}");
            }
        }

        // 解析PDF文档
        public Resume ParsePdfDocument(string filePath)
        {
            try
            {
                var resume = new Resume
                {
                    OriginalFilePath = filePath,
                    FileName = Path.GetFileName(filePath)
                };

                string content = ReadPdfContent(filePath);
                
                // 解析简历内容
                ParseResumeContent(content, resume);
                
                return resume;
            }
            catch (Exception ex)
            {
                throw new Exception($"解析PDF文档失败: {ex.Message}");
            }
        }

        // 解析简历内容
        private void ParseResumeContent(string content, Resume resume)
        {
            if (string.IsNullOrEmpty(content))
                return;

            // 提取姓名
            resume.Name = ExtractName(content);
            
            // 提取性别
            resume.Gender = ExtractGender(content);
            
            // 提取出生日期
            resume.BirthDate = ExtractBirthDate(content);
            
            // 提取地址
            resume.Address = ExtractAddress(content);
            
            // 提取联系方式
            resume.Phone = ExtractPhone(content);
            resume.Email = ExtractEmail(content);
            resume.IdCard = ExtractIdCard(content);
            
            // 提取教育信息
            ExtractEducationInfo(content, resume);
            
            // 提取工作经历
            ExtractWorkExperiences(content, resume);
            
            // 提取技能
            ExtractSkills(content, resume);
        }

        // 提取姓名>
        private string ExtractName(string content)
        {
            var lines = content.Split('\n');
            foreach (var line in lines)
            {
                if (line.Length >= 2 && line.Length <= 10 && IsChineseName(line.Trim()))
                {
                    return line.Trim();
                }
            }
            return "";
        }

        // 提取性别
        private string ExtractGender(string content)
        {
            if (content.Contains("男"))
                return "男";
            if (content.Contains("女"))
                return "女";
            return "";
        }

        // 提取出生日期
        private DateTime? ExtractBirthDate(string content)
        {
            var datePattern = @"(\d{4})[年\-/](\d{1,2})[月\-/](\d{1,2})[日]?";
            var match = Regex.Match(content, datePattern);
            
            if (match.Success)
            {
                if (int.TryParse(match.Groups[1].Value, out int year) &&
                    int.TryParse(match.Groups[2].Value, out int month) &&
                    int.TryParse(match.Groups[3].Value, out int day))
                {
                    try
                    {
                        return new DateTime(year, month, day);
                    }
                    catch
                    {
                        return null;
                    }
                }
            }
            return null;
        }

        // 提取地址
        private string ExtractAddress(string content)
        {
            // 简单的地址提取，查找包含"地址"、"住址"等关键词的行
            var addressKeywords = new[] { "地址", "住址", "现居", "居住地" };
            
            foreach (var keyword in addressKeywords)
            {
                var pattern = $@"{keyword}[：:]\s*([^\n\r]+)";
                var match = Regex.Match(content, pattern);
                if (match.Success)
                {
                    return match.Groups[1].Value.Trim();
                }
            }
            return "";
        }

        // 提取手机号码
        private string ExtractPhone(string content)
        {
            var phonePattern = @"1[3-9]\d{9}";
            var match = Regex.Match(content, phonePattern);
            return match.Success ? match.Value : "";
        }

        // 提取邮箱
        private string ExtractEmail(string content)
        {
            var emailPattern = @"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}";
            var match = Regex.Match(content, emailPattern);
            return match.Success ? match.Value : "";
        }

        // 提取身份证号
        private string ExtractIdCard(string content)
        {
            var idCardPattern = @"\d{17}[\dXx]";
            var match = Regex.Match(content, idCardPattern);
            return match.Success ? match.Value : "";
        }

        // 提取教育信息
        private void ExtractEducationInfo(string content, Resume resume)
        {
            // 提取所有学历信息
            var educationPattern = @"(大学|学院|学校)[：:]\s*([^\n\r]+)";
            var matches = Regex.Matches(content, educationPattern);
            
            if (matches.Count > 0)
            {
                // 第一个为第一学历
                resume.FirstEducationSchool = matches[0].Groups[2].Value.Trim();
                // 最后一个为最高学历
                resume.HighestEducationSchool = matches[matches.Count - 1].Groups[2].Value.Trim();
            }

            // 提取专业信息
            var majorPattern = @"专业[：:]\s*([^\n\r]+)";
            var majorMatches = Regex.Matches(content, majorPattern);
            if (majorMatches.Count > 0)
            {
                // 第一个为第一学历专业
                resume.FirstEducationMajor = majorMatches[0].Groups[1].Value.Trim();
                // 最后一个为最高学历专业
                resume.HighestEducationMajor = majorMatches[majorMatches.Count - 1].Groups[1].Value.Trim();
            }
            
            // 更精确地识别本科学历信息
            var bachelorPattern = @"本科[^\n\r]*[：:]?\s*([^\n\r]*(?:大学|学院))";
            var bachelorMatch = Regex.Match(content, bachelorPattern);
            if (bachelorMatch.Success && !string.IsNullOrEmpty(bachelorMatch.Groups[1].Value))
            {
                resume.FirstEducationSchool = bachelorMatch.Groups[1].Value.Trim();
            }
            
            // 识别本科专业
            var bachelorMajorPattern = @"本科[^\n\r]*专业[：:]\s*([^\n\r]+)";
            var bachelorMajorMatch = Regex.Match(content, bachelorMajorPattern);
            if (bachelorMajorMatch.Success)
            {
                resume.FirstEducationMajor = bachelorMajorMatch.Groups[1].Value.Trim();
            }
        }

        // 提取工作经历
        private void ExtractWorkExperiences(string content, Resume resume)
        {
            var workExperiencePattern = @"([^\n\r]+公司|[^\n\r]+集团|[^\n\r]+有限公司)[：:]?\s*([^\n\r]*岗位[^\n\r]*|[^\n\r]*职位[^\n\r]*|[^\n\r]*职务[^\n\r]*)";
            var matches = Regex.Matches(content, workExperiencePattern);
            
            foreach (Match match in matches)
            {
                var experience = new WorkExperience
                {
                    Company = match.Groups[1].Value.Trim(),
                    Position = match.Groups[2].Value.Trim()
                };
                resume.WorkExperiences.Add(experience);
            }
        }

        // 提取技能
        private void ExtractSkills(string content, Resume resume)
        {
            var skillKeywords = new[] { "技能", "专长", "能力", "掌握", "熟悉", "精通" };
            
            foreach (var keyword in skillKeywords)
            {
                var pattern = $@"{keyword}[：:]\s*([^\n\r]+)";
                var match = Regex.Match(content, pattern);
                if (match.Success)
                {
                    var skills = match.Groups[1].Value.Split(new[] { '、', '，', ',', '；', ';' });
                    foreach (var skill in skills)
                    {
                        var trimmedSkill = skill.Trim();
                        if (!string.IsNullOrEmpty(trimmedSkill))
                        {
                            resume.Skills.Add(trimmedSkill);
                        }
                    }
                }
            }
        }

        // 判断是否为中文姓名
        private bool IsChineseName(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;
            
            // 简单的中文姓名判断
            return text.All(c => char.IsLetter(c) || char.IsWhiteSpace(c)) && 
                   text.Length >= 2 && text.Length <= 10;
        }

        // 读取Word文档内容
        private string ReadWordContent(string filePath)
        {
            try
            {
                var extension = Path.GetExtension(filePath).ToLower();
                
                if (extension == ".docx")
                {
                    return ReadDocxContent(filePath);
                }
                else if (extension == ".doc")
                {
                    throw new NotSupportedException("doc格式不支持，我们尚未解决这个问题……");
                }
                else
                {
                    throw new NotSupportedException($"不支持的Word文档格式: {extension}");
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"读取Word文档失败: {ex.Message}");
            }
        }

        // 读取.docx文件内容
        private string ReadDocxContent(string filePath)
        {
            try
            {
                using (var document = DocX.Load(filePath))
                {
                    var sb = new StringBuilder();
                    
                    // 提取所有段落的文本
                    foreach (var paragraph in document.Paragraphs)
                    {
                        sb.AppendLine(paragraph.Text);
                    }
                    
                    // 提取表格中的文本
                    foreach (var table in document.Tables)
                    {
                        foreach (var row in table.Rows)
                        {
                            foreach (var cell in row.Cells)
                            {
                                foreach (var para in cell.Paragraphs)
                                {
                                    sb.AppendLine(para.Text);
                                }
                            }
                        }
                    }
                    
                    return sb.ToString();
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"读取DOCX文件失败: {ex.Message}");
            }
        }

        // 读取PDF文档内容
        private string ReadPdfContent(string filePath)
        {
            // 使用iText读取PDF文档
            using var pdfReader = new iText.Kernel.Pdf.PdfReader(filePath);
            using var pdfDocument = new iText.Kernel.Pdf.PdfDocument(pdfReader);
                
            var text = new System.Text.StringBuilder();
            for (int i = 1; i <= pdfDocument.GetNumberOfPages(); i++)
            {
                var page = pdfDocument.GetPage(i);
                var pageText = iText.Kernel.Pdf.Canvas.Parser.PdfTextExtractor.GetTextFromPage(page);
                text.AppendLine(pageText);
            }
                
            return text.ToString();
        }
    }
}
