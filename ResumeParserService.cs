using NPOI.HWPF;
using NPOI.HWPF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XWPF.UserModel;
using NPOI.POIFS.FileSystem;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Xceed.Words.NET;
using 页面.Models;

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
                
                // 如果从内容中提取不到姓名，尝试从文件名中提取
                if (string.IsNullOrWhiteSpace(resume.Name))
                {
                    resume.Name = ExtractNameFromFileName(resume.FileName);
                }
                
                return resume;
            }
            catch (Exception ex)
            {
                throw new Exception($"解析Word文档失败: {ex.Message}");
            }
        }
        
        // 从文件名中提取姓名
        private string ExtractNameFromFileName(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
                return "";
            
            var nameWithoutExt = Path.GetFileNameWithoutExtension(fileName);
            nameWithoutExt = Regex.Replace(nameWithoutExt, @"[\u200B-\u200D\uFEFF]", "");
            nameWithoutExt = Regex.Replace(nameWithoutExt, @"(个人)?简历.*$", "", RegexOptions.IgnoreCase);
            nameWithoutExt = nameWithoutExt.Trim();
            
            // 检查是否为中文姓名
            if (nameWithoutExt.Length >= 2 && nameWithoutExt.Length <= 4 && IsChineseName(nameWithoutExt))
            {
                return nameWithoutExt;
            }
            
            return "";
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

        // 提取姓名
        private string ExtractName(string content)
        {
            if (string.IsNullOrEmpty(content))
                return "";

            // 清理零宽字符和其他特殊字符
            content = Regex.Replace(content, @"[\u200B-\u200D\uFEFF]", ""); // 移除零宽字符
            
            var lines = content.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            
            // 优先查找前几行中的姓名
            for (int i = 0; i < Math.Min(10, lines.Length); i++)
            {
                var line = lines[i].Trim();
                // 移除可能的特殊字符和控制字符
                line = Regex.Replace(line, @"[\x00-\x1F\x7F\u200B-\u200D\uFEFF]", "");
                
                // 检查是否为中文姓名（2-4个中文字符）
                if (line.Length >= 2 && line.Length <= 4 && IsChineseName(line))
                {
                    return line;
                }
            }
            
            // 如果前几行没找到，尝试在整个内容中查找
            // 查找类似"姓名：XXX"的格式
            var namePatterns = new[]
            {
                @"姓名[：:]\s*([^\n\r]{2,4})",
                @"姓名[：:]\s*([^\s：:]{2,4})",
                @"^([^\n\r]{2,4})\s*$" // 单独一行的2-4个字符
            };
            
            foreach (var pattern in namePatterns)
            {
                var match = Regex.Match(content, pattern, RegexOptions.Multiline);
                if (match.Success)
                {
                    var name = match.Groups[1].Value.Trim();
                    name = Regex.Replace(name, @"[\u200B-\u200D\uFEFF]", ""); // 移除零宽字符
                    if (IsChineseName(name))
                    {
                        return name;
                    }
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
            if (string.IsNullOrEmpty(content))
                return;

            // 清理内容，确保换行符统一
            content = CleanContentFormat(content);
            
            // 多种工作经历匹配模式
            var patterns = new[]
            {
                // 模式1: 公司名称 + 职位（带冒号）
                @"([^\n\r]{2,30}(?:公司|集团|有限公司|企业|股份)[^\n\r]{0,20})[：:]\s*([^\n\r]{2,30}(?:岗位|职位|职务|担任)[^\n\r]{0,30})",
                // 模式2: 公司名称（不带冒号，后面跟职位关键词）
                @"([^\n\r]{2,30}(?:公司|集团|有限公司|企业|股份)[^\n\r]{0,20})\s+([^\n\r]{2,30}(?:岗位|职位|职务|担任)[^\n\r]{0,30})",
                // 模式3: 工作经历段落中的公司+职位
                @"工作经历[^\n\r]*\n[^\n\r]*([^\n\r]{2,30}(?:公司|集团|有限公司)[^\n\r]{0,20})[^\n\r]*([^\n\r]{2,30}(?:岗位|职位|职务)[^\n\r]{0,30})",
                // 模式4: 简单的公司名称（至少包含"公司"等关键词）
                @"([^\n\r]{2,30}(?:公司|集团|有限公司|企业|股份)[^\n\r]{0,20})"
            };
            
            var foundCompanies = new HashSet<string>();
            
            foreach (var pattern in patterns)
            {
                var matches = Regex.Matches(content, pattern, RegexOptions.Multiline);
                foreach (Match match in matches)
                {
                    string company = "";
                    string position = "";
                    
                    if (match.Groups.Count > 1)
                    {
                        company = match.Groups[1].Value.Trim();
                        // 清理公司名称中的多余字符
                        company = Regex.Replace(company, @"[：:\s]+$", "");
                        
                        if (match.Groups.Count > 2 && !string.IsNullOrWhiteSpace(match.Groups[2].Value))
                        {
                            position = match.Groups[2].Value.Trim();
                        }
                    }
                    
                    // 避免重复添加相同的公司
                    if (!string.IsNullOrWhiteSpace(company) && !foundCompanies.Contains(company))
                    {
                        foundCompanies.Add(company);
                        resume.WorkExperiences.Add(new WorkExperience
                        {
                            Company = company,
                            Position = position
                        });
                    }
                }
            }
        }

        // 提取技能
        private void ExtractSkills(string content, Resume resume)
        {
            if (string.IsNullOrEmpty(content))
                return;

            // 清理内容，包括零宽字符
            content = CleanContentFormat(content);
            content = Regex.Replace(content, @"[\u200B-\u200D\uFEFF]", ""); // 移除零宽字符
            
            var skillKeywords = new[] { "技能", "专长", "能力", "掌握", "熟悉", "精通", "擅长" };
            var foundSkills = new HashSet<string>();
            
            foreach (var keyword in skillKeywords)
            {
                // 模式1: 关键词后跟冒号，技能在同一行
                var pattern1 = $@"{keyword}[：:]\s*([^\n\r]+)";
                var match1 = Regex.Match(content, pattern1, RegexOptions.IgnoreCase);
                if (match1.Success)
                {
                    var skillText = match1.Groups[1].Value.Trim();
                    skillText = Regex.Replace(skillText, @"[\u200B-\u200D\uFEFF]", ""); // 移除零宽字符
                    if (!string.IsNullOrWhiteSpace(skillText))
                    {
                        ExtractSkillsFromText(skillText, resume, foundSkills);
                    }
                }
                
                // 模式2: 关键词后跟换行，技能在下一行或多行（最多5行）
                var pattern2 = $@"{keyword}[：:]\s*\n\s*([^\n\r]+(?:\n[^\n\r]+){{0,5}})";
                var match2 = Regex.Match(content, pattern2, RegexOptions.Multiline | RegexOptions.IgnoreCase);
                if (match2.Success)
                {
                    var skillText = match2.Groups[1].Value.Trim();
                    skillText = Regex.Replace(skillText, @"[\u200B-\u200D\uFEFF]", ""); // 移除零宽字符
                    if (!string.IsNullOrWhiteSpace(skillText))
                    {
                        ExtractSkillsFromText(skillText, resume, foundSkills);
                    }
                }
                
                // 模式3: 关键词在同一行，后面直接跟技能内容（至少5个字符）
                var pattern3 = $@"{keyword}[：:]\s*([^\n\r]{{5,200}})";
                var match3 = Regex.Match(content, pattern3, RegexOptions.IgnoreCase);
                if (match3.Success)
                {
                    var skillText = match3.Groups[1].Value.Trim();
                    skillText = Regex.Replace(skillText, @"[\u200B-\u200D\uFEFF]", ""); // 移除零宽字符
                    if (!string.IsNullOrWhiteSpace(skillText))
                    {
                        ExtractSkillsFromText(skillText, resume, foundSkills);
                    }
                }
                
                // 模式4: 查找"技能"关键词后的所有内容，直到下一个主要章节
                var pattern4 = $@"{keyword}[：:]\s*([^\n\r]*(?:\n(?!姓名|性别|出生|地址|手机|邮箱|身份证|学历|工作|教育)[^\n\r]+)*)";
                var match4 = Regex.Match(content, pattern4, RegexOptions.Multiline | RegexOptions.IgnoreCase);
                if (match4.Success)
                {
                    var skillText = match4.Groups[1].Value.Trim();
                    skillText = Regex.Replace(skillText, @"[\u200B-\u200D\uFEFF]", ""); // 移除零宽字符
                    if (!string.IsNullOrWhiteSpace(skillText))
                    {
                        ExtractSkillsFromText(skillText, resume, foundSkills);
                    }
                }
            }
            
            // 模式5: 专门处理"数字、技能关键词"格式（如：四、专业技能）
            // 匹配格式：一/二/三/四/五/六/七/八/九/十/1/2/3/4/5/6/7/8/9/10 + 、 + 技能关键词
            var numberedSkillPattern = @"[一二三四五六七八九十\d]+[、.]\s*([^\n\r]*(?:技能|专长|能力|技术)[^\n\r]*)";
            var numberedMatches = Regex.Matches(content, numberedSkillPattern, RegexOptions.IgnoreCase);
            
            foreach (Match match in numberedMatches)
            {
                if (match.Success && match.Groups.Count > 1)
                {
                    var titleLine = match.Groups[1].Value.Trim();
                    
                    var skillTitleKeywords = new[] { "技能", "专长", "能力", "技术", "专业", "核心", "主要" };
                    bool isSkillSection = skillTitleKeywords.Any(k => titleLine.Contains(k));
                    
                    if (isSkillSection)
                    {
                        var titleMatch = match.Groups[0];
                        var titleStartIndex = match.Index;
                        var sectionContent = ExtractNumberedSectionContent(content, titleStartIndex);
                        
                        if (!string.IsNullOrWhiteSpace(sectionContent))
                        {
                            ExtractSkillsFromParagraphs(sectionContent, resume, foundSkills);
                        }
                    }
                }
            }
        }
        
        private string ExtractNumberedSectionContent(string content, int startIndex)
        {
            if (startIndex >= content.Length)
                return "";
          
            var remainingContent = content.Substring(startIndex);
            var lines = remainingContent.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            
            var result = new StringBuilder();
            bool isFirstLine = true;
            
            foreach (var line in lines)
            {
                var trimmedLine = line.Trim();
                
                if (isFirstLine)
                {
                    isFirstLine = false;
                    continue;
                }
                
                if (Regex.IsMatch(trimmedLine, @"^[一二三四五六七八九十\d]+[、.]\s*"))
                {
                    break;
                }
                
                var sectionKeywords = new[] { "姓名", "性别", "出生", "地址", "手机", "邮箱", "身份证", 
                                            "学历", "工作", "教育", "经历", "项目", "获奖", "证书", 
                                            "自我评价", "个人评价", "文件信息" };
                bool isNextSection = sectionKeywords.Any(k => 
                    trimmedLine.StartsWith(k, StringComparison.OrdinalIgnoreCase) ||
                    trimmedLine.Contains($"{k}：") || 
                    trimmedLine.Contains($"{k}:"));
                
                if (isNextSection)
                    break;
               
                if (string.IsNullOrWhiteSpace(trimmedLine))
                    continue;
                
                if (result.Length > 0)
                    result.AppendLine();
                result.Append(trimmedLine);
            }
            
            return result.ToString();
        }
        
        private void ExtractSkillsFromParagraphs(string paragraphs, Resume resume, HashSet<string> foundSkills)
        {
            if (string.IsNullOrWhiteSpace(paragraphs))
                return;
            
            paragraphs = Regex.Replace(paragraphs, @"[\u200B-\u200D\uFEFF]", "");
            
            var lines = paragraphs.Split(new[] { '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries);
            
            foreach (var line in lines)
            {
                var trimmedLine = line.Trim();
                if (string.IsNullOrWhiteSpace(trimmedLine))
                    continue;
                
                // 从段落中提取技能关键词
                // 1. 提取括号中的技能（如：Excel、SPSS、Tableau）
                var bracketPattern = @"[（(]([^）)]+)[）)]";
                var bracketMatches = Regex.Matches(trimmedLine, bracketPattern);
                foreach (Match match in bracketMatches)
                {
                    if (match.Success && match.Groups.Count > 1)
                    {
                        var bracketContent = match.Groups[1].Value;
                        // 按分隔符分割括号内的内容
                        var skillsInBracket = bracketContent.Split(new[] { '、', '，', ',', '；', ';', '|', ' ' }, 
                            StringSplitOptions.RemoveEmptyEntries);
                        foreach (var skill in skillsInBracket)
                        {
                            var trimmedSkill = skill.Trim();
                            if (IsValidSkill(trimmedSkill) && !foundSkills.Contains(trimmedSkill.ToLowerInvariant()))
                            {
                                foundSkills.Add(trimmedSkill.ToLowerInvariant());
                                resume.Skills.Add(trimmedSkill);
                            }
                        }
                    }
                }
                
                // 2. 提取常见的技能关键词（如：Photoshop、Canva、Excel等）
                var commonSkillPatterns = new[]
                {
                    @"(?:熟练|精通|掌握|熟悉|了解|擅长|具备|使用|操作)\s*([A-Za-z][A-Za-z0-9\s]+?)(?:[，,。.、；;]|$)",
                    @"([A-Za-z][A-Za-z0-9]+(?:\s+[A-Za-z][A-Za-z0-9]+)*)", // 英文技能名称
                };
                
                foreach (var pattern in commonSkillPatterns)
                {
                    var matches = Regex.Matches(trimmedLine, pattern);
                    foreach (Match match in matches)
                    {
                        if (match.Success && match.Groups.Count > 1)
                        {
                            var skill = match.Groups[1].Value.Trim();
                            // 清理修饰词
                            skill = Regex.Replace(skill, @"^(熟练|精通|掌握|熟悉|了解|擅长|具备|使用|操作)\s*", "", RegexOptions.IgnoreCase);
                            skill = skill.Trim();
                            
                            if (IsValidSkill(skill) && !foundSkills.Contains(skill.ToLowerInvariant()))
                            {
                                foundSkills.Add(skill.ToLowerInvariant());
                                resume.Skills.Add(skill);
                            }
                        }
                    }
                }
                
                // 3. 提取中文技能关键词（如：市场营销、数据分析、项目管理等）
                // 3.1 提取带后缀的技能（如：数据分析工具、项目管理能力）
                var chineseSkillPattern1 = @"(?:熟练|精通|掌握|熟悉|了解|擅长|具备)\s*([\u4e00-\u9fa5]{2,10}(?:能力|技能|工具|软件|平台|系统|方法|技巧|经验))";
                var chineseMatches1 = Regex.Matches(trimmedLine, chineseSkillPattern1);
                foreach (Match match in chineseMatches1)
                {
                    if (match.Success && match.Groups.Count > 1)
                    {
                        var skill = match.Groups[1].Value.Trim();
                        if (IsValidSkill(skill) && !foundSkills.Contains(skill.ToLowerInvariant()))
                        {
                            foundSkills.Add(skill.ToLowerInvariant());
                            resume.Skills.Add(skill);
                        }
                    }
                }
                
                // 3.2 提取描述中的技能关键词（如：市场营销策划、品牌推广、活动执行）
                // 匹配格式：修饰词 + 技能关键词（2-8个中文字符）
                var chineseSkillPattern2 = @"(?:熟练|精通|掌握|熟悉|了解|擅长|具备|拥有|可)\s*([\u4e00-\u9fa5]{2,8}(?:策划|推广|执行|运营|管理|分析|设计|开发|创作|营销|推广|运营|管理|分析|设计|开发|创作))";
                var chineseMatches2 = Regex.Matches(trimmedLine, chineseSkillPattern2);
                foreach (Match match in chineseMatches2)
                {
                    if (match.Success && match.Groups.Count > 1)
                    {
                        var skill = match.Groups[1].Value.Trim();
                        if (IsValidSkill(skill) && !foundSkills.Contains(skill.ToLowerInvariant()))
                        {
                            foundSkills.Add(skill.ToLowerInvariant());
                            resume.Skills.Add(skill);
                        }
                    }
                }
                
                // 3.3 提取用顿号、逗号分隔的中文技能列表
                // 匹配格式：技能1、技能2、技能3
                var chineseListPattern = @"([\u4e00-\u9fa5]{2,8}(?:策划|推广|执行|运营|管理|分析|设计|开发|创作|营销|推广|运营|管理|分析|设计|开发|创作|平台|工具|软件|系统|方法|技巧|经验))(?:[、，,])";
                var chineseListMatches = Regex.Matches(trimmedLine, chineseListPattern);
                foreach (Match match in chineseListMatches)
                {
                    if (match.Success && match.Groups.Count > 1)
                    {
                        var skill = match.Groups[1].Value.Trim();
                        if (IsValidSkill(skill) && !foundSkills.Contains(skill.ToLowerInvariant()))
                        {
                            foundSkills.Add(skill.ToLowerInvariant());
                            resume.Skills.Add(skill);
                        }
                    }
                }
            }
        }
        
        // 验证是否为有效技能
        private bool IsValidSkill(string skill)
        {
            if (string.IsNullOrWhiteSpace(skill))
                return false;
            
            skill = skill.Trim();
            
            // 长度检查
            if (skill.Length < 2 || skill.Length > 30)
                return false;
            
            // 排除明显不是技能的内容
            var excludeKeywords = new[] { 
                "姓名", "性别", "地址", "手机", "电话", "邮箱", "身份证", 
                "学历", "工作", "经历", "教育", "学校", "专业", "公司",
                "文件信息", "导入时间", "目录", "未归类", "能够", "可以", "完成"
            };
            
            foreach (var keyword in excludeKeywords)
            {
                if (skill.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                    return false;
            }
            
            // 排除纯数字
            if (Regex.IsMatch(skill, @"^\d+$"))
                return false;
            
            return true;
        }
        
        // 从文本中提取技能列表
        private void ExtractSkillsFromText(string skillText, Resume resume, HashSet<string> foundSkills)
        {
            if (string.IsNullOrWhiteSpace(skillText))
                return;
            
            // 清理零宽字符和其他特殊字符
            skillText = Regex.Replace(skillText, @"[\u200B-\u200D\uFEFF]", "");
            skillText = Regex.Replace(skillText, @"[\x00-\x1F\x7F]", "");
            
            // 按多种分隔符分割技能
            var separators = new[] { '、', '，', ',', '；', ';', '\n', '\r', '|', ' ' };
            var skills = skillText.Split(separators, StringSplitOptions.RemoveEmptyEntries);
            
            foreach (var skill in skills)
            {
                var trimmedSkill = skill.Trim();
                
                // 排除明显不是技能的内容
                if (trimmedSkill.Contains("姓名") || trimmedSkill.Contains("性别") || 
                    trimmedSkill.Contains("地址") || trimmedSkill.Contains("手机") ||
                    trimmedSkill.Contains("邮箱") || trimmedSkill.Contains("身份证") ||
                    trimmedSkill.Contains("学历") || trimmedSkill.Contains("工作"))
                {
                    continue;
                }
                
                // 技能长度应该在2-50个字符之间
                if (!string.IsNullOrEmpty(trimmedSkill) && 
                    trimmedSkill.Length >= 2 && 
                    trimmedSkill.Length <= 50 &&
                    !foundSkills.Contains(trimmedSkill))
                {
                    foundSkills.Add(trimmedSkill);
                    resume.Skills.Add(trimmedSkill);
                }
            }
        }

        // 判断是否为中文姓名
        private bool IsChineseName(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;
            
            // 清理零宽字符
            text = Regex.Replace(text, @"[\u200B-\u200D\uFEFF]", "");
            text = text.Trim();
            
            if (text.Length < 2 || text.Length > 4)
                return false;
            
            // 排除明显不是姓名的内容
            var excludeKeywords = new[] { "：", ":", "公司", "学校", "地址", "电话", "邮箱", "手机", 
                                         "性别", "出生", "日期", "简历", "姓名", "专业", "学历" };
            foreach (var keyword in excludeKeywords)
            {
                if (text.Contains(keyword))
                    return false;
            }
            
            // 不能包含数字
            if (text.Any(char.IsDigit))
                return false;
            
            // 检查是否主要是中文字符
            int chineseCharCount = 0;
            int totalCharCount = 0;
            
            foreach (var c in text)
            {
                // 跳过空格和常见标点
                if (char.IsWhiteSpace(c) || c == '·' || c == '-')
                    continue;
                    
                totalCharCount++;
                
                // 中文字符范围：\u4E00-\u9FFF
                if (c >= 0x4E00 && c <= 0x9FFF)
                {
                    chineseCharCount++;
                }
                else
                {
                    // 如果包含非中文字符（除了空格和标点），则不是中文姓名
                    return false;
                }
            }
            
            // 至少包含2个中文字符，且中文字符占比超过80%
            return chineseCharCount >= 2 && totalCharCount > 0 && 
                   (chineseCharCount * 100 / totalCharCount) >= 80;
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
                    return ReadDocContent(filePath);
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

                    return CleanContentFormat(sb.ToString());
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"读取DOCX文件失败: {ex.Message}");
            }
        }

        // 读取.doc文件内容
        private string ReadDocContent(string filePath)
        {
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var poifs = new POIFSFileSystem(fs);
                var hwpfDocument = new HWPFDocument(poifs);
                
                var sb = new StringBuilder();
                
                // 按段落提取文本，保持换行格式
                var range = hwpfDocument.GetRange();
                var numParagraphs = range.NumParagraphs;
                
                for (int i = 0; i < numParagraphs; i++)
                {
                    var paragraph = range.GetParagraph(i);
                    var text = paragraph.Text;
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        sb.AppendLine(text.Trim());
                    }
                }
                
                if (sb.Length == 0)
                {
                    var fullText = range.Text;
                    if (!string.IsNullOrWhiteSpace(fullText))
                    {
                        // 将整体文本按行分割，保持格式
                        var lines = fullText.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var line in lines)
                        {
                            var trimmedLine = line.Trim();
                            if (!string.IsNullOrWhiteSpace(trimmedLine))
                            {
                                sb.AppendLine(trimmedLine);
                            }
                        }
                    }
                }
                return CleanContentFormat(sb.ToString());
            }
        }

        private string CleanContentFormat(string content)
        {
            content = content.Replace("\r\n", "\n").Replace("\r", "\n");
            return content.Trim();
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

