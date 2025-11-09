using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using 页面.Models;

namespace 页面.Services
{
    public class AiReviewService
    {
        private readonly HttpClient _httpClient;
        private readonly string _apiKey;
        private readonly string _baseUrl;
        private readonly string _model;

        public AiReviewService(HttpClient httpClient, string apiKey, string baseUrl, string model)
        {
            _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _apiKey = !string.IsNullOrWhiteSpace(apiKey) ? apiKey : throw new ArgumentException("API Key 不能为空", nameof(apiKey));
            _baseUrl = string.IsNullOrWhiteSpace(baseUrl) ? "https://openrouter.ai/api/v1" : baseUrl.TrimEnd('/');
            _model = string.IsNullOrWhiteSpace(model) ? "Qwen/Qwen3-Omni-30B-A3B-Instruct" : model;

            _httpClient.Timeout = TimeSpan.FromSeconds(60);
        }

        public async Task<string> GenerateReviewAsync(Resume resume)
        {
            if (resume == null)
            {
                throw new ArgumentNullException(nameof(resume));
            }

            var prompt = BuildPrompt(resume);

            // 默认按 OpenAI / OpenRouter 兼容的 chat.completions 接口
            var url = _baseUrl + "/chat/completions";
            using var request = new HttpRequestMessage(HttpMethod.Post, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);

            var requestBody = new
            {
                model = _model,
                messages = new object[]
                {
                    new { role = "system", content = "你是一名资深HR，请基于候选人的简历信息，输出客观、专业、可操作的点评，包括总体评价、核心优势、潜在风险、建议面试问题和改进建议。" },
                    new { role = "user", content = prompt }
                },
                temperature = 0.3,
                stream = false
            };

            var json = JsonSerializer.Serialize(requestBody);
            request.Content = new StringContent(json, Encoding.UTF8, "application/json");

            using var response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(false);
            response.EnsureSuccessStatusCode();

            var respStr = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            return ParseChatCompletion(respStr);
        }

        public async Task<(int score, string[] keywords)> ScoreSemanticMatchAsync(string jobDescription, string resumeSummary)
        {
            if (string.IsNullOrWhiteSpace(jobDescription)) throw new ArgumentException("职位描述不能为空", nameof(jobDescription));
            if (string.IsNullOrWhiteSpace(resumeSummary)) throw new ArgumentException("简历内容为空", nameof(resumeSummary));

            var system = "你是一个简历与岗位描述的匹配评估助手。请严格输出JSON，字段：score(0-100的整数), keywords(与岗位高度匹配的3-10个短语，按重要性降序)。不要输出多余文本。";
            var user = new StringBuilder();
            user.AppendLine("岗位描述：");
            user.AppendLine(jobDescription.Trim());
            user.AppendLine();
            user.AppendLine("候选人摘要：");
            user.AppendLine(resumeSummary.Trim());

            var url = _baseUrl + "/chat/completions";
            using var request = new HttpRequestMessage(HttpMethod.Post, url);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);

            var body = new
            {
                model = _model,
                messages = new object[]
                {
                    new { role = "system", content = system },
                    new { role = "user", content = user.ToString() }
                },
                temperature = 0.2,
                stream = false
            };

            request.Content = new StringContent(JsonSerializer.Serialize(body), Encoding.UTF8, "application/json");
            using var response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(false);
            var respStr = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
            if (!response.IsSuccessStatusCode)
            {
                throw new Exception($"HTTP {(int)response.StatusCode} {response.StatusCode}: {respStr}");
            }

            var content = ParseChatCompletion(respStr);

            try
            {
                using var doc = JsonDocument.Parse(content);
                var root = doc.RootElement;
                int score = root.TryGetProperty("score", out var s) && s.TryGetInt32(out var sv) ? Math.Clamp(sv, 0, 100) : 0;
                string[] keywords = Array.Empty<string>();
                if (root.TryGetProperty("keywords", out var ks) && ks.ValueKind == JsonValueKind.Array)
                {
                    var list = new System.Collections.Generic.List<string>();
                    foreach (var k in ks.EnumerateArray())
                    {
                        var v = k.GetString();
                        if (!string.IsNullOrWhiteSpace(v)) list.Add(v.Trim());
                    }
                    keywords = list.ToArray();
                }
                return (score, keywords);
            }
            catch
            {
                // 回退：若非JSON，尝试简单解析数字
                var fallback = 0;
                foreach (var ch in content)
                {
                    if (char.IsDigit(ch)) { fallback = Math.Min(100, fallback * 10 + (ch - '0')); }
                    else if (fallback > 0) break;
                }
                return (fallback, Array.Empty<string>());
            }
        }

        private static string BuildPrompt(Resume resume)
        {
            var sb = new StringBuilder();
            sb.AppendLine("候选人简历信息：");
            sb.AppendLine($"姓名：{resume.Name}");
            if (!string.IsNullOrWhiteSpace(resume.Gender)) sb.AppendLine($"性别：{resume.Gender}");
            if (resume.BirthDate.HasValue) sb.AppendLine($"出生日期：{resume.BirthDate:yyyy-MM-dd}");
            if (!string.IsNullOrWhiteSpace(resume.Address)) sb.AppendLine($"地址：{resume.Address}");
            if (!string.IsNullOrWhiteSpace(resume.Phone)) sb.AppendLine($"手机：{resume.Phone}");
            if (!string.IsNullOrWhiteSpace(resume.Email)) sb.AppendLine($"邮箱：{resume.Email}");
            if (!string.IsNullOrWhiteSpace(resume.FirstEducationSchool)) sb.AppendLine($"第一学历学校：{resume.FirstEducationSchool}");
            if (!string.IsNullOrWhiteSpace(resume.FirstEducationMajor)) sb.AppendLine($"第一学历专业：{resume.FirstEducationMajor}");
            if (!string.IsNullOrWhiteSpace(resume.HighestEducationSchool)) sb.AppendLine($"最高学历学校：{resume.HighestEducationSchool}");
            if (!string.IsNullOrWhiteSpace(resume.HighestEducationMajor)) sb.AppendLine($"最高学历专业：{resume.HighestEducationMajor}");

            if (resume.Skills != null && resume.Skills.Count > 0)
            {
                sb.AppendLine("技能：" + string.Join("，", resume.Skills));
            }

            if (resume.WorkExperiences != null && resume.WorkExperiences.Count > 0)
            {
                sb.AppendLine("工作经历：");
                foreach (var w in resume.WorkExperiences)
                {
                    var line = new StringBuilder();
                    if (!string.IsNullOrWhiteSpace(w.Company)) line.Append(w.Company + " ");
                    if (!string.IsNullOrWhiteSpace(w.Position)) line.Append(w.Position + " ");
                    if (!string.IsNullOrWhiteSpace(w.Title)) line.Append(w.Title + " ");
                    if (!string.IsNullOrWhiteSpace(w.Project)) line.Append("项目:" + w.Project + " ");
                    if (!string.IsNullOrWhiteSpace(w.Responsibilities)) line.Append("职责:" + w.Responsibilities);
                    sb.AppendLine("- " + line.ToString().Trim());
                }
            }

            sb.AppendLine();
            sb.AppendLine("请按如下结构输出：\n1) 总体评价\n2) 核心优势（3-5条）\n3) 潜在风险（2-4条）\n4) 建议面试问题（4-6个）\n5) 改进建议（2-4条）\n用中文输出，控制在400字。");
            return sb.ToString();
        }

        private static string ParseChatCompletion(string json)
        {
            try
            {
                using var doc = JsonDocument.Parse(json);
                var root = doc.RootElement;

                // 兼容 OpenAI / OpenRouter 风格：choices[0].message.content
                if (root.TryGetProperty("choices", out var choices) && choices.ValueKind == JsonValueKind.Array && choices.GetArrayLength() > 0)
                {
                    var first = choices[0];
                    if (first.TryGetProperty("message", out var message) && message.TryGetProperty("content", out var content))
                    {
                        return content.GetString() ?? string.Empty;
                    }

                    // 兼容一些返回格式：choices[0].text
                    if (first.TryGetProperty("text", out var text))
                    {
                        return text.GetString() ?? string.Empty;
                    }
                }
            }
            catch
            {
                // 忽略解析异常，回退到原始字符串
            }

            return json; // 解析失败时返回原始JSON，便于诊断
        }
    }
}


