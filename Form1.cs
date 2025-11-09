using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using 页面.Models;
using 页面.Services;

namespace 页面
{
    public partial class Form1 : Form
    {
        private readonly ResumeManagementService _managementService;
        private readonly AiReviewService _aiReviewService;
        private readonly string _appDataDirectory;
        private readonly string _reviewDataFile;
        private List<Resume> _currentResumes = new List<Resume>();
        private List<ResumeDirectory> _directories = new List<ResumeDirectory>();
        private List<Resume> _reviewResumes = new List<Resume>();
        private List<Resume> _exportResumes = new List<Resume>();
        private Dictionary<string, string> _reviewNotes = new Dictionary<string, string>();

        public Form1()
        {
            InitializeComponent();
            _managementService = new ResumeManagementService();
            _aiReviewService = InitializeAiServiceOrNull();
            _appDataDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "智能简历解析系统");
            if (!Directory.Exists(_appDataDirectory))
            {
                Directory.CreateDirectory(_appDataDirectory);
            }
            _reviewDataFile = Path.Combine(_appDataDirectory, "resume_reviews.json");
            LoadReviewNotes();

            InitializeUI();
            panelAdd.Show();
        }

        private AiReviewService InitializeAiServiceOrNull()
        {
            try
            {
                // 从环境变量读取配置，便于部署时只设置密钥即可
                string apiKey = Environment.GetEnvironmentVariable("QWEN_API_KEY")
                                   ?? Environment.GetEnvironmentVariable("OPENROUTER_API_KEY");
                string baseUrl = Environment.GetEnvironmentVariable("AI_BASE_URL") ?? "https://openrouter.ai/api/v1";
                string model = Environment.GetEnvironmentVariable("AI_MODEL") ?? "Qwen/Qwen3-Omni-30B-A3B-Instruct";

                if (string.IsNullOrWhiteSpace(apiKey))
                {
                    return null; // 未配置密钥，延迟到点击按钮时提示
                }

                var httpClient = new System.Net.Http.HttpClient();
                return new AiReviewService(httpClient, apiKey, baseUrl, model);
            }
            catch
            {
                return null;
            }
        }

        private void InitializeUI()
        {
            // 导入按钮
            var importButton = new Button
            {
                Text = "选择文件导入",
                Location = new Point(10, 10),
                Size = new Size(150, 40)
            };
            importButton.Click += ImportButton_Click;
            panelAdd.Controls.Add(importButton);

            // 拖拽提示标签
            var dragLabel = new Label
            {
                Text = "或者将文件拖拽到此区域",
                Location = new Point(10, 60),
                Size = new Size(200, 30)
            };
            panelAdd.Controls.Add(dragLabel);

            // 目录管理区域
            var directoryLabel = new Label
            {
                Text = "目录:",
                Location = new Point(210, 65),
                Size = new Size(50, 20)
            };
            panelAdd.Controls.Add(directoryLabel);

            // 目录选择下拉框
            var directoryCombo = new ComboBox
            {
                Name = "directoryCombo",
                Location = new Point(260, 60),
                Size = new Size(120, 25),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            directoryCombo.SelectedIndexChanged += DirectoryCombo_SelectedIndexChanged;
            panelAdd.Controls.Add(directoryCombo);

            // 新建目录按钮
            var newDirButton = new Button
            {
                Text = "新建目录",
                Location = new Point(390, 60),
                Size = new Size(100, 30)
            };
            newDirButton.Click += NewDirectoryButton_Click;
            panelAdd.Controls.Add(newDirButton);

            // 添加目录删除按钮
            var deleteDirButton = new Button
            {
                Text = "删除目录",
                Location = new Point(500, 60),  // 位置在新建目录按钮右侧
                Size = new Size(100, 30)
            };
            deleteDirButton.Click += DeleteDirectoryButton_Click;  // 绑定之前实现的点击事件
            panelAdd.Controls.Add(deleteDirButton);

            // 已上传简历标签
            var uploadedLabel = new Label
            {
                Text = "已上传的简历:",
                Location = new Point(10, 100),
                Size = new Size(150, 30),
                Font = new Font("微软雅黑", 9, FontStyle.Bold)
            };
            panelAdd.Controls.Add(uploadedLabel);

            // 已上传简历列表
            var uploadedListBox = new ListBox
            {
                Name = "uploadedListBox",
                Location = new Point(10, 135),
                Size = new Size(560, 190),
                SelectionMode = SelectionMode.One
            };
            uploadedListBox.DoubleClick += UploadedListBox_DoubleClick;
            panelAdd.Controls.Add(uploadedListBox);

            // 移动到目录按钮
            var moveToDirButton = new Button
            {
                Text = "移动到目录",
                Location = new Point(250, 330),
                Size = new Size(120, 35),
                BackColor = Color.Orange,
                ForeColor = Color.White
            };
            moveToDirButton.Click += MoveToDirectoryButton_Click;
            panelAdd.Controls.Add(moveToDirButton);

            // 删除按钮
            var deleteButton = new Button
            {
                Text = "删除所选",
                Location = new Point(10, 330),
                Size = new Size(120, 35),
                BackColor = Color.Red,
                ForeColor = Color.White
            };
            deleteButton.Click += DeleteResumeButton_Click;
            panelAdd.Controls.Add(deleteButton);

            // 查看详情按钮
            var viewDetailButton = new Button
            {
                Text = "查看详情",
                Location = new Point(140, 330),
                Size = new Size(100, 35),
                BackColor = Color.Green,
                ForeColor = Color.White
            };
            viewDetailButton.Click += ViewDetailButton_Click;
            panelAdd.Controls.Add(viewDetailButton);

            // 设置拖拽功能
            panelAdd.AllowDrop = true;
            panelAdd.DragEnter += PanelAdd_DragEnter;
            panelAdd.DragDrop += PanelAdd_DragDrop;

            // 初始化简历检索面板
            InitializeSearchPanel();

            // 初始化简历查重面板
            InitializeDuplicatePanel();

            // 初始化简历点评面板
            InitializeReviewPanel();

            // 初始化数据统计面板
            InitializeStatisticsPanel();

            // 初始化简历导出面板
            InitializeExportPanel();

            // 初始化语义匹配面板
            InitializeMatchPanel();

            // 加载目录数据
            LoadDirectories();

        }

        // 需添加的辅助方法
        private void LoadDirectories()
        {
            // 加载目录数据到下拉框
            _directories = _managementService.GetAllDirectories(); // 需在ResumeManagementService中实现
            var directoryCombo = panelAdd.Controls.Find("directoryCombo", false).FirstOrDefault() as ComboBox;
            if (directoryCombo != null)
            {
                directoryCombo.Items.Clear();
                directoryCombo.Items.Add("所有目录"); // 默认选项
                directoryCombo.Items.Add("未归类");
                foreach (var dir in _directories)
                {
                    directoryCombo.Items.Add(dir.Name);
                }
                directoryCombo.SelectedIndex = 0;
            }
        }

        private void NewDirectoryButton_Click(object sender, EventArgs e)
        {
            // 新建目录逻辑
            var inputForm = new Form
            {
                Text = "新建目录",
                Size = new Size(300, 150),
                StartPosition = FormStartPosition.CenterParent
            };

            var label = new Label
            {
                Text = "目录名称:",
                Location = new Point(20, 20),
                Size = new Size(80, 20)
            };
            inputForm.Controls.Add(label);

            var textBox = new TextBox
            {
                Location = new Point(100, 20),
                Size = new Size(150, 20)
            };
            inputForm.Controls.Add(textBox);

            var confirmButton = new Button
            {
                Text = "确认",
                Location = new Point(80, 60),
                Size = new Size(70, 30)
            };
            confirmButton.Click += (s, args) =>
            {
                string dirName = textBox.Text.Trim();
                if (!string.IsNullOrEmpty(dirName))
                {
                    _managementService.CreateDirectory(dirName); // 需在ResumeManagementService中实现
                    LoadDirectories(); // 刷新目录列表
                    inputForm.Close();
                }
                else
                {
                    MessageBox.Show("目录名称不能为空", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            };
            inputForm.Controls.Add(confirmButton);

            var cancelButton = new Button
            {
                Text = "取消",
                Location = new Point(160, 60),
                Size = new Size(70, 30)
            };
            cancelButton.Click += (s, args) => inputForm.Close();
            inputForm.Controls.Add(cancelButton);

            inputForm.ShowDialog();
        }

        // 实现选择事件：筛选并刷新简历列表
        private void DirectoryCombo_SelectedIndexChanged(object sender, EventArgs e)
        {
            var combo = sender as ComboBox;
            if (combo == null) return;

            // 启用/禁用“删除目录”按钮（仅自定义目录可删）
            var deleteDirBtn = panelAdd.Controls.Find("deleteDirButton", true).FirstOrDefault() as Button;
            if (deleteDirBtn != null)
            {
                deleteDirBtn.Enabled = combo.SelectedIndex > 1; // 索引>1是自定义目录
            }

            // 筛选对应目录的简历
            if (combo.SelectedIndex == 0)
            {
                // 所有目录：包含未归类（unclassified）和所有自定义目录的简历
                _currentResumes = _managementService.LoadResumes();
            }
            else if (combo.SelectedIndex == 1)
            {
                // 未归类：仅显示DirectoryId为"unclassified"的简历
                _currentResumes = _managementService.LoadResumes()
                    .Where(r => r.DirectoryId == "unclassified")
                    .ToList();
            }
            else
            {
                // 自定义目录：仅显示DirectoryId等于该目录ID的简历
                string selectedDirName = combo.SelectedItem.ToString();
                var targetDir = _directories.FirstOrDefault(d => d.Name == selectedDirName);
                if (targetDir != null)
                {
                    _currentResumes = _managementService.LoadResumes()
                        .Where(r => r.DirectoryId == targetDir.Id)
                        .ToList();
                }
            }

            // 刷新简历列表
            UpdateUploadedListBox();
        }

        // 辅助方法：更新简历列表显示
        private void UpdateUploadedListBox()
        {
            var uploadedListBox = panelAdd.Controls.Find("uploadedListBox", true).FirstOrDefault() as ListBox;
            if (uploadedListBox == null) return;

            uploadedListBox.Items.Clear();
            foreach (var resume in _currentResumes.OrderByDescending(r => r.ImportTime))
            {
                uploadedListBox.Items.Add($"{resume.Name} - {resume.FileName} - {resume.ImportTime:yyyy-MM-dd HH:mm}");
            }
        }

        private void MoveToDirectoryButton_Click(object sender, EventArgs e)
        {
            // 获取当前选中的简历
            var uploadedListBox = panelAdd.Controls.Find("uploadedListBox", false).FirstOrDefault() as ListBox;
            var directoryCombo = panelAdd.Controls.Find("directoryCombo", false).FirstOrDefault() as ComboBox;
            if (uploadedListBox == null)
            {
                MessageBox.Show("未找到简历列表控件", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (uploadedListBox.SelectedIndex < 0)
            {
                MessageBox.Show("请先选择要移动的简历", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 获取所有可用目录（排除系统目录）
            var allDirectories = _managementService.GetAllDirectories()
                .Where(d => !string.IsNullOrEmpty(d.Id) && !string.IsNullOrEmpty(d.Name))
                .ToList();

            if (allDirectories.Count == 0)
            {
                MessageBox.Show("没有可用的目录，请先创建目录", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 创建目录选择对话框
            using (var selectForm = new Form())
            {
                selectForm.Text = "选择目标目录";
                selectForm.Size = new Size(300, 400);
                selectForm.StartPosition = FormStartPosition.CenterParent;

                // 添加目录列表
                var dirListBox = new ListBox
                {
                    Location = new Point(20, 20),
                    Size = new Size(240, 280),
                    SelectionMode = SelectionMode.One
                };
                // 添加未归类选项
                dirListBox.Items.Add("未归类");
                // 添加自定义目录
                foreach (var dir in allDirectories)
                {
                    dirListBox.Items.Add(dir.Name);
                }
                dirListBox.SelectedIndex = 0; // 默认选中未归类
                selectForm.Controls.Add(dirListBox);

                // 确认按钮
                var confirmBtn = new Button
                {
                    Text = "确认",
                    Location = new Point(50, 320),
                    Size = new Size(80, 30)
                };
                confirmBtn.DialogResult = DialogResult.OK;
                selectForm.Controls.Add(confirmBtn);

                // 取消按钮
                var cancelBtn = new Button
                {
                    Text = "取消",
                    Location = new Point(160, 320),
                    Size = new Size(80, 30)
                };
                cancelBtn.DialogResult = DialogResult.Cancel;
                selectForm.Controls.Add(cancelBtn);

                // 显示对话框并处理结果
                if (selectForm.ShowDialog() == DialogResult.OK)
                {
                    string targetDirName = dirListBox.SelectedItem.ToString();
                    string targetDirId;

                    // 处理未归类目录
                    if (targetDirName == "未归类")
                    {
                        targetDirId = "unclassified";
                    }
                    else
                    {
                        // 查找选中的自定义目录ID
                        var targetDir = allDirectories.FirstOrDefault(d => d.Name == targetDirName);
                        if (targetDir == null)
                        {
                            MessageBox.Show("选中的目录不存在", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                        targetDirId = targetDir.Id;
                    }

                    // 执行移动操作
                    string selectedItemText = uploadedListBox.SelectedItem.ToString();
                    var selectedResume = _currentResumes.FirstOrDefault(r =>
                        $"{r.Name} - {r.FileName} - {r.ImportTime:yyyy-MM-dd HH:mm}" == selectedItemText
                    );

                    if (selectedResume == null || string.IsNullOrWhiteSpace(selectedResume.Id))
                    {
                        MessageBox.Show("选中的简历不存在或数据异常", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    // 记录当前选中的目录索引，用于刷新后恢复选择
                    int currentDirIndex = directoryCombo.SelectedIndex;

                    // 检查是否已在目标目录
                    if (selectedResume.DirectoryId == targetDirId)
                    {
                        MessageBox.Show("简历已在目标目录中，无需重复移动", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    // 执行移动
                    bool success = _managementService.MoveResumeToDirectory(selectedResume.Id, targetDirId);
                    if (success)
                    {
                        // 关键修复：先刷新所有简历数据
                        _currentResumes = _managementService.LoadResumes();

                        // 重新应用当前目录筛选，确保移动的简历从原目录消失
                        DirectoryCombo_SelectedIndexChanged(directoryCombo, EventArgs.Empty);

                        // 恢复目录选择状态
                        directoryCombo.SelectedIndex = currentDirIndex;

                        MessageBox.Show("简历已移动到指定目录", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("移动简历失败，请重试", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        // Form1.cs 添加删除目录按钮事件
        private void DeleteDirectoryButton_Click(object sender, EventArgs e)
        {
            var directoryCombo = panelAdd.Controls.Find("directoryCombo", false).FirstOrDefault() as ComboBox;
            if (directoryCombo == null || directoryCombo.SelectedIndex <= 0)
            {
                MessageBox.Show("请选择要删除的目录", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string targetDirName = directoryCombo.SelectedItem.ToString();
            var latestDirectories = _managementService.GetAllDirectories();
            var targetDir = latestDirectories.FirstOrDefault(d => d.Name == targetDirName);
            if (targetDir == null) return;
            if (targetDir == null)
            {
                MessageBox.Show("所选目录不存在或已被删除", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            // 禁止删除系统默认目录
            if (targetDir.Id == "unclassified")
            {
                MessageBox.Show("系统默认目录不可删除", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 额外检查目录ID是否有效
            if (string.IsNullOrWhiteSpace(targetDir.Id))
            {
                MessageBox.Show("目录数据异常，无法删除", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                var result = MessageBox.Show(
                    $"确定要删除目录 '{targetDirName}' 吗？",
                    "确认删除",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (result != DialogResult.Yes)
                {
                    return; // 用户取消操作
                }

                // 检查并移动目录下的简历
                var allResumes = _managementService.LoadResumes();
                var residualResumes = allResumes.Where(r => r.DirectoryId == targetDir.Id).ToList();

                if (residualResumes.Count > 0)
                {
                    // 批量移动简历到未归类目录
                    int successCount = 0;
                    foreach (var resume in residualResumes)
                    {
                        if (_managementService.MoveResumeToDirectory(resume.Id, "unclassified"))
                        {
                            successCount++;
                        }
                    }

                    if (successCount > 0)
                    {
                        MessageBox.Show(
                            $"已将 {successCount}/{residualResumes.Count} 份简历移动到'未归类'目录",
                            "移动完成",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                }

                // 尝试删除目录
                bool deleteSuccess = _managementService.DeleteDirectory(targetDir.Id);
                if (deleteSuccess)
                {
                    MessageBox.Show("目录删除成功", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadDirectories(); // 刷新目录列表
                    LoadResumes();     // 刷新简历列表
                }
                else
                {
                    MessageBox.Show("目录删除失败，请重试", "失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show(ex.Message, "操作失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"删除失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadResumes();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ShowPanel(panelAdd);
            LoadResumes();  // 切换到添加面板时刷新简历列表
        }

        private void buttonSearchResume_Click(object sender, EventArgs e)
        {
            ShowPanel(panelSearch);
        }

        private void buttonCheckDuplicate_Click(object sender, EventArgs e)
        {
            ShowPanel(panelDuplicate);
        }

        private void buttonReview_Click(object sender, EventArgs e)
        {
            ShowPanel(panelReview);
            LoadReviewResumes();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            ShowPanel(panelStatistics);
            LoadStatistics();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ShowPanel(panelExport);
            LoadExportResumes();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ShowPanel(panelMatch);
            LoadMatchJobs();

        }

        private void ShowPanel(Panel panelToShow)
        {
            panelAdd.Hide();
            panelSearch.Hide();
            panelDuplicate.Hide();
            panelReview.Hide();
            panelStatistics.Hide();
            panelExport.Hide();
            panelMatch.Hide();
            panelToShow.Show();
        }

        private void ImportButton_Click(object sender, EventArgs e)
        {
            using (var openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "简历文件|*.doc;*.docx;*.pdf";
                openFileDialog.Multiselect = true;
                openFileDialog.Title = "选择要导入的简历文件";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string targetDirectoryId = null;
                    var directoryCombo = panelAdd.Controls.Find("directoryCombo", false).FirstOrDefault() as ComboBox;
                    if (directoryCombo != null && directoryCombo.SelectedIndex > 0)
                    {
                        string selectedDirName = directoryCombo.SelectedItem.ToString();
                        // 从已加载的目录列表中匹配目录ID
                        var selectedDir = _directories.FirstOrDefault(d => d.Name == selectedDirName);
                        targetDirectoryId = selectedDir?.Id; // 得到目标目录ID
                    }

                    // 2. 调用导入方法时传入目录ID（需修改ImportFiles方法接收该参数）
                    ImportFiles(openFileDialog.FileNames, targetDirectoryId);
                }
            }
        }

        private void PanelAdd_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void PanelAdd_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                var validFiles = files.Where(f =>
                    Path.GetExtension(f).ToLower() == ".doc" ||
                    Path.GetExtension(f).ToLower() == ".docx" ||
                    Path.GetExtension(f).ToLower() == ".pdf").ToArray();

                if (validFiles.Length > 0)
                {
                    string targetDirectoryId = null;
                    var directoryCombo = panelAdd.Controls.Find("directoryCombo", false).FirstOrDefault() as ComboBox;
                    if (directoryCombo != null && directoryCombo.SelectedIndex > 0)
                    {
                        string selectedDirName = directoryCombo.SelectedItem.ToString();
                        var selectedDir = _directories.FirstOrDefault(d => d.Name == selectedDirName);
                        targetDirectoryId = selectedDir?.Id;
                    }

                    ImportFiles(validFiles, targetDirectoryId);
                }
                else
                {
                    MessageBox.Show("请选择Word或PDF格式的简历文件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void ImportFiles(string[] filePaths, string targetDirectoryId)
        {
            try
            {

                ImportResult importResult = _managementService.ImportResumes(filePaths.ToList(), false);

                if (importResult.ImportedResumes.Count > 0)
                {
                    // 1. 单独弹出“新增成功”提示（仅当有新增时显示）
                    if (importResult.NewCount > 0)
                    {
                        MessageBox.Show(
                            $"成功新增 {importResult.NewCount} 份简历。",
                            "导入成功",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }

                    // 2. 单独弹出“替换成功”提示（仅当有替换时显示）
                    if (importResult.ReplacedCount > 0)
                    {
                        MessageBox.Show(
                            $"成功替换 {importResult.ReplacedCount} 份简历。",
                            "替换成功",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }

                    LoadResumes(); // 刷新简历列表
                }
                else
                {
                    // 无任何导入时的提示（优化原有文案，更清晰）
                    MessageBox.Show(
                        "未导入任何简历（可能全部跳过或文件格式不支持）",
                        "提示",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导入失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void ResultListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            var listBox = sender as ListBox;
            if (listBox?.SelectedIndex >= 0 && listBox.SelectedIndex < _currentResumes.Count)
            {
                var selectedResume = _currentResumes[listBox.SelectedIndex];
                ShowResumeDetails(selectedResume);
            }
        }

        private void ShowResumeDetails(Resume resume)
        {
            var detailsForm = new Form
            {
                Text = $"简历详情 - {resume.Name}",
                Size = new Size(600, 500),
                StartPosition = FormStartPosition.CenterParent
            };

            var textBox = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                ReadOnly = true,
                Text = GetResumeDetailsText(resume)
            };

            detailsForm.Controls.Add(textBox);
            detailsForm.ShowDialog();
        }

        private string GetResumeDetailsText(Resume resume)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"姓名: {resume.Name}");
            sb.AppendLine($"性别: {resume.Gender}");
            sb.AppendLine($"出生日期: {resume.BirthDate?.ToString("yyyy-MM-dd") ?? ""}");
            sb.AppendLine($"地址: {resume.Address}");
            sb.AppendLine($"手机: {resume.Phone}");
            sb.AppendLine($"邮箱: {resume.Email}");
            sb.AppendLine($"身份证号: {resume.IdCard}");
            sb.AppendLine();
            sb.AppendLine("学历信息");
            sb.AppendLine($"第一学历（本科）学校: {resume.FirstEducationSchool}");
            sb.AppendLine($"第一学历专业: {resume.FirstEducationMajor}");
            sb.AppendLine($"最高学历学校: {resume.HighestEducationSchool}");
            sb.AppendLine($"最高学历专业: {resume.HighestEducationMajor}");
            sb.AppendLine();
            sb.AppendLine("工作经历");
            foreach (var work in resume.WorkExperiences)
            {
                sb.AppendLine($"  - {work.Company} - {work.Position}");
                if (!string.IsNullOrEmpty(work.Responsibilities))
                {
                    sb.AppendLine($"    职责: {work.Responsibilities}");
                }
            }
            sb.AppendLine();
            sb.AppendLine("技能");
            foreach (var skill in resume.Skills)
            {
                sb.AppendLine($"  - {skill}");
            }
            sb.AppendLine();
            sb.AppendLine("文件信息");
            sb.AppendLine($"文件名: {resume.FileName}");
            sb.AppendLine($"导入时间: {resume.ImportTime:yyyy-MM-dd HH:mm:ss}");

            var allDirectories = _managementService.GetAllDirectories(); // 从服务获取所有目录
            string directoryName = allDirectories.FirstOrDefault(d => d.Id == resume.DirectoryId)?.Name ?? "未归类";
            sb.AppendLine($"目录: {directoryName}");

            return sb.ToString();
        }




        private void LoadResumes()
        {
            _currentResumes = _managementService.LoadResumes();
            UpdateUploadedResumesList();
        }

        private void UpdateUploadedResumesList()
        {
            var uploadedListBox = panelAdd.Controls.Find("uploadedListBox", false).FirstOrDefault() as ListBox;
            if (uploadedListBox != null)
            {
                uploadedListBox.Items.Clear();
                foreach (var resume in _currentResumes.OrderByDescending(r => r.ImportTime))
                {
                    uploadedListBox.Items.Add($"{resume.Name} - {resume.FileName} - {resume.ImportTime:yyyy-MM-dd HH:mm}");
                }
            }
        }

        private void UploadedListBox_DoubleClick(object sender, EventArgs e)
        {
            var listBox = sender as ListBox;
            if (listBox?.SelectedIndex >= 0 && listBox.SelectedIndex < _currentResumes.Count)
            {
                var orderedResumes = _currentResumes.OrderByDescending(r => r.ImportTime).ToList();
                var selectedResume = orderedResumes[listBox.SelectedIndex];
                ShowResumeDetails(selectedResume);
            }
        }

        private void DeleteResumeButton_Click(object sender, EventArgs e)
        {
            var uploadedListBox = panelAdd.Controls.Find("uploadedListBox", false).FirstOrDefault() as ListBox;

            if (uploadedListBox == null || uploadedListBox.SelectedIndex < 0)
            {
                MessageBox.Show("请先选择要删除的简历。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var orderedResumes = _currentResumes.OrderByDescending(r => r.ImportTime).ToList();
            var selectedResume = orderedResumes[uploadedListBox.SelectedIndex];

            var result = MessageBox.Show(
                $"确定要删除简历 '{selectedResume.Name} - {selectedResume.FileName}' 吗？\n此操作不可撤销！",
                "确认删除",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                try
                {
                    bool success = _managementService.DeleteResume(selectedResume.Id);

                    if (success)
                    {
                        MessageBox.Show("简历删除成功！", "删除成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadResumes(); // 刷新列表
                    }
                    else
                    {
                        MessageBox.Show("删除失败，未找到该简历。", "删除失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"删除失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void ViewDetailButton_Click(object sender, EventArgs e)
        {
            var uploadedListBox = panelAdd.Controls.Find("uploadedListBox", false).FirstOrDefault() as ListBox;

            if (uploadedListBox == null || uploadedListBox.SelectedIndex < 0)
            {
                MessageBox.Show("请先选择要查看的简历。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var orderedResumes = _currentResumes.OrderByDescending(r => r.ImportTime).ToList();
            var selectedResume = orderedResumes[uploadedListBox.SelectedIndex];
            ShowResumeDetails(selectedResume);
        }


        private void panelAdd_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panelSearch_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panelDuplicate_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panelReview_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panelStatistics_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panelExport_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panelMatch_Paint(object sender, PaintEventArgs e)
        {

        }

        // 简历检索
        private void InitializeSearchPanel()
        {
            // 标题
            var titleLabel = new Label
            {
                Text = "简历检索",
                Location = new Point(10, 10),
                Size = new Size(150, 30),
                Font = new Font("微软雅黑", 12, FontStyle.Bold)
            };
            panelSearch.Controls.Add(titleLabel);

            // 搜索字段选择（移除文件名选项）
            var searchFieldLabel = new Label
            {
                Text = "搜索:",
                Location = new Point(10, 50),
                Size = new Size(60, 20)
            };
            panelSearch.Controls.Add(searchFieldLabel);

            var searchFieldCombo = new ComboBox
            {
                Name = "searchFieldCombo",
                Location = new Point(100, 48),
                Size = new Size(120, 40),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            // 移除了"文件名"选项
            searchFieldCombo.Items.AddRange(new object[] { "全部", "姓名", "手机", "邮箱" });
            searchFieldCombo.SelectedIndex = 0;
            panelSearch.Controls.Add(searchFieldCombo);

            // 关键词输入
            var keywordLabel = new Label
            {
                Text = "关键词:",
                Location = new Point(230, 50),
                Size = new Size(70, 30)
            };
            panelSearch.Controls.Add(keywordLabel);

            var keywordTextBox = new TextBox
            {
                Name = "keywordTextBox",
                Location = new Point(305, 48),
                Size = new Size(150, 40)
            };
            panelSearch.Controls.Add(keywordTextBox);

            // 新增文件名搜索框
            var fileNameLabel = new Label
            {
                Text = "文件名:",
                Location = new Point(10, 120),  // 调整位置避免与日期控件重叠
                Size = new Size(70, 30)
            };
            panelSearch.Controls.Add(fileNameLabel);

            var fileNameTextBox = new TextBox
            {
                Name = "fileNameTextBox",  // 新增的文件名搜索框
                Location = new Point(80, 118),
                Size = new Size(200, 40)
            };
            panelSearch.Controls.Add(fileNameTextBox);

            // 开始日期（调整位置）
            var startDateLabel = new Label
            {
                Text = "开始:",
                Location = new Point(10, 160),  // 下移位置
                Size = new Size(60, 30)
            };
            panelSearch.Controls.Add(startDateLabel);

            var startDatePicker = new DateTimePicker
            {
                Name = "startDatePicker",
                Location = new Point(70, 158),  // 下移位置
                Size = new Size(150, 25),
                Format = DateTimePickerFormat.Short
            };
            startDatePicker.Value = DateTime.Now.AddMonths(-1);
            panelSearch.Controls.Add(startDatePicker);

            var startDateCheck = new CheckBox
            {
                Name = "startDateCheck",
                Text = "是",
                Location = new Point(225, 160),  // 下移位置
                Size = new Size(60, 30)
            };
            panelSearch.Controls.Add(startDateCheck);

            // 结束日期（调整位置）
            var endDateLabel = new Label
            {
                Text = "结束:",
                Location = new Point(290, 160),  // 下移位置
                Size = new Size(60, 30)
            };
            panelSearch.Controls.Add(endDateLabel);

            var endDatePicker = new DateTimePicker
            {
                Name = "endDatePicker",
                Location = new Point(355, 158),  // 下移位置
                Size = new Size(150, 25),
                Format = DateTimePickerFormat.Short
            };
            panelSearch.Controls.Add(endDatePicker);

            var endDateCheck = new CheckBox
            {
                Name = "endDateCheck",
                Text = "是",
                Location = new Point(520, 160),  // 下移位置
                Size = new Size(60, 30)
            };
            panelSearch.Controls.Add(endDateCheck);

            // 搜索按钮
            var searchButton = new Button
            {
                Text = "搜索",
                Location = new Point(465, 48),
                Size = new Size(100, 30),
                BackColor = Color.Blue,
                ForeColor = Color.White
            };
            searchButton.Click += SearchButton_Click;
            panelSearch.Controls.Add(searchButton);

            // 结果列表（调整位置）
            var resultLabel = new Label
            {
                Text = "搜索结果:",
                Location = new Point(10, 200),  // 下移位置
                Size = new Size(100, 30),
                Font = new Font("微软雅黑", 9, FontStyle.Bold)
            };
            panelSearch.Controls.Add(resultLabel);

            var searchResultListBox = new ListBox
            {
                Name = "searchResultListBox",
                Location = new Point(10, 235),  // 下移位置
                Size = new Size(560, 200),
                SelectionMode = SelectionMode.One
            };
            searchResultListBox.DoubleClick += SearchResultListBox_DoubleClick;
            panelSearch.Controls.Add(searchResultListBox);

            // 操作按钮（调整位置）
            var viewButton = new Button
            {
                Text = "查看详情",
                Location = new Point(10, 440),  // 下移位置
                Size = new Size(100, 35),
                BackColor = Color.Green,
                ForeColor = Color.White
            };
            viewButton.Click += SearchViewDetailButton_Click;
            panelSearch.Controls.Add(viewButton);

            var openFileButton = new Button
            {
                Text = "打开原稿",
                Location = new Point(120, 440),  // 下移位置
                Size = new Size(100, 35),
                BackColor = Color.Orange,
                ForeColor = Color.White
            };
            openFileButton.Click += OpenOriginalFileButton_Click;
            panelSearch.Controls.Add(openFileButton);
        }

        private void InitializeMatchPanel()
        {
            var title = new Label
            {
                Text = "语义匹配",
                Location = new Point(10, 10),
                Size = new Size(160, 30),
                Font = new Font("微软雅黑", 12, FontStyle.Bold)
            };
            panelMatch.Controls.Add(title);

            var jdLabel = new Label
            {
                Text = "职位描述:",
                Location = new Point(10, 50),
                Size = new Size(160, 30)
            };
            panelMatch.Controls.Add(jdLabel);

            var jdTextBox = new TextBox
            {
                Name = "jdTextBox",
                Location = new Point(10, 82),
                Size = new Size(560, 100),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical
            };
            panelMatch.Controls.Add(jdTextBox);

            var limitLabel = new Label
            {
                Text = "评估数量:",
                Location = new Point(10, 190),
                Size = new Size(80, 25)
            };
            panelMatch.Controls.Add(limitLabel);

            var limitNumeric = new NumericUpDown
            {
                Name = "matchLimitNumeric",
                Location = new Point(90, 188),
                Size = new Size(80, 25),
                Minimum = 1,
                Maximum = 200,
                Value = 20
            };
            panelMatch.Controls.Add(limitNumeric);

            var evaluateButton = new Button
            {
                Text = "匹配",
                Location = new Point(180, 186),
                Size = new Size(90, 28),
                BackColor = Color.MediumPurple,
                ForeColor = Color.White
            };
            evaluateButton.Click += EvaluateMatchButton_Click;
            panelMatch.Controls.Add(evaluateButton);

            var matchListView = new ListView
            {
                Name = "matchListView",
                Location = new Point(10, 225),
                Size = new Size(560, 200),
                View = View.Details,
                FullRowSelect = true,
                HideSelection = false,
                HeaderStyle = ColumnHeaderStyle.Nonclickable
            };
            matchListView.Columns.Add("姓名", 120, HorizontalAlignment.Left);
            matchListView.Columns.Add("分数", 60, HorizontalAlignment.Left);
            matchListView.Columns.Add("文件名", 220, HorizontalAlignment.Left);
            matchListView.Columns.Add("导入时间", 140, HorizontalAlignment.Left);
            panelMatch.Controls.Add(matchListView);

            var viewHighlightButton = new Button
            {
                Text = "查看高亮",
                Location = new Point(10, 430),
                Size = new Size(100, 28),
                BackColor = Color.SteelBlue,
                ForeColor = Color.White
            };
            viewHighlightButton.Click += ViewHighlightButton_Click;
            panelMatch.Controls.Add(viewHighlightButton);
        }

        private void SearchButton_Click(object sender, EventArgs e)
        {
            var keywordTextBox = panelSearch.Controls.Find("keywordTextBox", false).FirstOrDefault() as TextBox;
            var searchFieldCombo = panelSearch.Controls.Find("searchFieldCombo", false).FirstOrDefault() as ComboBox;
            var startDatePicker = panelSearch.Controls.Find("startDatePicker", false).FirstOrDefault() as DateTimePicker;
            var endDatePicker = panelSearch.Controls.Find("endDatePicker", false).FirstOrDefault() as DateTimePicker;
            var startDateCheck = panelSearch.Controls.Find("startDateCheck", false).FirstOrDefault() as CheckBox;
            var endDateCheck = panelSearch.Controls.Find("endDateCheck", false).FirstOrDefault() as CheckBox;
            var searchResultListBox = panelSearch.Controls.Find("searchResultListBox", false).FirstOrDefault() as ListBox;
            // 新增：获取文件名搜索框控件
            var fileNameTextBox = panelSearch.Controls.Find("fileNameTextBox", false).FirstOrDefault() as TextBox;

            string keyword = keywordTextBox?.Text ?? "";
            string searchField = searchFieldCombo?.SelectedItem?.ToString() ?? "全部";
            DateTime? startDate = startDateCheck?.Checked == true ? startDatePicker?.Value : null;
            DateTime? endDate = endDateCheck?.Checked == true ? endDatePicker?.Value : null;
            // 新增：获取文件名搜索值（去空格处理）
            string fileNameKeyword = fileNameTextBox?.Text?.Trim() ?? "";

            // 调整：调用搜索服务时传入文件名参数，参与筛选
            var results = _managementService.SearchResumes(
                keyword,
                startDate,
                endDate,
                searchField,
                fileNameKeyword // 新增：文件名搜索条件
            );
            _currentResumes = results;

            searchResultListBox?.Items.Clear();
            foreach (var resume in results.OrderByDescending(r => r.ImportTime))
            {
                searchResultListBox?.Items.Add($"{resume.Name} - {resume.FileName} - {resume.ImportTime:yyyy-MM-dd HH:mm}");
            }

            MessageBox.Show($"找到 {results.Count} 份简历", "搜索完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void SearchResultListBox_DoubleClick(object sender, EventArgs e)
        {
            var listBox = sender as ListBox;
            if (listBox?.SelectedIndex >= 0 && listBox.SelectedIndex < _currentResumes.Count)
            {
                var orderedResumes = _currentResumes.OrderByDescending(r => r.ImportTime).ToList();
                var selectedResume = orderedResumes[listBox.SelectedIndex];
                ShowResumeDetails(selectedResume);
            }
        }

        private class SemanticMatchResult
        {
            public string ResumeId { get; set; } = string.Empty;
            public int Score { get; set; }
            public List<string> Keywords { get; set; } = new List<string>();
        }

        private Dictionary<string, SemanticMatchResult> _semanticMatchResults = new Dictionary<string, SemanticMatchResult>();

        private async void EvaluateMatchButton_Click(object sender, EventArgs e)
        {
            try
            {
                var jdTextBox = panelMatch.Controls.Find("jdTextBox", false).FirstOrDefault() as TextBox;
                var limitNumeric = panelMatch.Controls.Find("matchLimitNumeric", false).FirstOrDefault() as NumericUpDown;
                var listView = panelMatch.Controls.Find("matchListView", false).FirstOrDefault() as ListView;

                if (_aiReviewService == null)
                {
                    MessageBox.Show("未配置AI密钥", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var jd = (jdTextBox.Text ?? string.Empty).Trim();
                if (jd.Length < 50 || jd.Length > 800)
                {
                    MessageBox.Show("请输入100-400字左右的职位描述。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var resumes = _managementService.LoadResumes().OrderByDescending(r => r.ImportTime).ToList();

                int limit = (int)limitNumeric.Value;
                resumes = resumes.Take(limit).ToList();
                if (resumes.Count == 0)
                {
                    MessageBox.Show("没有可评估的简历。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                listView.BeginUpdate();
                listView.Items.Clear();
                listView.EndUpdate();

                _semanticMatchResults.Clear();

                foreach (var resume in resumes)
                {
                    var summary = BuildResumeSummaryForAi(resume);
                    (int score, string[] keywords) result;
                    try
                    {
                        result = await _aiReviewService.ScoreSemanticMatchAsync(jd, summary);
                    }
                    catch (Exception ex)
                    {
                        result = (0, Array.Empty<string>());
                        Console.WriteLine("Match error: " + ex.Message);
                    }

                    var match = new SemanticMatchResult
                    {
                        ResumeId = resume.Id,
                        Score = result.score,
                        Keywords = result.keywords?.Where(k => !string.IsNullOrWhiteSpace(k)).Select(k => k.Trim()).Distinct().Take(10).ToList() ?? new List<string>()
                    };
                    _semanticMatchResults[resume.Id] = match;
                }

                var ordered = resumes
                    .OrderByDescending(r => _semanticMatchResults.TryGetValue(r.Id, out var m) ? m.Score : 0)
                    .ToList();

                listView.BeginUpdate();
                foreach (var r in ordered)
                {
                    var score = _semanticMatchResults.TryGetValue(r.Id, out var m) ? m.Score : 0;
                    var item = new ListViewItem(new[]
                    {
                        r.Name,
                        score.ToString(),
                        r.FileName,
                        r.ImportTime.ToString("yyyy-MM-dd HH:mm")
                    })
                    {
                        Tag = r.Id
                    };
                    listView.Items.Add(item);
                }
                listView.EndUpdate();

                MessageBox.Show($"评估完成，共计 {ordered.Count} 份。", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"评估失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string BuildResumeSummaryForAi(Resume resume)
        {
            var sb = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(resume.Name)) sb.AppendLine($"姓名：{resume.Name}");
            if (!string.IsNullOrWhiteSpace(resume.HighestEducationSchool)) sb.AppendLine($"最高学历学校：{resume.HighestEducationSchool}");
            if (!string.IsNullOrWhiteSpace(resume.HighestEducationMajor)) sb.AppendLine($"最高学历专业：{resume.HighestEducationMajor}");
            if (resume.Skills != null && resume.Skills.Count > 0) sb.AppendLine("技能：" + string.Join("、", resume.Skills));
            if (resume.WorkExperiences != null && resume.WorkExperiences.Count > 0)
            {
                sb.AppendLine("经历摘要：");
                foreach (var w in resume.WorkExperiences.Take(3))
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
            return sb.ToString();
        }

        private void ViewHighlightButton_Click(object sender, EventArgs e)
        {
            var listView = panelMatch.Controls.Find("matchListView", false).FirstOrDefault() as ListView;
            if (listView == null || listView.SelectedItems.Count == 0)
            {
                MessageBox.Show("请先在匹配结果中选择一条。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var id = listView.SelectedItems[0].Tag as string;
            var resume = (_currentResumes ?? new List<Resume>()).FirstOrDefault(r => r.Id == id)
                         ?? _managementService.GetResumeById(id);
            if (resume == null)
            {
                MessageBox.Show("未找到该简历。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var terms = _semanticMatchResults.TryGetValue(resume.Id, out var m) ? m.Keywords : new List<string>();
            ShowResumeDetailsWithHighlight(resume, terms);
        }

        private void LoadMatchJobs()
        {
            try
            {
                _currentResumes = _managementService.LoadResumes()
                    .OrderByDescending(r => r.ImportTime)
                    .ToList();
            }
            catch
            {
                _currentResumes = new List<Resume>();
            }
            var listView = panelMatch.Controls.Find("matchListView", false).FirstOrDefault() as ListView;
            if (listView != null)
            {
                listView.Items.Clear();
            }
            _semanticMatchResults.Clear();
        }

        private void ShowResumeDetailsWithHighlight(Resume resume, List<string> highlightTerms)
        {
            var detailsForm = new Form
            {
                Text = $"简历详情(高亮) - {resume.Name}",
                Size = new Size(700, 540),
                StartPosition = FormStartPosition.CenterParent
            };

            var rich = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                DetectUrls = false,
                HideSelection = false
            };

            var text = GetResumeDetailsText(resume);
            rich.Text = text;

            // 简单高亮：按关键短语标注；颜色循环
            var colors = new[] { Color.Yellow, Color.LightGreen, Color.LightPink, Color.LightBlue, Color.Khaki, Color.Thistle };
            int ci = 0;
            foreach (var term in highlightTerms.Where(t => !string.IsNullOrWhiteSpace(t)).Distinct())
            {
                var termTrim = term.Trim();
                if (termTrim.Length < 2) continue;

                int start = 0;
                var color = colors[ci % colors.Length];
                ci++;
                while (start < rich.TextLength)
                {
                    int idx = rich.Find(termTrim, start, RichTextBoxFinds.None);
                    if (idx < 0) break;
                    rich.Select(idx, termTrim.Length);
                    rich.SelectionBackColor = color;
                    start = idx + termTrim.Length;
                }
            }

            detailsForm.Controls.Add(rich);
            detailsForm.ShowDialog();
        }

        private void SearchViewDetailButton_Click(object sender, EventArgs e)
        {
            var searchResultListBox = panelSearch.Controls.Find("searchResultListBox", false).FirstOrDefault() as ListBox;

            if (searchResultListBox == null || searchResultListBox.SelectedIndex < 0)
            {
                MessageBox.Show("请先选择要查看的简历。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var orderedResumes = _currentResumes.OrderByDescending(r => r.ImportTime).ToList();
            var selectedResume = orderedResumes[searchResultListBox.SelectedIndex];
            ShowResumeDetails(selectedResume);
        }

        private void OpenOriginalFileButton_Click(object sender, EventArgs e)
        {
            var searchResultListBox = panelSearch.Controls.Find("searchResultListBox", false).FirstOrDefault() as ListBox;

            if (searchResultListBox == null || searchResultListBox.SelectedIndex < 0)
            {
                MessageBox.Show("请先选择要打开的简历。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var orderedResumes = _currentResumes.OrderByDescending(r => r.ImportTime).ToList();
            var selectedResume = orderedResumes[searchResultListBox.SelectedIndex];

            if (string.IsNullOrEmpty(selectedResume.OriginalFilePath) || !File.Exists(selectedResume.OriginalFilePath))
            {
                MessageBox.Show("原始文件不存在或路径无效。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = selectedResume.OriginalFilePath,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开文件失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 简历查重
        private void InitializeDuplicatePanel()
        {
            // 标题
            var titleLabel = new Label
            {
                Text = "简历查重",
                Location = new Point(10, 10),
                Size = new Size(150, 30),
                Font = new Font("微软雅黑", 12, FontStyle.Bold)
            };
            panelDuplicate.Controls.Add(titleLabel);

            // 查重属性选择
            var checkOptionsLabel = new Label
            {
                Text = "查重属性:",
                Location = new Point(10, 50),
                Size = new Size(100, 20),
                Font = new Font("微软雅黑", 9, FontStyle.Bold)
            };
            panelDuplicate.Controls.Add(checkOptionsLabel);

            var checkNameCheckBox = new CheckBox
            {
                Name = "checkNameCheckBox",
                Text = "姓名",
                Location = new Point(10, 75),
                Size = new Size(80, 30),
                Checked = true
            };
            panelDuplicate.Controls.Add(checkNameCheckBox);

            var checkPhoneCheckBox = new CheckBox
            {
                Name = "checkPhoneCheckBox",
                Text = "手机号码",
                Location = new Point(100, 75),
                Size = new Size(100, 30),
                Checked = true
            };
            panelDuplicate.Controls.Add(checkPhoneCheckBox);

            var checkEmailCheckBox = new CheckBox
            {
                Name = "checkEmailCheckBox",
                Text = "邮箱",
                Location = new Point(210, 75),
                Size = new Size(80, 30),
                Checked = true
            };
            panelDuplicate.Controls.Add(checkEmailCheckBox);

            var checkIdCardCheckBox = new CheckBox
            {
                Name = "checkIdCardCheckBox",
                Text = "身份证号",
                Location = new Point(300, 75),
                Size = new Size(100, 30),
                Checked = true
            };
            panelDuplicate.Controls.Add(checkIdCardCheckBox);

            // 查重按钮
            var checkDuplicateButton = new Button
            {
                Text = "开始查重",
                Location = new Point(10, 105),
                Size = new Size(120, 35),
                BackColor = Color.Blue,
                ForeColor = Color.White
            };
            checkDuplicateButton.Click += CheckDuplicateButton_Click;
            panelDuplicate.Controls.Add(checkDuplicateButton);

            // 导出按钮
            var exportTxtButton = new Button
            {
                Text = "导出为TXT",
                Location = new Point(140, 105),
                Size = new Size(120, 35),
                BackColor = Color.Green,
                ForeColor = Color.White
            };
            exportTxtButton.Click += ExportTxtButton_Click;
            panelDuplicate.Controls.Add(exportTxtButton);

            var exportWordButton = new Button
            {
                Text = "导出为Word",
                Location = new Point(270, 105),
                Size = new Size(120, 35),
                BackColor = Color.Green,
                ForeColor = Color.White
            };
            exportWordButton.Click += ExportWordButton_Click;
            panelDuplicate.Controls.Add(exportWordButton);

            // 结果显示
            var resultLabel = new Label
            {
                Text = "查重结果:",
                Location = new Point(10, 150),
                Size = new Size(100, 30),
                Font = new Font("微软雅黑", 9, FontStyle.Bold)
            };
            panelDuplicate.Controls.Add(resultLabel);

            var duplicateResultTextBox = new TextBox
            {
                Name = "duplicateResultTextBox",
                Location = new Point(10, 205),
                Size = new Size(560, 280),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true
            };
            panelDuplicate.Controls.Add(duplicateResultTextBox);
        }

        private void InitializeReviewPanel()
        {
            var titleLabel = new Label
            {
                Text = "简历点评",
                Location = new Point(10, 10),
                Size = new Size(150, 30),
                Font = new Font("微软雅黑", 12, FontStyle.Bold)
            };
            panelReview.Controls.Add(titleLabel);

            var refreshButton = new Button
            {
                Text = "刷新",
                Location = new Point(480, 10),
                Size = new Size(90, 32),
                Name = "refreshReviewButton"
            };
            refreshButton.Click += RefreshReviewButton_Click;
            panelReview.Controls.Add(refreshButton);

            var resumeListBox = new ListBox
            {
                Name = "reviewResumeListBox",
                Location = new Point(10, 50),
                Size = new Size(200, 400)
            };
            resumeListBox.SelectedIndexChanged += ReviewResumeListBox_SelectedIndexChanged;
            panelReview.Controls.Add(resumeListBox);

            var detailsLabel = new Label
            {
                Text = "简历详情:",
                Location = new Point(220, 50),
                Size = new Size(120, 25)
            };
            panelReview.Controls.Add(detailsLabel);

            var detailsTextBox = new TextBox
            {
                Name = "reviewDetailsTextBox",
                Location = new Point(220, 78),
                Size = new Size(350, 180),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                ReadOnly = true
            };
            panelReview.Controls.Add(detailsTextBox);

            var manualLabel = new Label
            {
                Text = "点评:",
                Location = new Point(220, 270),
                Size = new Size(120, 25)
            };
            panelReview.Controls.Add(manualLabel);

            var manualTextBox = new TextBox
            {
                Name = "manualReviewTextBox",
                Location = new Point(220, 298),
                Size = new Size(350, 130),
                Multiline = true,
                ScrollBars = ScrollBars.Vertical
            };
            panelReview.Controls.Add(manualTextBox);

            var statusLabel = new Label
            {
                Name = "reviewStatusLabel",
                Location = new Point(220, 430),
                Size = new Size(350, 30),
                ForeColor = Color.Gray
            };
            panelReview.Controls.Add(statusLabel);

            var generateButton = new Button
            {
                Text = "调用AI生成",
                Location = new Point(220, 460),
                Size = new Size(120, 35),
                Name = "generateReviewButton"
            };
            generateButton.Click += GenerateReviewButton_Click;
            panelReview.Controls.Add(generateButton);

            var saveButton = new Button
            {
                Text = "保存点评",
                Location = new Point(350, 460),
                Size = new Size(120, 35),
                Name = "saveReviewButton"
            };
            saveButton.Click += SaveReviewButton_Click;
            panelReview.Controls.Add(saveButton);
        }

        private void InitializeStatisticsPanel()
        {
            var titleLabel = new Label
            {
                Text = "数据分析与统计",
                Location = new Point(10, 10),
                Size = new Size(200, 30),
                Font = new Font("微软雅黑", 12, FontStyle.Bold)
            };
            panelStatistics.Controls.Add(titleLabel);

            var refreshButton = new Button
            {
                Text = "刷新数据",
                Location = new Point(380, 10),
                Size = new Size(90, 32),
                Name = "refreshStatisticsButton"
            };
            refreshButton.Click += RefreshStatisticsButton_Click;
            panelStatistics.Controls.Add(refreshButton);

            var generateReportButton = new Button
            {
                Text = "生成报告",
                Location = new Point(480, 10),
                Size = new Size(90, 32),
                Name = "generateReportButton",
                BackColor = Color.Green,
                ForeColor = Color.White
            };
            generateReportButton.Click += GenerateReportButton_Click;
            panelStatistics.Controls.Add(generateReportButton);

            var summaryLabel = new Label
            {
                Name = "statisticsSummaryLabel",
                Location = new Point(10, 50),
                Size = new Size(560, 60),
                Font = new Font("微软雅黑", 9)
            };
            panelStatistics.Controls.Add(summaryLabel);

            // 使用TabControl来分别显示数据列表和图表
            var tabControl = new TabControl
            {
                Name = "statisticsTabControl",
                Location = new Point(10, 120),
                Size = new Size(560, 380)
            };

            // 数据列表标签页
            var listTabPage = new TabPage("数据列表");
            var listView = new ListView
            {
                Name = "statisticsListView",
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                HeaderStyle = ColumnHeaderStyle.Nonclickable,
                HideSelection = false,
                ShowGroups = true
            };
            listView.Columns.Add("分类", 420, HorizontalAlignment.Left);
            listView.Columns.Add("数量", 100, HorizontalAlignment.Left);
            listTabPage.Controls.Add(listView);
            tabControl.TabPages.Add(listTabPage);

            // 图表标签页
            var chartTabPage = new TabPage("图表展示");
            var pictureBox = new PictureBox
            {
                Name = "statisticsPictureBox",
                Dock = DockStyle.Fill,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = Color.White
            };
            chartTabPage.Controls.Add(pictureBox);
            tabControl.TabPages.Add(chartTabPage);

            panelStatistics.Controls.Add(tabControl);
        }

        private void InitializeExportPanel()
        {
            var titleLabel = new Label
            {
                Text = "简历导出",
                Location = new Point(10, 10),
                Size = new Size(150, 30),
                Font = new Font("微软雅黑", 12, FontStyle.Bold)
            };
            panelExport.Controls.Add(titleLabel);

            var refreshButton = new Button
            {
                Text = "刷新",
                Location = new Point(480, 10),
                Size = new Size(90, 32),
                Name = "refreshExportButton"
            };
            refreshButton.Click += RefreshExportButton_Click;
            panelExport.Controls.Add(refreshButton);

            var exportTabControl = new TabControl
            {
                Name = "exportTabControl",
                Location = new Point(10, 50),
                Size = new Size(560, 450)
            };

            var directExportTab = new TabPage("直接导出");
            var formatLabel = new Label
            {
                Text = "导出格式:",
                Location = new Point(10, 10),
                Size = new Size(90, 25)
            };
            directExportTab.Controls.Add(formatLabel);

            var formatCombo = new ComboBox
            {
                Name = "exportFormatCombo",
                Location = new Point(100, 8),
                Size = new Size(120, 30),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            formatCombo.Items.AddRange(new object[] { "JSON", "CSV", "Word"});
            formatCombo.SelectedIndex = 0;
            directExportTab.Controls.Add(formatCombo);

            var exportList = new CheckedListBox
            {
                Name = "exportCheckedListBox",
                Location = new Point(10, 45),
                Size = new Size(300, 350),
                CheckOnClick = true
            };
            directExportTab.Controls.Add(exportList);

            var exportButton = new Button
            {
                Text = "导出所选",
                Location = new Point(330, 45),
                Size = new Size(120, 35),
                Name = "exportSelectedButton"
            };
            exportButton.Click += ExportSelectedButton_Click;
            directExportTab.Controls.Add(exportButton);

            var summaryLabel = new Label
            {
                Name = "exportSummaryLabel",
                Location = new Point(330, 95),
                Size = new Size(250, 80),
                Font = new Font("微软雅黑", 9)
            };
            directExportTab.Controls.Add(summaryLabel);

            exportTabControl.TabPages.Add(directExportTab);

            var convertTab = new TabPage("格式转换");

            var convertTitleLabel = new Label
            {
                Text = "格式转换工具",
                Location = new Point(10, 10),
                Size = new Size(200, 25),
                Font = new Font("微软雅黑", 11, FontStyle.Bold)
            };
            convertTab.Controls.Add(convertTitleLabel);

            var sourceFormatLabel = new Label
            {
                Text = "源格式:",
                Location = new Point(10, 45),
                Size = new Size(80, 25)
            };
            convertTab.Controls.Add(sourceFormatLabel);

            var sourceFormatCombo = new ComboBox
            {
                Name = "sourceFormatCombo",
                Location = new Point(100, 43),
                Size = new Size(120, 30),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            sourceFormatCombo.Items.AddRange(new object[] { "JSON", "CSV" });
            sourceFormatCombo.SelectedIndex = 0;
            convertTab.Controls.Add(sourceFormatCombo);

            var targetFormatLabel = new Label
            {
                Text = "目标格式:",
                Location = new Point(240, 45),
                Size = new Size(80, 25)
            };
            convertTab.Controls.Add(targetFormatLabel);

            var targetFormatCombo = new ComboBox
            {
                Name = "targetFormatCombo",
                Location = new Point(330, 43),
                Size = new Size(120, 30),
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            targetFormatCombo.Items.AddRange(new object[] { "JSON", "CSV", "Word" });
            targetFormatCombo.SelectedIndex = 1;
            convertTab.Controls.Add(targetFormatCombo);

            var convertButton = new Button
            {
                Text = "选择文件转换",
                Location = new Point(460, 42),
                Size = new Size(90, 32),
                Name = "convertFormatButton",
                BackColor = Color.Orange,
                ForeColor = Color.White
            };
            convertButton.Click += ConvertFormatButton_Click;
            convertTab.Controls.Add(convertButton);
            var convertStatusLabel = new Label
            {
                Name = "convertStatusLabel",
                Location = new Point(10, 175),
                Size = new Size(540, 25),
                Font = new Font("微软雅黑", 9)
            };
            convertTab.Controls.Add(convertStatusLabel);

            exportTabControl.TabPages.Add(convertTab);
            panelExport.Controls.Add(exportTabControl);
        }

        private Dictionary<string, List<Resume>> _currentDuplicates = new Dictionary<string, List<Resume>>();

        private void RefreshReviewButton_Click(object sender, EventArgs e)
        {
            LoadReviewResumes();
        }

        private async void GenerateReviewButton_Click(object sender, EventArgs e)
        {
            var button = sender as Button;
            if (button != null) button.Enabled = false;

            try
            {
                if (_aiReviewService == null)
                {
                    MessageBox.Show("尚未配置AI密钥。请在系统环境变量中设置QWEN_API_KEY，然后重启程序。可选：AI_BASE_URL、AI_MODEL。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var resume = GetSelectedReviewResume();
                if (resume == null)
                {
                    UpdateReviewStatus("请先在左侧列表选择一份简历。", true);
                    return;
                }

                UpdateReviewStatus("AI正在生成点评...", false);

                var review = await _aiReviewService.GenerateReviewAsync(resume);

                var manualReviewTextBox = panelReview.Controls.Find("manualReviewTextBox", false).FirstOrDefault() as TextBox;
                if (manualReviewTextBox != null)
                {
                    manualReviewTextBox.Text = review?.Trim() ?? string.Empty;
                }

                UpdateReviewStatus("AI点评生成完成，已填入文本框。", false);
            }
            catch (Exception ex)
            {
                UpdateReviewStatus($"生成失败: {ex.Message}", true);
                MessageBox.Show($"生成失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (button != null) button.Enabled = true;
            }
        }

        private Resume GetSelectedReviewResume()
        {
            var listBox = panelReview.Controls.Find("reviewResumeListBox", false).FirstOrDefault() as ListBox;
            if (listBox == null || listBox.SelectedIndex < 0 || listBox.SelectedIndex >= _reviewResumes.Count)
            {
                return null;
            }

            return _reviewResumes[listBox.SelectedIndex];
        }


        private void SaveReviewButton_Click(object sender, EventArgs e)
        {
            var resumeListBox = panelReview.Controls.Find("reviewResumeListBox", false).FirstOrDefault() as ListBox;
            var manualReviewTextBox = panelReview.Controls.Find("manualReviewTextBox", false).FirstOrDefault() as TextBox;

            if (resumeListBox == null || manualReviewTextBox == null)
            {
                UpdateReviewStatus("界面初始化失败，无法保存点评。", true);
                return;
            }

            if (resumeListBox.SelectedIndex < 0 || resumeListBox.SelectedIndex >= _reviewResumes.Count)
            {
                UpdateReviewStatus("请先选择一份简历。", true);
                return;
            }

            var resume = _reviewResumes[resumeListBox.SelectedIndex];
            _reviewNotes[resume.Id] = manualReviewTextBox.Text ?? string.Empty;

            try
            {
                SaveReviewNotes();
                UpdateReviewStatus("点评已保存。", false);
            }
            catch (Exception ex)
            {
                UpdateReviewStatus($"保存失败: {ex.Message}", true);
            }
        }

        private void ReviewResumeListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            var listBox = sender as ListBox;
            var detailsTextBox = panelReview.Controls.Find("reviewDetailsTextBox", false).FirstOrDefault() as TextBox;
            var manualReviewTextBox = panelReview.Controls.Find("manualReviewTextBox", false).FirstOrDefault() as TextBox;

            if (listBox == null || detailsTextBox == null || manualReviewTextBox == null)
            {
                return;
            }

            if (listBox.SelectedIndex < 0 || listBox.SelectedIndex >= _reviewResumes.Count)
            {
                detailsTextBox.Text = string.Empty;
                manualReviewTextBox.Text = string.Empty;
                return;
            }

            var resume = _reviewResumes[listBox.SelectedIndex];
            detailsTextBox.Text = GetResumeDetailsText(resume);

            if (_reviewNotes.TryGetValue(resume.Id, out var note))
            {
                manualReviewTextBox.Text = note;
            }
            else
            {
                manualReviewTextBox.Text = string.Empty;
            }

            UpdateReviewStatus("已加载选择的简历。", false);
        }

        private void LoadReviewResumes()
        {
            try
            {
                _reviewResumes = _managementService.LoadResumes()
                    .OrderByDescending(r => r.ImportTime)
                    .ToList();
            }
            catch (Exception ex)
            {
                UpdateReviewStatus($"加载简历失败: {ex.Message}", true);
                return;
            }

            var resumeListBox = panelReview.Controls.Find("reviewResumeListBox", false).FirstOrDefault() as ListBox;
            var detailsTextBox = panelReview.Controls.Find("reviewDetailsTextBox", false).FirstOrDefault() as TextBox;
            var manualReviewTextBox = panelReview.Controls.Find("manualReviewTextBox", false).FirstOrDefault() as TextBox;

            if (resumeListBox == null || detailsTextBox == null || manualReviewTextBox == null)
            {
                UpdateReviewStatus("界面初始化失败。", true);
                return;
            }

            resumeListBox.Items.Clear();

            foreach (var resume in _reviewResumes)
            {
                resumeListBox.Items.Add($"{resume.Name} - {resume.FileName}");
            }

            if (_reviewResumes.Count > 0)
            {
                resumeListBox.SelectedIndex = 0;
                UpdateReviewStatus($"共 {_reviewResumes.Count} 份简历。", false);
            }
            else
            {
                detailsTextBox.Text = string.Empty;
                manualReviewTextBox.Text = string.Empty;
                UpdateReviewStatus("暂无简历数据。", false);
            }
        }

        private void RefreshStatisticsButton_Click(object sender, EventArgs e)
        {
            LoadStatistics();
        }

        private void LoadStatistics()
        {
            var summaryLabel = panelStatistics.Controls.Find("statisticsSummaryLabel", false).FirstOrDefault() as Label;

            var tabControl = panelStatistics.Controls.Find("statisticsTabControl", false).FirstOrDefault() as TabControl;
            if (tabControl == null)
            {
                return;
            }
            ListView listView = null;
            if (tabControl.TabPages.Count > 0)
            {
                listView = tabControl.TabPages[0].Controls.Find("statisticsListView", false).FirstOrDefault() as ListView;
            }
            PictureBox pictureBox = null;
            if (tabControl.TabPages.Count > 1)
            {
                pictureBox = tabControl.TabPages[1].Controls.Find("statisticsPictureBox", false).FirstOrDefault() as PictureBox;
            }

            List<Resume> allResumes;
            try
            {
                allResumes = _managementService.LoadResumes();
            }
            catch (Exception ex)
            {
                summaryLabel.Text = $"加载数据失败: {ex.Message}";
                listView.Items.Clear();
                listView.Groups.Clear();
                return;
            }

            if (allResumes.Count == 0)
            {
                summaryLabel.Text = "暂无简历数据。";
                listView.Items.Clear();
                listView.Groups.Clear();
                pictureBox.Image = null;
                return;
            }

            var earliest = allResumes.Min(r => r.ImportTime);
            var latest = allResumes.Max(r => r.ImportTime);

            //年龄分布
            var ageGroups = CalculateAgeDistribution(allResumes);
            var totalWithAge = ageGroups.Values.Sum();

            //学历
            var educationGroups = CalculateEducationDistribution(allResumes);

            //计算技能
            var skillGroups = CalculateSkillDistribution(allResumes);
            var topSkills = skillGroups.OrderByDescending(kvp => kvp.Value).Take(10).ToList();

            summaryLabel.Text = $"总简历数: {allResumes.Count} 份 | " +
                               $"最早导入: {earliest:yyyy-MM-dd} | " +
                               $"最近导入: {latest:yyyy-MM-dd} | " +
                               $"有效年龄数据: {totalWithAge} 份 | " +
                               $"热门技能: {(topSkills.Count > 0 ? topSkills[0].Key : "无")}";

            // 填充列表视图
            listView.BeginUpdate();
            listView.Items.Clear();
            listView.Groups.Clear();

            //性别
            var genderGroup = new ListViewGroup("按性别分布");
            listView.Groups.Add(genderGroup);
            foreach (var group in allResumes
                         .GroupBy(r => string.IsNullOrWhiteSpace(r.Gender) ? "未填写" : r.Gender)
                         .OrderByDescending(g => g.Count()))
            {
                var item = new ListViewItem(new[] { group.Key, $"{group.Count()} 份 ({group.Count() * 100.0 / allResumes.Count:F1}%)" })
                {
                    Group = genderGroup
                };
                listView.Items.Add(item);
            }

            //年龄
            var ageGroup = new ListViewGroup("按年龄分布");
            listView.Groups.Add(ageGroup);
            foreach (var kvp in ageGroups.OrderByDescending(g => g.Value))
            {
                var percentage = totalWithAge > 0 ? kvp.Value * 100.0 / totalWithAge : 0;
                var item = new ListViewItem(new[] { kvp.Key, $"{kvp.Value} 人 ({percentage:F1}%)" })
                {
                    Group = ageGroup
                };
                listView.Items.Add(item);
            }

            //学历
            var educationGroup = new ListViewGroup("按学历分布");
            listView.Groups.Add(educationGroup);
            foreach (var kvp in educationGroups.OrderByDescending(g => g.Value))
            {
                var percentage = kvp.Value * 100.0 / allResumes.Count;
                var item = new ListViewItem(new[] { kvp.Key, $"{kvp.Value} 份 ({percentage:F1}%)" })
                {
                    Group = educationGroup
                };
                listView.Items.Add(item);
            }

            //技能（Top 10）
            var skillGroup = new ListViewGroup("热门技能 Top 10");
            listView.Groups.Add(skillGroup);
            foreach (var kvp in topSkills)
            {
                var item = new ListViewItem(new[] { kvp.Key, $"{kvp.Value} 份" })
                {
                    Group = skillGroup
                };
                listView.Items.Add(item);
            }

            //月份
            var monthGroup = new ListViewGroup("按导入月份分布");
            listView.Groups.Add(monthGroup);
            foreach (var group in allResumes
                         .GroupBy(r => r.ImportTime.ToString("yyyy-MM"))
                         .OrderByDescending(g => g.Key))
            {
                var percentage = group.Count() * 100.0 / allResumes.Count;
                var item = new ListViewItem(new[] { group.Key, $"{group.Count()} 份 ({percentage:F1}%)" })
                {
                    Group = monthGroup
                };
                listView.Items.Add(item);
            }

            listView.EndUpdate();

            // 生成图表
            GenerateStatisticsCharts(pictureBox, allResumes, ageGroups, educationGroups, skillGroups);
        }

        private Dictionary<string, int> CalculateAgeDistribution(List<Resume> resumes)
        {
            var ageGroups = new Dictionary<string, int>
            {
                { "18-25岁", 0 },
                { "26-30岁", 0 },
                { "31-35岁", 0 },
                { "36-40岁", 0 },
                { "41-50岁", 0 },
                { "50岁以上", 0 }
            };

            foreach (var resume in resumes)
            {
                if (resume.BirthDate.HasValue)
                {
                    var age = DateTime.Now.Year - resume.BirthDate.Value.Year;
                    if (resume.BirthDate.Value.Date > DateTime.Now.AddYears(-age)) age--;

                    if (age >= 18 && age <= 25) ageGroups["18-25岁"]++;
                    else if (age >= 26 && age <= 30) ageGroups["26-30岁"]++;
                    else if (age >= 31 && age <= 35) ageGroups["31-35岁"]++;
                    else if (age >= 36 && age <= 40) ageGroups["36-40岁"]++;
                    else if (age >= 41 && age <= 50) ageGroups["41-50岁"]++;
                    else if (age > 50) ageGroups["50岁以上"]++;
                }
            }

            return ageGroups;
        }

        private Dictionary<string, int> CalculateEducationDistribution(List<Resume> resumes)
        {
            var educationGroups = new Dictionary<string, int>();

            foreach (var resume in resumes)
            {
                var education = string.IsNullOrWhiteSpace(resume.HighestEducationSchool)
                    ? "未填写"
                    : ExtractEducationLevel(resume.HighestEducationSchool);

                if (!educationGroups.ContainsKey(education))
                {
                    educationGroups[education] = 0;
                }
                educationGroups[education]++;
            }

            return educationGroups;
        }

        private string ExtractEducationLevel(string school)
        {
            if (string.IsNullOrWhiteSpace(school)) return "未填写";

            school = school.ToLower();
            if (school.Contains("博士") || school.Contains("phd")) return "博士";
            if (school.Contains("硕士") || school.Contains("研究生") || school.Contains("master")) return "硕士";
            if (school.Contains("本科") || school.Contains("学士") || school.Contains("bachelor") || school.Contains("大学")) return "本科";
            if (school.Contains("专科") || school.Contains("大专")) return "专科";
            if (school.Contains("高中") || school.Contains("中专")) return "高中及以下";

            return "其他";
        }

        private Dictionary<string, int> CalculateSkillDistribution(List<Resume> resumes)
        {
            var skillGroups = new Dictionary<string, int>();

            foreach (var resume in resumes)
            {
                foreach (var skill in resume.Skills)
                {
                    if (!string.IsNullOrWhiteSpace(skill))
                    {
                        var skillKey = skill.Trim();
                        if (!skillGroups.ContainsKey(skillKey))
                        {
                            skillGroups[skillKey] = 0;
                        }
                        skillGroups[skillKey]++;
                    }
                }
            }

            return skillGroups;
        }

        private void GenerateStatisticsCharts(PictureBox pictureBox, List<Resume> allResumes,
            Dictionary<string, int> ageGroups, Dictionary<string, int> educationGroups,
            Dictionary<string, int> skillGroups)
        {
            if (pictureBox == null) return;

            const int width = 800;
            const int height = 800;
            var bitmap = new Bitmap(width, height);
            using (var g = Graphics.FromImage(bitmap))
            {
                g.Clear(Color.White);
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;

                var font = new Font("微软雅黑", 10);
                var titleFont = new Font("微软雅黑", 14, FontStyle.Bold);
                var smallFont = new Font("微软雅黑", 8);
                var brush = new SolidBrush(Color.Black);

                g.DrawString("简历数据统计图表", titleFont, brush, width / 2 - 100, 20);

                // 绘制性别分布饼图
                var genderData = allResumes.GroupBy(r => string.IsNullOrWhiteSpace(r.Gender) ? "未填写" : r.Gender)
                    .ToDictionary(g => g.Key, g => g.Count());
                DrawPieChart(g, "性别分布", genderData, new Rectangle(20, 80, 180, 180), font, smallFont);

                // 绘制学历分布柱状图
                var topEducation = educationGroups.OrderByDescending(kvp => kvp.Value).Take(6).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
                DrawBarChart(g, "学历分布", topEducation, new Rectangle(290, 80, 260, 180), font, smallFont);

                // 绘制年龄分布柱状图
                var validAgeGroups = ageGroups.Where(kvp => kvp.Value > 0).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
                if (validAgeGroups.Count > 0)
                {
                    DrawBarChart(g, "年龄分布", validAgeGroups, new Rectangle(20, 300, 260, 180), font, smallFont);
                }

                // 绘制技能分布（Top 5）
                var topSkills = skillGroups.OrderByDescending(kvp => kvp.Value).Take(5).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
                if (topSkills.Count > 0)
                {
                    DrawBarChart(g, "热门技能 Top 5", topSkills, new Rectangle(300, 300, 260, 180), font, smallFont);
                }
            }

            // 释放旧图片资源
            var oldImage = pictureBox.Image;
            if (oldImage != null && oldImage != bitmap)
            {
                oldImage.Dispose();
            }
            pictureBox.Image = bitmap;
        }

        private void DrawPieChart(Graphics g, string title, Dictionary<string, int> data, Rectangle bounds, Font font, Font smallFont)
        {
            if (data.Count == 0) return;

            var total = data.Values.Sum();
            if (total == 0) return;

            var colors = new[] { Color.Red, Color.Blue, Color.Green, Color.Orange, Color.Purple, Color.Yellow, Color.Cyan, Color.Magenta };
            var centerX = bounds.X + bounds.Width / 2;
            var centerY = bounds.Y + bounds.Height / 2;
            var radius = Math.Min(bounds.Width, bounds.Height) / 2 - 30;
            var rect = new Rectangle(centerX - radius, centerY - radius, radius * 2, radius * 2);

            float startAngle = -90;
            int colorIndex = 0;

            g.DrawString(title, font, Brushes.Black, bounds.X, bounds.Y - 20);

            foreach (var kvp in data.OrderByDescending(d => d.Value))
            {
                var sweepAngle = (float)(kvp.Value * 360.0 / total);
                var brush = new SolidBrush(colors[colorIndex % colors.Length]);

                g.FillPie(brush, rect, startAngle, sweepAngle);
                g.DrawPie(Pens.Black, rect, startAngle, sweepAngle);

                // 绘制图例
                var legendY = bounds.Y + bounds.Height - (data.Count - colorIndex) * 20;
                g.FillRectangle(brush, bounds.X + bounds.Width - 60, legendY - 10, 15, 15);
                g.DrawRectangle(Pens.Black, bounds.X + bounds.Width - 60, legendY - 10, 15, 15);

                var percentage = kvp.Value * 100.0 / total;
                var legendText = $"{kvp.Key}: {kvp.Value} ({percentage:F1}%)";
                g.DrawString(legendText, smallFont, Brushes.Black, bounds.X + bounds.Width - 43, legendY - 12);

                startAngle += sweepAngle;
                colorIndex++;
                brush.Dispose();
            }
        }

        private void DrawBarChart(Graphics g, string title, Dictionary<string, int> data, Rectangle bounds, Font font, Font smallFont)
        {
            if (data.Count == 0) return;

            var maxValue = data.Values.Max();
            if (maxValue == 0) return;

            g.DrawString(title, font, Brushes.Black, bounds.X, bounds.Y);

            var chartArea = new Rectangle(bounds.X, bounds.Y + 25, bounds.Width - 60, bounds.Height - 45);
            var barWidth = chartArea.Width / (data.Count * 2);
            var colors = new[] { Color.SteelBlue, Color.DarkGreen, Color.Orange, Color.Red, Color.Purple, Color.Teal };

            int index = 0;
            foreach (var kvp in data.OrderByDescending(d => d.Value))
            {
                var barHeight = (int)(kvp.Value * chartArea.Height / maxValue);
                var x = chartArea.X + index * (chartArea.Width / data.Count);
                var y = chartArea.Y + chartArea.Height - barHeight;
                var barRect = new Rectangle(x + 5, y, barWidth, barHeight);

                var brush = new SolidBrush(colors[index % colors.Length]);
                g.FillRectangle(brush, barRect);
                g.DrawRectangle(Pens.Black, barRect);
                brush.Dispose();

                // 绘制数值
                var valueText = kvp.Value.ToString();
                var textSize = g.MeasureString(valueText, smallFont);
                g.DrawString(valueText, smallFont, Brushes.Black,
                    x + 5 + (barWidth - textSize.Width) / 2, y - textSize.Height - 2);

                // 绘制标签（旋转）
                var labelText = kvp.Key.Length > 6 ? kvp.Key.Substring(0, 6) + "..." : kvp.Key;
                var labelSize = g.MeasureString(labelText, smallFont);
                g.TranslateTransform(x + 5 + barWidth / 2, chartArea.Y + chartArea.Height + 10);
                g.RotateTransform(-45);
                g.DrawString(labelText, smallFont, Brushes.Black, -labelSize.Width / 2, 0);
                g.ResetTransform();

                index++;
            }

            // 绘制Y轴
            g.DrawLine(Pens.Black, chartArea.X, chartArea.Y, chartArea.X, chartArea.Y + chartArea.Height);
            g.DrawLine(Pens.Black, chartArea.X, chartArea.Y + chartArea.Height,
                chartArea.X + chartArea.Width, chartArea.Y + chartArea.Height);
        }

        private void GenerateReportButton_Click(object sender, EventArgs e)
        {
            try
            {
                var allResumes = _managementService.LoadResumes();
                if (allResumes.Count == 0)
                {
                    MessageBox.Show("暂无简历数据，无法生成报告。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (var saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Word文档|*.docx";
                    saveFileDialog.Title = "保存分析报告";
                    saveFileDialog.FileName = $"简历数据分析报告_{DateTime.Now:yyyyMMdd_HHmmss}.docx";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        GenerateAnalysisReport(allResumes, saveFileDialog.FileName);
                        MessageBox.Show("报告生成成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"生成报告失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void GenerateAnalysisReport(List<Resume> allResumes, string filePath)
        {
            using (var document = Xceed.Words.NET.DocX.Create(filePath))
            {
                // 标题
                var title = document.InsertParagraph("简历数据分析报告");
                title.Font("微软雅黑").FontSize(18).Bold().Alignment = Xceed.Document.NET.Alignment.center;
                title.SpacingAfter(20);

                // 生成时间
                var dateInfo = document.InsertParagraph($"生成时间: {DateTime.Now:yyyy年MM月dd日 HH:mm:ss}");
                dateInfo.Font("微软雅黑").FontSize(10).Color(System.Drawing.Color.Gray);
                dateInfo.Alignment = Xceed.Document.NET.Alignment.center;
                dateInfo.SpacingAfter(30);

                // 一、数据概览
                var section1 = document.InsertParagraph("一、数据概览");
                section1.Font("微软雅黑").FontSize(14).Bold();
                section1.SpacingAfter(10);

                var earliest = allResumes.Min(r => r.ImportTime);
                var latest = allResumes.Max(r => r.ImportTime);
                var totalResumes = allResumes.Count;

                var overviewText = $"本报告共分析 {totalResumes} 份简历数据。\n" +
                                  $"数据导入时间范围：{earliest:yyyy-MM-dd} 至 {latest:yyyy-MM-dd}。\n" +
                                  $"以下是详细的数据统计分析结果：\n";
                var overview = document.InsertParagraph(overviewText);
                overview.Font("微软雅黑").FontSize(11);
                overview.SpacingAfter(20);

                // 二、性别分布
                var section2 = document.InsertParagraph("二、性别分布分析");
                section2.Font("微软雅黑").FontSize(14).Bold();
                section2.SpacingAfter(10);

                var genderGroups = allResumes
                    .GroupBy(r => string.IsNullOrWhiteSpace(r.Gender) ? "未填写" : r.Gender)
                    .OrderByDescending(g => g.Count())
                    .ToList();

                var genderText = new StringBuilder();
                foreach (var group in genderGroups)
                {
                    var percentage = group.Count() * 100.0 / totalResumes;
                    genderText.AppendLine($"{group.Key}: {group.Count()} 人，占比 {percentage:F1}%");
                }
                var genderPara = document.InsertParagraph(genderText.ToString());
                genderPara.Font("微软雅黑").FontSize(11);
                genderPara.SpacingAfter(20);

                // 三、年龄分布
                var section3 = document.InsertParagraph("三、年龄分布分析");
                section3.Font("微软雅黑").FontSize(14).Bold();
                section3.SpacingAfter(10);

                var ageGroups = CalculateAgeDistribution(allResumes);
                var totalWithAge = ageGroups.Values.Sum();
                var ageText = new StringBuilder();
                if (totalWithAge > 0)
                {
                    ageText.AppendLine($"有效年龄数据: {totalWithAge} 人（占总数的 {totalWithAge * 100.0 / totalResumes:F1}%）\n");
                    foreach (var kvp in ageGroups.Where(kvp => kvp.Value > 0).OrderByDescending(kvp => kvp.Value))
                    {
                        var percentage = kvp.Value * 100.0 / totalWithAge;
                        ageText.AppendLine($"{kvp.Key}: {kvp.Value} 人，占比 {percentage:F1}%");
                    }
                }
                else
                {
                    ageText.AppendLine("暂无有效的年龄数据。");
                }
                var agePara = document.InsertParagraph(ageText.ToString());
                agePara.Font("微软雅黑").FontSize(11);
                agePara.SpacingAfter(20);

                // 四、学历分布
                var section4 = document.InsertParagraph("四、学历分布分析");
                section4.Font("微软雅黑").FontSize(14).Bold();
                section4.SpacingAfter(10);

                var educationGroups = CalculateEducationDistribution(allResumes);
                var educationText = new StringBuilder();
                foreach (var kvp in educationGroups.OrderByDescending(kvp => kvp.Value))
                {
                    var percentage = kvp.Value * 100.0 / totalResumes;
                    educationText.AppendLine($"{kvp.Key}: {kvp.Value} 人，占比 {percentage:F1}%");
                }
                var educationPara = document.InsertParagraph(educationText.ToString());
                educationPara.Font("微软雅黑").FontSize(11);
                educationPara.SpacingAfter(20);

                // 五、技能分布
                var section5 = document.InsertParagraph("五、技能分布分析");
                section5.Font("微软雅黑").FontSize(14).Bold();
                section5.SpacingAfter(10);

                var skillGroups = CalculateSkillDistribution(allResumes);
                var topSkills = skillGroups.OrderByDescending(kvp => kvp.Value).Take(15).ToList();
                var skillText = new StringBuilder();
                if (topSkills.Count > 0)
                {
                    skillText.AppendLine($"共识别出 {skillGroups.Count} 种不同技能，以下是出现频率最高的 Top 15：\n");
                    int rank = 1;
                    foreach (var kvp in topSkills)
                    {
                        skillText.AppendLine($"{rank}. {kvp.Key}: 出现在 {kvp.Value} 份简历中");
                        rank++;
                    }
                }
                else
                {
                    skillText.AppendLine("暂无技能数据。");
                }
                var skillPara = document.InsertParagraph(skillText.ToString());
                skillPara.Font("微软雅黑").FontSize(11);
                skillPara.SpacingAfter(20);

                // 六、时间趋势
                var section6 = document.InsertParagraph("六、导入时间趋势分析");
                section6.Font("微软雅黑").FontSize(14).Bold();
                section6.SpacingAfter(10);

                var monthGroups = allResumes
                    .GroupBy(r => r.ImportTime.ToString("yyyy-MM"))
                    .OrderBy(g => g.Key)
                    .ToList();

                var trendText = new StringBuilder();
                trendText.AppendLine("按月份统计简历导入情况：\n");
                foreach (var group in monthGroups)
                {
                    var percentage = group.Count() * 100.0 / totalResumes;
                    trendText.AppendLine($"{group.Key}: {group.Count()} 份，占比 {percentage:F1}%");
                }

                if (monthGroups.Count > 1)
                {
                    var firstMonth = monthGroups.First().Count();
                    var lastMonth = monthGroups.Last().Count();
                    if (firstMonth > 0)
                    {
                        var growthRate = (lastMonth - firstMonth) * 100.0 / firstMonth;
                        trendText.AppendLine($"\n趋势分析: 从 {monthGroups.First().Key} 到 {monthGroups.Last().Key}，");
                        trendText.AppendLine($"月度导入量从 {firstMonth} 份变化为 {lastMonth} 份，");
                        trendText.AppendLine($"变化幅度: {(growthRate >= 0 ? "+" : "")}{growthRate:F1}%");
                    }
                }

                var trendPara = document.InsertParagraph(trendText.ToString());
                trendPara.Font("微软雅黑").FontSize(11);
                trendPara.SpacingAfter(20);

                // 七、总结
                var section7 = document.InsertParagraph("七、总结");
                section7.Font("微软雅黑").FontSize(14).Bold();
                section7.SpacingAfter(10);

                var summaryText = new StringBuilder();
                summaryText.AppendLine($"通过对 {totalResumes} 份简历数据的分析，我们可以得出以下结论：\n");

                if (genderGroups.Count > 0)
                {
                    var maxGender = genderGroups.First();
                    summaryText.AppendLine($"1. 性别分布：{maxGender.Key}占比最高（{maxGender.Count() * 100.0 / totalResumes:F1}%）。\n");
                }

                var maxEducation = educationGroups.OrderByDescending(kvp => kvp.Value).FirstOrDefault();
                if (maxEducation.Value > 0)
                {
                    summaryText.AppendLine($"2. 学历分布：{maxEducation.Key}学历占比最高（{maxEducation.Value * 100.0 / totalResumes:F1}%）。\n");
                }

                if (topSkills.Count > 0)
                {
                    var topSkill = topSkills.First();
                    summaryText.AppendLine($"3. 技能分析：{topSkill.Key}是最热门的技能，出现在 {topSkill.Value} 份简历中。\n");
                }

                var summaryPara = document.InsertParagraph(summaryText.ToString());
                summaryPara.Font("微软雅黑").FontSize(11);
                summaryPara.SpacingAfter(30);

                // 页脚
                var footer = document.InsertParagraph($"报告生成于 {DateTime.Now:yyyy年MM月dd日 HH:mm:ss} | 智能简历解析系统");
                footer.Font("微软雅黑").FontSize(9).Color(System.Drawing.Color.Gray);
                footer.Alignment = Xceed.Document.NET.Alignment.center;

                document.Save();
            }
        }

        private void RefreshExportButton_Click(object sender, EventArgs e)
        {
            LoadExportResumes();
        }

        private void LoadExportResumes()
        {
            var exportTabControl = panelExport.Controls.Find("exportTabControl", false).FirstOrDefault() as TabControl;
            if (exportTabControl == null || exportTabControl.TabPages.Count == 0)
            {
                return;
            }

            var directExportTab = exportTabControl.TabPages[0];
            var exportList = directExportTab.Controls.Find("exportCheckedListBox", false).FirstOrDefault() as CheckedListBox;
            var summaryLabel = directExportTab.Controls.Find("exportSummaryLabel", false).FirstOrDefault() as Label;

            if (exportList == null || summaryLabel == null)
            {
                return;
            }

            try
            {
                _exportResumes = _managementService.LoadResumes()
                    .OrderByDescending(r => r.ImportTime)
                    .ToList();
            }
            catch (Exception ex)
            {
                summaryLabel.Text = $"加载数据失败: {ex.Message}";
                exportList.Items.Clear();
                return;
            }

            exportList.Items.Clear();

            foreach (var resume in _exportResumes)
            {
                exportList.Items.Add($"{resume.Name} - {resume.FileName}");
            }

            summaryLabel.Text = _exportResumes.Count > 0
                ? $"可导出简历数量: {_exportResumes.Count}"
                : "暂无可导出的简历。";
        }

        private void ExportSelectedButton_Click(object sender, EventArgs e)
        {
            // 从TabControl中查找控件
            var exportTabControl = panelExport.Controls.Find("exportTabControl", false).FirstOrDefault() as TabControl;
            if (exportTabControl == null || exportTabControl.TabPages.Count == 0)
            {
                MessageBox.Show("界面初始化失败，无法导出。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var directExportTab = exportTabControl.TabPages[0];
            var exportList = directExportTab.Controls.Find("exportCheckedListBox", false).FirstOrDefault() as CheckedListBox;
            var formatCombo = directExportTab.Controls.Find("exportFormatCombo", false).FirstOrDefault() as ComboBox;

            if (exportList == null || formatCombo == null)
            {
                MessageBox.Show("界面初始化失败，无法导出。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (exportList.CheckedItems.Count == 0)
            {
                MessageBox.Show("请先勾选需要导出的简历。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var selectedResumes = exportList.CheckedIndices
                .Cast<int>()
                .Where(index => index >= 0 && index < _exportResumes.Count)
                .Select(index => _exportResumes[index])
                .ToList();

            if (selectedResumes.Count == 0)
            {
                MessageBox.Show("选中的简历列表为空。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var selectedFormat = formatCombo.SelectedItem?.ToString() ?? "JSON";
            using (var saveFileDialog = new SaveFileDialog())
            {
                switch (selectedFormat.ToUpper())
                {
                    case "CSV":
                        saveFileDialog.Filter = "CSV 文件|*.csv";
                        saveFileDialog.FileName = $"简历导出_{DateTime.Now:yyyyMMdd_HHmmss}.csv";
                        break;
                    case "WORD":
                        saveFileDialog.Filter = "Word文档|*.docx";
                        saveFileDialog.FileName = $"简历导出_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
                        break;
                    default: // JSON
                        saveFileDialog.Filter = "JSON 文件|*.json";
                        saveFileDialog.FileName = $"简历导出_{DateTime.Now:yyyyMMdd_HHmmss}.json";
                        break;
                }

                if (saveFileDialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                try
                {
                    switch (selectedFormat.ToUpper())
                    {
                        case "CSV":
                            ExportToCsv(selectedResumes, saveFileDialog.FileName);
                            break;
                        case "WORD":
                            ExportToWord(selectedResumes, saveFileDialog.FileName);
                            break;
                        default: // JSON
                            ExportToJson(selectedResumes, saveFileDialog.FileName);
                            break;
                    }

                    MessageBox.Show("导出成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void LoadReviewNotes()
        {
            try
            {
                if (File.Exists(_reviewDataFile))
                {
                    var json = File.ReadAllText(_reviewDataFile);
                    _reviewNotes = JsonSerializer.Deserialize<Dictionary<string, string>>(json) ?? new Dictionary<string, string>();
                }
            }
            catch
            {
                _reviewNotes = new Dictionary<string, string>();
            }
        }

        private void SaveReviewNotes()
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };
            var json = JsonSerializer.Serialize(_reviewNotes, options);
            File.WriteAllText(_reviewDataFile, json, Encoding.UTF8);
        }

        private void UpdateReviewStatus(string message, bool isError)
        {
            var statusLabel = panelReview.Controls.Find("reviewStatusLabel", false).FirstOrDefault() as Label;
            if (statusLabel == null)
            {
                return;
            }

            statusLabel.ForeColor = isError ? Color.Red : Color.Gray;
            statusLabel.Text = message;
        }

        private void ExportToJson(IEnumerable<Resume> resumes, string filePath)
        {
            var options = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };
            var json = JsonSerializer.Serialize(resumes, options);
            File.WriteAllText(filePath, json, Encoding.UTF8);
        }

        private void ExportToCsv(IEnumerable<Resume> resumes, string filePath)
        {
            var allDirectories = _managementService.GetAllDirectories();
            var sb = new StringBuilder();
            //sb.AppendLine("姓名,文件名,手机,邮箱,导入时间,目录");
            // 写入最终版CSV表头
            sb.AppendLine("ID,姓名,性别,出生日期,家庭住址,手机号码,邮箱,身份证号,最高学历院校,最高学历专业,第一学历院校,第一学历专业,技能列表,工作经历（公司-岗位-时间）,原始文件路径,文件名,导入时间,所属目录ID,所属目录名称");

            foreach (var resume in resumes)
            {
                string directoryName = allDirectories
            .FirstOrDefault(d => d.Id == resume.DirectoryId)?
            .Name ?? "未归类";
                // 2. 处理日期字段（格式化+null兼容）
                string birthDateStr = resume.BirthDate.HasValue
                    ? resume.BirthDate.Value.ToString("yyyy-MM-dd")
                    : string.Empty;

                // 3. 处理技能列表（数组→字符串）
                string skillsStr = string.Join(",", resume.Skills ?? Enumerable.Empty<string>());

                // 4. 处理工作经历（数组→结构化字符串）
                var workExpList = new List<string>();
                foreach (var exp in resume.WorkExperiences ?? Enumerable.Empty<WorkExperience>())
                {
                    string startDate = exp.StartDate.HasValue ? exp.StartDate.Value.ToString("yyyy-MM-dd") : "无";
                    string endDate = exp.IsCurrentJob ? "当前" : (exp.EndDate.HasValue ? exp.EndDate.Value.ToString("yyyy-MM-dd") : "无");
                    // 单条经历格式：公司-岗位-开始时间-结束时间
                    string expStr = $"{exp.Company}-{exp.Position}-{startDate}-{endDate}";
                    workExpList.Add(expStr);
                }
                string workExpStr = string.Join("||", workExpList); // 多经历用“||”分隔，避免与内部“-”冲突

                // 5. 拼接所有字段（顺序与表头严格一致）
                var fields = new[]
                {
            EscapeCsv(resume.Id),
            EscapeCsv(resume.Name),
            EscapeCsv(resume.Gender),
            EscapeCsv(birthDateStr),
            EscapeCsv(resume.Address),
            EscapeCsv(resume.Phone),
            EscapeCsv(resume.Email),
            EscapeCsv(resume.IdCard),
            EscapeCsv(resume.HighestEducationSchool),
            EscapeCsv(resume.HighestEducationMajor),
            EscapeCsv(resume.FirstEducationSchool),
            EscapeCsv(resume.FirstEducationMajor),
            EscapeCsv(skillsStr),
            EscapeCsv(workExpStr),
            EscapeCsv(resume.OriginalFilePath),
            EscapeCsv(resume.FileName),
            EscapeCsv(resume.ImportTime.ToString("yyyy-MM-dd HH:mm:ss")),
            EscapeCsv(resume.DirectoryId),
            EscapeCsv(directoryName)
        };

                sb.AppendLine(string.Join(",", fields));
            }
            /* var fields = new[]
             {
                 EscapeCsv(resume.Name),
                 EscapeCsv(resume.FileName),
                 EscapeCsv(resume.Phone),
                 EscapeCsv(resume.Email),
                 EscapeCsv(resume.ImportTime.ToString("yyyy-MM-dd HH:mm:ss")),
                  EscapeCsv(directoryName)
             };
             sb.AppendLine(string.Join(",", fields));
         }*/

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }

        private static string EscapeCsv(string? value)
        {
            value ??= string.Empty;
            if (value.Contains('"'))
            {
                value = value.Replace("\"", "\"\"");
            }

            if (value.Contains(',') || value.Contains('\n') || value.Contains('\r'))
            {
                value = $"\"{value}\"";
            }

            return value;
        }

        private void ConvertFormatButton_Click(object sender, EventArgs e)
        {
            // 从TabControl中查找控件
            var exportTabControl = panelExport.Controls.Find("exportTabControl", false).FirstOrDefault() as TabControl;
            if (exportTabControl == null || exportTabControl.TabPages.Count < 2)
            {
                MessageBox.Show("界面初始化失败，无法进行格式转换。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var convertTab = exportTabControl.TabPages[1];
            var sourceFormatCombo = convertTab.Controls.Find("sourceFormatCombo", false).FirstOrDefault() as ComboBox;
            var targetFormatCombo = convertTab.Controls.Find("targetFormatCombo", false).FirstOrDefault() as ComboBox;
            var statusLabel = convertTab.Controls.Find("convertStatusLabel", false).FirstOrDefault() as Label;

            if (sourceFormatCombo == null || targetFormatCombo == null)
            {
                MessageBox.Show("界面初始化失败，无法进行格式转换。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var sourceFormat = sourceFormatCombo.SelectedItem?.ToString() ?? "JSON";
            var targetFormat = targetFormatCombo.SelectedItem?.ToString() ?? "CSV";

            if (sourceFormat == targetFormat)
            {
                MessageBox.Show("源格式和目标格式不能相同。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // 选择源文件
            using (var openFileDialog = new OpenFileDialog())
            {
                if (sourceFormat == "JSON")
                {
                    openFileDialog.Filter = "JSON 文件|*.json";
                }
                else if (sourceFormat == "CSV")
                {
                    openFileDialog.Filter = "CSV 文件|*.csv";
                }
                else
                {
                    MessageBox.Show("不支持的源格式。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                openFileDialog.Title = "选择要转换的源文件";

                if (openFileDialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                // 读取源文件
                List<Resume> resumes;
                try
                {
                    if (sourceFormat == "JSON")
                    {
                        resumes = LoadResumesFromJson(openFileDialog.FileName);
                    }
                    else // CSV
                    {
                        resumes = LoadResumesFromCsv(openFileDialog.FileName);
                    }

                    if (resumes == null || resumes.Count == 0)
                    {
                        MessageBox.Show("未能从源文件中读取到简历数据。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        if (statusLabel != null) statusLabel.Text = "转换失败：未读取到数据";
                        return;
                    }

                    if (statusLabel != null) statusLabel.Text = $"成功读取 {resumes.Count} 份简历数据，准备转换...";
                    Application.DoEvents();

                    // 选择保存位置
                    using (var saveFileDialog = new SaveFileDialog())
                    {
                        switch (targetFormat.ToUpper())
                        {
                            case "JSON":
                                saveFileDialog.Filter = "JSON 文件|*.json";
                                saveFileDialog.FileName = Path.GetFileNameWithoutExtension(openFileDialog.FileName) + "_转换.json";
                                break;
                            case "CSV":
                                saveFileDialog.Filter = "CSV 文件|*.csv";
                                saveFileDialog.FileName = Path.GetFileNameWithoutExtension(openFileDialog.FileName) + "_转换.csv";
                                break;
                            case "WORD":
                                saveFileDialog.Filter = "Word文档|*.docx";
                                saveFileDialog.FileName = Path.GetFileNameWithoutExtension(openFileDialog.FileName) + "_转换.docx";
                                break;
                            default:
                                MessageBox.Show("不支持的目标格式。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                        }

                        saveFileDialog.Title = "选择保存位置";

                        if (saveFileDialog.ShowDialog() != DialogResult.OK)
                        {
                            return;
                        }

                        // 执行转换
                        try
                        {
                            switch (targetFormat.ToUpper())
                            {
                                case "JSON":
                                    ExportToJson(resumes, saveFileDialog.FileName);
                                    break;
                                case "CSV":
                                    ExportToCsv(resumes, saveFileDialog.FileName);
                                    break;
                                case "WORD":
                                    ExportToWord(resumes, saveFileDialog.FileName);
                                    break;
                            }

                            if (statusLabel != null)
                            {
                                statusLabel.Text = $"转换成功！已将 {resumes.Count} 份简历从 {sourceFormat} 格式转换为 {targetFormat} 格式。";
                                statusLabel.ForeColor = Color.Green;
                            }
                            MessageBox.Show($"转换成功！\n已将 {resumes.Count} 份简历从 {sourceFormat} 格式转换为 {targetFormat} 格式。",
                                "转换成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception ex)
                        {
                            if (statusLabel != null)
                            {
                                statusLabel.Text = $"转换失败: {ex.Message}";
                                statusLabel.ForeColor = Color.Red;
                            }
                            MessageBox.Show($"转换失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (statusLabel != null)
                    {
                        statusLabel.Text = $"读取源文件失败: {ex.Message}";
                        statusLabel.ForeColor = Color.Red;
                    }
                    MessageBox.Show($"读取源文件失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private List<Resume> LoadResumesFromJson(string filePath)
        {
            try
            {
                var json = File.ReadAllText(filePath, Encoding.UTF8);
                var options = new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                };
                return JsonSerializer.Deserialize<List<Resume>>(json, options) ?? new List<Resume>();
            }
            catch (Exception ex)
            {
                throw new Exception($"读取JSON文件失败: {ex.Message}");
            }
        }

        private List<Resume> LoadResumesFromCsv(string filePath)
        {
            try
            {
                var resumes = new List<Resume>();
                var lines = File.ReadAllLines(filePath, Encoding.UTF8);

                var allDirectories = _managementService.GetAllDirectories();
                if (lines.Length < 2)
                {
                    return resumes; // 至少需要表头和数据行
                }

                // 解析表头
                var headers = lines[0].Split(',').Select(h => h.Trim('"').Trim()).ToArray();

                // 解析数据行
                for (int i = 1; i < lines.Length; i++)
                {
                    var line = lines[i];
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    var values = ParseCsvLine(line);
                    if (values.Count < headers.Length) continue;

                    var resume = new Resume();

                    /*for (int j = 0; j < headers.Length && j < values.Count; j++)
                    {
                        var header = headers[j].ToLower();
                        var value = values[j].Trim('"').Trim();

                        switch (header)
                        {
                            case "姓名":
                            case "name":
                                resume.Name = value;
                                break;
                            case "文件名":
                            case "filename":
                                resume.FileName = value;
                                break;
                            case "手机":
                            case "phone":
                                resume.Phone = value;
                                break;
                            case "邮箱":
                            case "email":
                                resume.Email = value;
                                break;
                            case "导入时间":
                            case "importtime":
                                if (DateTime.TryParse(value, out var importTime))
                                {
                                    resume.ImportTime = importTime;
                                }
                                break;
                            case "目录":
                            case "directory":
                            var matchedDir = allDirectories.FirstOrDefault(d =>
                                d.Name.Equals(value, StringComparison.OrdinalIgnoreCase));
                            resume.DirectoryId = matchedDir?.Id;
                                break;*/

                    // 字段映射（按表头顺序匹配）
                    for (int j = 0; j < headers.Length && j < values.Count; j++)
                    {
                        var header = headers[j];
                        var value = values[j].Trim('"').Trim(); // 去除前后引号和空格

                        switch (header)
                        {
                            case "id":
                                resume.Id = value;
                                break;
                            case "姓名":
                            case "name":
                                resume.Name = value;
                                break;
                            case "性别":
                            case "gender":
                                resume.Gender = value;
                                break;
                            case "出生日期":
                            case "birthdate":
                                if (DateTime.TryParse(value, out var birthDate))
                                    resume.BirthDate = birthDate;
                                break;
                            case "家庭住址":
                            case "address":
                                resume.Address = value;
                                break;
                            case "手机号码":
                            case "phone":
                                resume.Phone = value;
                                break;
                            case "邮箱":
                            case "email":
                                resume.Email = value;
                                break;
                            case "身份证号":
                            case "idcard":
                                resume.IdCard = value;
                                break;
                            case "最高学历院校":
                            case "highesteducationschool":
                                resume.HighestEducationSchool = value;
                                break;
                            case "最高学历专业":
                            case "highesteducationmajor":
                                resume.HighestEducationMajor = value;
                                break;
                            case "第一学历院校":
                            case "firsteducationschool":
                                resume.FirstEducationSchool = value;
                                break;
                            case "第一学历专业":
                            case "firsteducationmajor":
                                resume.FirstEducationMajor = value;
                                break;
                            case "技能列表":
                            case "skills":
                                // 字符串→技能列表（逗号分隔）
                                if (!string.IsNullOrEmpty(value))
                                    resume.Skills = value.Split(',', StringSplitOptions.RemoveEmptyEntries).ToList();
                                break;
                            case "工作经历（公司-岗位-时间）":
                            case "workexperiences":
                                // 字符串→工作经历列表（||分隔多经历，-分隔字段）
                                if (!string.IsNullOrEmpty(value))
                                {
                                    var expItems = value.Split("||", StringSplitOptions.RemoveEmptyEntries);
                                    foreach (var expStr in expItems)
                                    {
                                        var expParts = expStr.Split('-', 4); // 按前3个“-”分割（避免字段含“-”）
                                        var workExp = new WorkExperience();

                                        workExp.Company = expParts.Length > 0 ? expParts[0].Trim() : string.Empty;
                                        workExp.Position = expParts.Length > 1 ? expParts[1].Trim() : string.Empty;

                                        // 处理开始时间
                                        if (expParts.Length > 2 && expParts[2].Trim() != "无" && DateTime.TryParse(expParts[2].Trim(), out var startDate))
                                            workExp.StartDate = startDate;

                                        // 处理结束时间和是否当前工作
                                        if (expParts.Length > 3)
                                        {
                                            string endDateStr = expParts[3].Trim();
                                            if (endDateStr == "当前")
                                            {
                                                workExp.IsCurrentJob = true;
                                                workExp.EndDate = null;
                                            }
                                            else if (endDateStr != "无" && DateTime.TryParse(endDateStr, out var endDate))
                                            {
                                                workExp.EndDate = endDate;
                                            }
                                        }

                                        resume.WorkExperiences.Add(workExp);
                                    }
                                }
                                break;
                            case "原始文件路径":
                            case "originalfilepath":
                                resume.OriginalFilePath = value;
                                break;
                            case "文件名":
                            case "filename":
                                resume.FileName = value;
                                break;
                            case "导入时间":
                            case "importtime":
                                if (DateTime.TryParse(value, out var importTime))
                                    resume.ImportTime = importTime;
                                break;
                            case "所属目录id":
                            case "directoryid":
                                resume.DirectoryId = value;
                                break;
                            case "所属目录名称":
                            case "directoryname":
                                break;
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(resume.Name))
                    {
                        resumes.Add(resume);
                    }
                }

                return resumes;
            }
            catch (Exception ex)
            {
                throw new Exception($"读取CSV文件失败: {ex.Message}");
            }
        }

        private List<string> ParseCsvLine(string line)
        {
            var values = new List<string>();
            var currentValue = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        // 转义的引号
                        currentValue.Append('"');
                        i++; // 跳过下一个引号
                    }
                    else
                    {
                        // 切换引号状态
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    // 字段结束
                    values.Add(currentValue.ToString());
                    currentValue.Clear();
                }
                else
                {
                    currentValue.Append(c);
                }
            }

            // 添加最后一个字段
            values.Add(currentValue.ToString());

            return values;
        }

        private void ExportToWord(IEnumerable<Resume> resumes, string filePath)
        {
            var allDirectories = _managementService.GetAllDirectories();
            using (var document = Xceed.Words.NET.DocX.Create(filePath))
            {
                var title = document.InsertParagraph("简历数据导出");
                title.Font("微软雅黑").FontSize(16).Bold().Alignment = Xceed.Document.NET.Alignment.center;
                title.SpacingAfter(20);

                var dateInfo = document.InsertParagraph($"导出时间: {DateTime.Now:yyyy年MM月dd日 HH:mm:ss}");
                dateInfo.Font("微软雅黑").FontSize(10).Color(System.Drawing.Color.Gray);
                dateInfo.Alignment = Xceed.Document.NET.Alignment.center;
                dateInfo.SpacingAfter(30);

                int index = 1;
                foreach (var resume in resumes)
                {
                    var resumeTitle = document.InsertParagraph($"简历 {index}");
                    resumeTitle.Font("微软雅黑").FontSize(14).Bold();
                    resumeTitle.SpacingAfter(10);

                    var content = new StringBuilder();
                    content.AppendLine($"姓名: {resume.Name}");
                    content.AppendLine($"性别: {resume.Gender}");
                    content.AppendLine($"手机: {resume.Phone}");
                    content.AppendLine($"邮箱: {resume.Email}");
                    content.AppendLine($"地址: {resume.Address}");
                    if (resume.BirthDate.HasValue)
                    {
                        content.AppendLine($"出生日期: {resume.BirthDate.Value:yyyy-MM-dd}");
                    }
                    content.AppendLine($"最高学历学校: {resume.HighestEducationSchool}");
                    content.AppendLine($"最高学历专业: {resume.HighestEducationMajor}");
                    content.AppendLine($"第一学历学校: {resume.FirstEducationSchool}");
                    content.AppendLine($"第一学历专业: {resume.FirstEducationMajor}");
                    content.AppendLine($"导入时间: {resume.ImportTime:yyyy-MM-dd HH:mm:ss}");
                    string directoryName = allDirectories
                                   .FirstOrDefault(d => d.Id == resume.DirectoryId)?
                                   .Name ?? "未归类";
                    content.AppendLine($"目录: {directoryName}");

                    if (resume.Skills.Count > 0)
                    {
                        content.AppendLine($"技能: {string.Join(", ", resume.Skills)}");
                    }

                    if (resume.WorkExperiences.Count > 0)
                    {
                        content.AppendLine("工作经历:");
                        foreach (var work in resume.WorkExperiences)
                        {
                            content.AppendLine($"  - {work.Company} - {work.Position}");
                        }
                    }

                    var resumeContent = document.InsertParagraph(content.ToString());
                    resumeContent.Font("微软雅黑").FontSize(11);
                    resumeContent.SpacingAfter(20);

                    index++;
                }

                document.Save();
            }
        }

        private void CheckDuplicateButton_Click(object sender, EventArgs e)
        {
            var checkNameCheckBox = panelDuplicate.Controls.Find("checkNameCheckBox", false).FirstOrDefault() as CheckBox;
            var checkPhoneCheckBox = panelDuplicate.Controls.Find("checkPhoneCheckBox", false).FirstOrDefault() as CheckBox;
            var checkEmailCheckBox = panelDuplicate.Controls.Find("checkEmailCheckBox", false).FirstOrDefault() as CheckBox;
            var checkIdCardCheckBox = panelDuplicate.Controls.Find("checkIdCardCheckBox", false).FirstOrDefault() as CheckBox;
            var duplicateResultTextBox = panelDuplicate.Controls.Find("duplicateResultTextBox", false).FirstOrDefault() as TextBox;

            bool checkName = checkNameCheckBox?.Checked ?? false;
            bool checkPhone = checkPhoneCheckBox?.Checked ?? false;
            bool checkEmail = checkEmailCheckBox?.Checked ?? false;
            bool checkIdCard = checkIdCardCheckBox?.Checked ?? false;

            if (!checkName && !checkPhone && !checkEmail && !checkIdCard)
            {
                MessageBox.Show("请至少选择一个查重属性。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _currentDuplicates = _managementService.FindDuplicateResumes(checkName, checkPhone, checkEmail, checkIdCard);

            var sb = new StringBuilder();
            sb.AppendLine($"查重完成！发现 {_currentDuplicates.Count} 组重复简历\n");

            foreach (var duplicate in _currentDuplicates)
            {
                sb.AppendLine($"\n{duplicate.Key}");
                sb.AppendLine($"重复数量: {duplicate.Value.Count} 份\n");

                foreach (var resume in duplicate.Value)
                {
                    sb.AppendLine($"    姓名: {resume.Name}");
                    sb.AppendLine($"    文件名: {resume.FileName}");
                    sb.AppendLine($"    手机: {resume.Phone}");
                    sb.AppendLine($"    邮箱: {resume.Email}");
                    sb.AppendLine($"    身份证号: {resume.IdCard}");
                    sb.AppendLine($"    导入时间: {resume.ImportTime:yyyy-MM-dd HH:mm:ss}\n");
                }

                sb.AppendLine(new string('-', 60));
            }

            if (duplicateResultTextBox != null)
            {
                duplicateResultTextBox.Text = sb.ToString();
            }

            MessageBox.Show($"查重完成！发现 {_currentDuplicates.Count} 组重复简历", "查重结果", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ExportTxtButton_Click(object sender, EventArgs e)
        {
            if (_currentDuplicates == null || _currentDuplicates.Count == 0)
            {
                MessageBox.Show("请先执行查重操作。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (var saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "文本文件|*.txt";
                saveFileDialog.Title = "导出查重结果";
                saveFileDialog.FileName = $"简历查重结果_{DateTime.Now:yyyyMMdd_HHmmss}.txt";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        _managementService.ExportDuplicatesToTxt(_currentDuplicates, saveFileDialog.FileName);
                        MessageBox.Show("导出成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void ExportWordButton_Click(object sender, EventArgs e)
        {
            if (_currentDuplicates == null || _currentDuplicates.Count == 0)
            {
                MessageBox.Show("请先执行查重操作。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (var saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Word文档|*.docx";
                saveFileDialog.Title = "导出查重结果";
                saveFileDialog.FileName = $"简历查重结果_{DateTime.Now:yyyyMMdd_HHmmss}.docx";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        _managementService.ExportDuplicatesToWord(_currentDuplicates, saveFileDialog.FileName);
                        MessageBox.Show("导出成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"导出失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

    }
}
