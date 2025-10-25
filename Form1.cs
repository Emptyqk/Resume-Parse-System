using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using 页面.Models;
using 页面.Services;

namespace 页面
{
    public partial class Form1 : Form
    {
        private readonly ResumeManagementService _managementService;
        private List<Resume> _currentResumes = new List<Resume>();
        private List<ResumeDirectory> _directories = new List<ResumeDirectory>();

        public Form1()
        {
            InitializeComponent();
            _managementService = new ResumeManagementService();
            
            InitializeUI();
            panelAdd.Show();
        }

        private void InitializeUI()
        {          
            // 在panelAdd添加已上传简历列表
            var uploadedLabel = new Label
            {
                Text = "已上传的简历:",
                Location = new Point(10, 90),
                Size = new Size(150, 30),
                Font = new Font("微软雅黑", 9, FontStyle.Bold)
            };
            panelAdd.Controls.Add(uploadedLabel);
            
            var uploadedListBox = new ListBox
            {
                Name = "uploadedListBox",
                Location = new Point(10, 125),
                Size = new Size(560, 200),
                SelectionMode = SelectionMode.One
                //一次只能选一个
            };
            uploadedListBox.DoubleClick += UploadedListBox_DoubleClick;
            panelAdd.Controls.Add(uploadedListBox);

            var importButton = new Button
            {
                Text = "选择文件导入",
                Location = new Point(10, 10),
                Size = new Size(150, 40)
            };
            importButton.Click += ImportButton_Click;
            panelAdd.Controls.Add(importButton);
            
            // 添加拖拽提示标签
            var dragLabel = new Label
            {
                Text = "或者将文件拖拽到此区域",
                Location = new Point(10, 60),
                Size = new Size(200, 30)
            };
            panelAdd.Controls.Add(dragLabel);
            
            // 添加删除按钮
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

            // 添加查看详情按钮
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

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadResumes();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            panelAdd.Show();
            panelSearch.Hide();
            panelDuplicate.Hide();
            LoadResumes();  // 切换到添加面板时刷新简历列表
        }

        private void buttonSearchResume_Click(object sender, EventArgs e)
        {
             panelAdd.Hide();
            panelSearch.Show();
            panelDuplicate.Hide();
        }

        private void buttonCheckDuplicate_Click(object sender, EventArgs e)
        {
            panelAdd.Hide();
            panelSearch.Hide();
            panelDuplicate.Show();
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
                    ImportFiles(openFileDialog.FileNames);
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
                    ImportFiles(validFiles);
                }
                else
                {
                    MessageBox.Show("请选择Word或PDF格式的简历文件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void ImportFiles(string[] filePaths)
        {
            try
            {
                var importedResumes = _managementService.ImportResumes(filePaths.ToList(), false);
                
                if (importedResumes.Count > 0)
                {
                    MessageBox.Show($"成功导入 {importedResumes.Count} 份简历。", "导入成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadResumes();
                }
                else
                {
                    MessageBox.Show("导入失败", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            sb.AppendLine($"目录: {resume.Directory}");

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

            // 搜索字段选择
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
            searchFieldCombo.Items.AddRange(new object[] { "全部", "姓名", "文件名", "手机", "邮箱" });
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

            // 开始日期
            var startDateLabel = new Label
            {
                Text = "开始:",
                Location = new Point(10, 85),
                Size = new Size(60, 30)
            };
            panelSearch.Controls.Add(startDateLabel);

            var startDatePicker = new DateTimePicker
            {
                Name = "startDatePicker",
                Location = new Point(70, 83),
                Size = new Size(150, 25),
                Format = DateTimePickerFormat.Short
            };
            startDatePicker.Value = DateTime.Now.AddMonths(-1);
            panelSearch.Controls.Add(startDatePicker);

            var startDateCheck = new CheckBox
            {
                Name = "startDateCheck",
                Text = "是",
                Location = new Point(225, 85),
                Size = new Size(60, 30)
            };
            panelSearch.Controls.Add(startDateCheck);

            // 结束日期
            var endDateLabel = new Label
            {
                Text = "结束:",
                Location = new Point(290, 85),
                Size = new Size(60, 30)
            };
            panelSearch.Controls.Add(endDateLabel);

            var endDatePicker = new DateTimePicker
            {
                Name = "endDatePicker",
                Location = new Point(355, 83),
                Size = new Size(150, 25),
                Format = DateTimePickerFormat.Short
            };
            panelSearch.Controls.Add(endDatePicker);

            var endDateCheck = new CheckBox
            {
                Name = "endDateCheck",
                Text = "是",
                Location = new Point(520, 85),
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

            // 结果列表
            var resultLabel = new Label
            {
                Text = "搜索结果:",
                Location = new Point(10, 120),
                Size = new Size(100, 30),
                Font = new Font("微软雅黑", 9, FontStyle.Bold)
            };
            panelSearch.Controls.Add(resultLabel);

            var searchResultListBox = new ListBox
            {
                Name = "searchResultListBox",
                Location = new Point(10, 155),
                Size = new Size(560, 200),
                SelectionMode = SelectionMode.One
            };
            searchResultListBox.DoubleClick += SearchResultListBox_DoubleClick;
            panelSearch.Controls.Add(searchResultListBox);

            // 操作按钮
            var viewButton = new Button
            {
                Text = "查看详情",
                Location = new Point(10, 355),
                Size = new Size(100, 35),
                BackColor = Color.Green,
                ForeColor = Color.White
            };
            viewButton.Click += SearchViewDetailButton_Click;
            panelSearch.Controls.Add(viewButton);

            var openFileButton = new Button
            {
                Text = "打开原稿",
                Location = new Point(120, 355),
                Size = new Size(100, 35),
                BackColor = Color.Orange,
                ForeColor = Color.White
            };
            openFileButton.Click += OpenOriginalFileButton_Click;
            panelSearch.Controls.Add(openFileButton);
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

            string keyword = keywordTextBox?.Text ?? "";
            string searchField = searchFieldCombo?.SelectedItem?.ToString() ?? "全部";
            DateTime? startDate = startDateCheck?.Checked == true ? startDatePicker?.Value : null;
            DateTime? endDate = endDateCheck?.Checked == true ? endDatePicker?.Value : null;

            var results = _managementService.SearchResumes(keyword, startDate, endDate, searchField);
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

        private Dictionary<string, List<Resume>> _currentDuplicates = new Dictionary<string, List<Resume>>();

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
