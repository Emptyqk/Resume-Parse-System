using System;
using System.Collections.Generic;

namespace 页面.Models
{
    // 简历数据模型
    public class Resume
    {
        // 唯一标识符
        public string Id { get; set; } = Guid.NewGuid().ToString();

        // 姓名
        public string Name { get; set; } = "";

        // 性别
        public string Gender { get; set; } = "";

        // 出生日期
        public DateTime? BirthDate { get; set; }

        // 家庭住址
        public string Address { get; set; } = "";

        // 最高学历
        public string HighestEducationSchool { get; set; } = "";

        // 最高学历的专业
        public string HighestEducationMajor { get; set; } = "";

        // 第一学历（本科）的毕业院校
        public string FirstEducationSchool { get; set; } = "";

        // 第一学历的专业
        public string FirstEducationMajor { get; set; } = "";

        // 手机号码
        public string Phone { get; set; } = "";

        // 邮箱
        public string Email { get; set; } = "";


        // 身份证号
        public string IdCard { get; set; } = "";

        // 工作经历列表
        public List<WorkExperience> WorkExperiences { get; set; } = new List<WorkExperience>();

        // 技能列表
        public List<string> Skills { get; set; } = new List<string>();

        // 原始文件路径
        public string OriginalFilePath { get; set; } = "";

        // 文件名
        public string FileName { get; set; } = "";

        // 导入时间
        public DateTime ImportTime { get; set; } = DateTime.Now;

        // 所属目录
        public string Directory { get; set; } = "默认目录";
    }

    // 工作经历模型
    public class WorkExperience
    {
        // 单位名称
        public string Company { get; set; } = "";

        // 岗位
        public string Position { get; set; } = "";

        // 项目
        public string Project { get; set; } = "";

        // 职务
        public string Title { get; set; } = "";

        // 职责描述
        public string Responsibilities { get; set; } = "";

        // 开始时间
        public DateTime? StartDate { get; set; }

        // 结束时间
        public DateTime? EndDate { get; set; }

        // 是否当前工作
        public bool IsCurrentJob { get; set; } = false;
    }

    // 简历目录模型
    public class ResumeDirectory
    {
        // 目录名称
        public string Name { get; set; } = "";

        // 创建时间
        public DateTime CreatedTime { get; set; } = DateTime.Now;
    }

}
