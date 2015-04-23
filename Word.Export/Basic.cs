using System;

namespace Word.Export
{
    public class Basic
    {
        /// <summary>
        /// 公司
        /// </summary>
        public string Company { get; set; }

        /// <summary>
        /// 项目
        /// </summary>
        public string Project { get; set; }

        /// <summary>
        /// 班组
        /// </summary>
        public string Team { get; set; }

        /// <summary>
        /// 开始日期
        /// </summary>
        public DateTime StartDate { get; set; }

        /// <summary>
        /// 结束日期
        /// </summary>
        public DateTime EndDate { get; set; }

        /// <summary>
        /// 进场总数
        /// </summary>
        public int EnterCount { get; set; }

        /// <summary>
        /// 离场总数
        /// </summary>
        public int DepartureCount { get; set; }

        /// <summary>
        /// 现场总数
        /// </summary>
        public int CurrentCount { get; set; }
    }
}
