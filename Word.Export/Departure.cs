using System;

namespace Word.Export
{
    public class Departure
    {
        /// <summary>
        /// 编号
        /// </summary>
        public int Order { get; set; }

        /// <summary>
        /// 姓名
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 身份证号
        /// </summary>
        public string IdNum { get; set; }

        /// <summary>
        /// 离场时间
        /// </summary>
        public DateTime DepartureDate { get; set; }

        /// <summary>
        /// 务工时间
        /// </summary>
        public double WorkDays { get; set; }

        /// <summary>
        /// 工种（或岗位）
        /// </summary>
        public string Job { get; set; }

        /// <summary>
        /// 工资结算、支付情况(是否已结算)
        /// </summary>
        public bool IsSettled { get; set; }

        /// <summary>
        /// 是否有离场承诺书
        /// </summary>
        public bool IsCommited { get; set; }
    }
}
