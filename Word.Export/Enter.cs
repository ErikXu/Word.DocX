using System;

namespace Word.Export
{
    public class Enter
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
        /// 进场时间
        /// </summary>
        public DateTime EnterDate { get; set; }

        /// <summary>
        /// 工种（或岗位）
        /// </summary>
        public string Job { get; set; }

        /// <summary>
        /// 是否在市建委备案
        /// </summary>
        public bool IsRecord { get; set; }

        /// <summary>
        /// 是否签订《劳动合同》
        /// </summary>
        public bool HasContract { get; set; }
    }
}
