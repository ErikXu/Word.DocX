using System.Collections.Generic;

namespace Word.Export
{
    public class Info
    {
        /// <summary>
        /// 基本信息
        /// </summary>
        public Basic Basic { get; set; }
        
        /// <summary>
        /// 进场人员信息
        /// </summary>
        public List<Enter> Enters { get; set; }

        /// <summary>
        /// 离场人员信息
        /// </summary>
        public List<Departure> Departures { get; set; }
    }
}
