//------------------------------------------------------------------------------
// <auto-generated>
//     此代码已从模板生成。
//
//     手动更改此文件可能导致应用程序出现意外的行为。
//     如果重新生成代码，将覆盖对此文件的手动更改。
// </auto-generated>
//------------------------------------------------------------------------------

namespace PatientDataExport.Data
{
    using System;
    using System.Collections.Generic;
    
    public partial class hstatdiags
    {
        public string statdiag_code { get; set; }
        public string p_statdiag_code { get; set; }
        public string statdiag_name { get; set; }
        public string statdiag_advice { get; set; }
        public bool invalidFlag { get; set; }
        public bool isTopFlag { get; set; }
        public bool isBottomFlag { get; set; }
        public int statdiag_seq { get; set; }
        public string statview_code { get; set; }
        public string vdiag_code { get; set; }
    }
}