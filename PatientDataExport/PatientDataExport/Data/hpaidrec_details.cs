//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace PatientDataExport.Data
{
    using System;
    using System.Collections.Generic;
    
    public partial class hpaidrec_details
    {
        public System.Guid paid_id { get; set; }
        public string unioncode { get; set; }
        public string unionname { get; set; }
        public Nullable<decimal> paid_unionfee { get; set; }
        public Nullable<decimal> paid_discountRate { get; set; }
        public Nullable<int> paid_checkqty { get; set; }
        public Nullable<decimal> paid_checkfee { get; set; }
        public Nullable<int> paid_plusqty { get; set; }
        public Nullable<decimal> paid_plusfee { get; set; }
        public Nullable<int> paid_undoqty { get; set; }
        public Nullable<decimal> paid_undofee { get; set; }
    }
}
