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
    
    public partial class hpaidrec
    {
        public System.Guid paid_id { get; set; }
        public string if_forsettle { get; set; }
        public string settlecode { get; set; }
        public string discount_type { get; set; }
        public string paid_title { get; set; }
        public Nullable<int> paid_qty { get; set; }
        public string paid_way { get; set; }
        public Nullable<decimal> paid_checkfee { get; set; }
        public Nullable<decimal> paid_plusfee { get; set; }
        public Nullable<decimal> paid_undofee { get; set; }
        public Nullable<decimal> paid_sumfee { get; set; }
        public Nullable<decimal> paid_round { get; set; }
        public Nullable<decimal> paid_real { get; set; }
        public string paid_invoice { get; set; }
        public string paid_refcode { get; set; }
        public string paid_contract { get; set; }
        public string paid_remark { get; set; }
        public string paid_oper { get; set; }
        public Nullable<System.DateTime> paid_date { get; set; }
        public string if_confirm { get; set; }
        public string confirm_oper { get; set; }
        public Nullable<System.DateTime> confirm_date { get; set; }
        public string discount_name { get; set; }
        public Nullable<decimal> paid_discountRate { get; set; }
        public Nullable<decimal> paid_payment { get; set; }
    }
}