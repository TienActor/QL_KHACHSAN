//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace QLKS.Models
{
    using System;
    using System.Collections.Generic;
    
    public partial class tblTinNhan
    {
        public int id { get; set; }
        public string ma_kh { get; set; }
        public string ho_ten { get; set; }
        public string mail { get; set; }
        public string noi_dung { get; set; }
        public Nullable<System.DateTime> ngay_gui { get; set; }
        public int danh_gia { get; set; }
    
        public virtual tblKhachHang tblKhachHang { get; set; }
    }
}
