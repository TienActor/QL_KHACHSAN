﻿//------------------------------------------------------------------------------
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
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    [Table("tblKhachHang")]
    public partial class tblKhachHang
    {
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2214:DoNotCallOverridableMethodsInConstructors")]
        public tblKhachHang()
        {
            this.tblPhieuDatPhongs = new HashSet<tblPhieuDatPhong>();
            this.tblTinNhans = new HashSet<tblTinNhan>();
        }

        [Key]
        public string ma_kh { get; set; }
        [Required(ErrorMessage = "Vui lòng nhập đầy đủ thông tin.")]
        [MinLength(6, ErrorMessage = "Mật khẩu phải chứa ít nhất 6 kí tự.")]
        public string mat_khau { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập đầy đủ thông tin.")]
        public string ho_ten { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập đầy đủ thông tin.")]
        [MinLength(9, ErrorMessage = "Số CMT phải có ít nhất 9 chữ số.")]
        public string cmt { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập đầy đủ thông tin.")]
        [RegularExpression(@"^0[0-9]{8,}$", ErrorMessage = "Số điện thoại phải bắt đầu bằng số 0 và có ít nhất 9 chữ số.")]
        public string sdt { get; set; }

        [Required(ErrorMessage = "Vui lòng nhập đầy đủ thông tin.")]

        [RegularExpression(@".*@gmail\.com$", ErrorMessage = "Email phải kết thúc bằng '@gmail.com'.")]
        public string mail { get; set; }
        public Nullable<int> diem { get; set; }


        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<tblPhieuDatPhong> tblPhieuDatPhongs { get; set; }
        [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public virtual ICollection<tblTinNhan> tblTinNhans { get; set; }
    }
}
