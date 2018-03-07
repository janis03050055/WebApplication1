using System;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace WebApplication1.Models
{
    public class Calendar
    {
        [Key]
        public int ID { get; set; }

        [Required, DisplayName("日期"), DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}", ApplyFormatInEditMode = true)]
        public DateTime date { get; set; }

        [DisplayName("日期名稱")]
        public string name { get; set; }

        [Required, DisplayName("是否放假")]
        public string isHoliday { get; set; }

        [Required, DisplayName("日期類別")]
        public string holidayCategory { get; set; }

        [DisplayName("日期描述")]
        public string description { get; set; }
    }
}
