using System.ComponentModel.DataAnnotations;

namespace RDLCReportByAsp.net.Models
{
    public class Book
    {
        [Required]
        public int Id { get; set; }
        [Required]
        public string Auther { get; set; }
        [Required]
        public string Title { get; set; }
        [Required]
        public string Price { get; set; }
    }
}