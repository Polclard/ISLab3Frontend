using System.ComponentModel.DataAnnotations;

namespace AdminApplication.Models
{
    public class Concert
    {
        [Required]
        public string ConcertName { get; set; }
        [Required]
        public string ConcertDescription { get; set; }
        [Required]
        public string ConcertImage { get; set; }
        [Required]
        public double Rating { get; set; }
    }
}
