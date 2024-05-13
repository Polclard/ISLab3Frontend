using System.ComponentModel.DataAnnotations;

namespace AdminApplication.Models
{
    public class Ticket
    {
        public Guid ConcertId { get; set; }
        public Concert? Concert { get; set; }

        [Required]
        public double Price { get; set; }

        [Required]
        public double Rating { get; set; }
    }
}
