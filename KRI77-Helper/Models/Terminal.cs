using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KRI77_Helper.Models
{
    public class Terminal
    {
        public string? DeviceName { get; set; }
        public string? Model { get; set; }
        public string? OSVersion { get; set; }
        public required string SerialNumber { get; set; }
        public string? Location { get; set; }
        public string? TimeZone { get; set; }
        public string? Locale { get; set; }
        public string? Group { get; set; }
    }
}
