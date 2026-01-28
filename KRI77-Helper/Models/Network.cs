using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KRI77_Helper.Models
{
    public class Network
    {
        public string? Country { get; set; }
        public required string HostName { get; set; }
        public string? SerialNumber { get; set; }
        public string? ModelName { get; set; }
        public string? CountryLocation { get; set; }
    }
}
