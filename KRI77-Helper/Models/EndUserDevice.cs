using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KRI77_Helper.Models
{
    public class EndUserDevice
    {
        public string? ChassisType { get; set; }
        public string? CpuName { get; set; }
        public string? ComputerId { get; set; }
        public string? ComputerName { get; set; }
        public required string SerialNumber { get; set; }
        public string? OsPlatform { get; set; }
        public string? OperatingSystem { get; set; }
        public string? ServicePack { get; set; }
        public string? Manufacturer { get; set; }
        public string? Model { get; set; }
        public string? IPAddress { get; set; }
        public string? UserName { get; set; }
        public string? CreatedDate { get; set; }
        public string? UpdatedDate { get; set; }
        public string? LastSeen { get; set; }
        public string? IsVirtual { get; set; }
        public string? OSVersion { get; set; }
        public string? Count { get; set; }
    }
}
