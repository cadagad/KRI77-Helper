using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KRI77_Helper.Models
{
    public  class MobileDevice
    {
        public string? DeviceID { get; set; }
        public string? Devicename { get; set; }
        public string? Enrollmentdate { get; set; }
        public string? Lastcheckin { get; set; }
        public string? AzureADDeviceID { get; set; }
        public string? OSversion { get; set; }
        public string? AzureADregistered { get; set; }
        public string? EASactivationID { get; set; }
        public required string Serialnumber { get; set; }
        public string? Manufacturer { get; set; }
        public string? Model { get; set; }
        public string? EASactivated { get; set; }
        public string? IMEI { get; set; }
        public string? LastEASsynctime { get; set; }
        public string? EASreason { get; set; }
        public string? EASstatus { get; set; }
        public string? Compliancegraceperiodexpiration { get; set; }
        public string? Securitypatchlevel { get; set; }
        public string? WifiMAC { get; set; }
        public string? MEID { get; set; }
        public string? Subscribercarrier { get; set; }
        public string? Totalstorage { get; set; }
        public string? Freestorage { get; set; }
        public string? Managementname { get; set; }
        public string? Category { get; set; }
        public string? UserId { get; set; }
        public string? PrimaryuserUPN { get; set; }
        public string? Primaryuseremailaddress { get; set; }
        public string? Primaryuserdisplayname { get; set; }
        public string? WiFiIPv4Address { get; set; }
        public string? WiFiSubnetID { get; set; }
        public string? Compliance { get; set; }
        public string? Managedby { get; set; }
        public string? Ownership { get; set; }
        public string? Devicestate { get; set; }
        public string? Intuneregistered { get; set; }
        public string? Supervised { get; set; }
        public string? Encrypted { get; set; }
        public string? OS { get; set; }
        public string? SkuFamily { get; set; }
        public string? JoinType { get; set; }
        public string? Phonenumber { get; set; }
        public string? Jailbroken { get; set; }
        public string? ICCID { get; set; }
        public string? EthernetMAC { get; set; }
        public string? CellularTechnology { get; set; }
        public string? ProcessorArchitecture { get; set; }
        public string? EID { get; set; }
        public string? SystemManagementBIOSVersion { get; set; }
        public string? TPMManufacturerId { get; set; }
        public string? TPMManufacturerVersion { get; set; }
        public string? ProductName { get; set; }
        public string? Managementcertificateexpirationdate { get; set; }

    }
}
