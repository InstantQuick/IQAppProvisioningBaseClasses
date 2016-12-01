using Microsoft.SharePoint.Client;

namespace IQAppProvisioningBaseClasses.Provisioning
{
    public class SiteAuditSettings
    {
        public AuditMaskType AuditMaskType { get; set; }
        public int AuditLogTrimmingRetention { get; set; }
        public bool TrimAuditLog { get; set; }
    }
}
