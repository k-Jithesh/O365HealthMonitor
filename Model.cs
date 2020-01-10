using System;
using System.Collections.Generic;
using System.Text;

namespace O365HealthMonitor
{
    class Token
    {
        public string token_type { get; set; }
        public int expires_in { get; set; }
        public int ext_expires_in { get; set; }
        public int expires_on { get; set; }
        public int not_before { get; set; }
        public string resource { get; set; }
        public string access_token { get; set; }
    }

    public class FeatureStatus
    {
        public string FeatureDisplayName { get; set; }
        public string FeatureName { get; set; }
        public string FeatureServiceStatus { get; set; }
        public string FeatureServiceStatusDisplayName { get; set; }
    }

    public class FeatureValue
    {
        public List<FeatureStatus> FeatureStatus { get; set; }
        public Guid id { get; set; }
        public string Id { get; set; }
        public List<object> IncidentIds { get; set; }
        public string Status { get; set; }
        public string StatusDisplayName { get; set; }
        public DateTime StatusTime { get; set; }
        public string Workload { get; set; }
        public string WorkloadDisplayName { get; set; }
        public long BatchNumber { get; set; }
    }

    public class CurrentStatus
    {
        public List<FeatureValue> value { get; set; }
    }


    public class Message
    {
        public string MessageText { get; set; }
        public DateTime PublishedTime { get; set; }
    }

    public class Messages
    {
        public List<MessageValue> value { get; set; }
    }

    public class MessageValue
    {
        public List<object> AffectedWorkloadDisplayNames { get; set; }
        public List<object> AffectedWorkloadNames { get; set; }
        public string Status { get; set; }
        public string Workload { get; set; }
        public string WorkloadDisplayName { get; set; }
        public string ActionType { get; set; }
        public int AffectedTenantCount { get; set; }
        public object AffectedUserCount { get; set; }
        public string Classification { get; set; }
        public DateTime? EndTime { get; set; }
        public string Feature { get; set; }
        public string FeatureDisplayName { get; set; }
        public string UserFunctionalImpact { get; set; }
        public string Id { get; set; }
        public Guid id { get; set; }
        public string ImpactDescription { get; set; }
        public DateTime LastUpdatedTime { get; set; }
        public string MessageType { get; set; }
        public List<Message> Messages { get; set; }
        public string PostIncidentDocumentUrl { get; set; }
        public string Severity { get; set; }
        public DateTime StartTime { get; set; }
        public string Title { get; set; }
        public DateTime? ActionRequiredByDate { get; set; }
        public int? AnnouncementId { get; set; }
        public string Category { get; set; }
        public List<object> MessageTagNames { get; set; }
        public string ExternalLink { get; set; }
        public bool? IsDismissed { get; set; }
        public bool? IsRead { get; set; }
        public bool? IsMajorChange { get; set; }
        public object PreviewDuration { get; set; }
        public string AppliesTo { get; set; }
        public DateTime? MilestoneDate { get; set; }
        public string Milestone { get; set; }
        public string BlogLink { get; set; }
        public string HelpLink { get; set; }
        public object FlightName { get; set; }
        public string FeatureName { get; set; }
        public long BatchNumber { get; set; }
    }
}
