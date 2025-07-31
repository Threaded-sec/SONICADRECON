using System;

namespace ADAUDIT.Models
{
    public class GroupPolicy
    {
        public string Name { get; set; } = string.Empty;
        public string Id { get; set; } = string.Empty;
        public string Status { get; set; } = string.Empty;
        public string DistinguishedName { get; set; } = string.Empty;
        public DateTime? Created { get; set; }
        public DateTime? Modified { get; set; }
        public string Owner { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public bool IsEnabled { get; set; }
        public string Version { get; set; } = string.Empty;
    }
} 