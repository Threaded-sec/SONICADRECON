using System;
using System.Collections.Generic;

namespace ADAUDIT.Models
{
    public class Group
    {
        public string Name { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;
        public string DistinguishedName { get; set; } = string.Empty;
        public string GroupType { get; set; } = string.Empty;
        public string Scope { get; set; } = string.Empty;
        public int MemberCount { get; set; }
        public DateTime? Created { get; set; }
        public DateTime? Modified { get; set; }
        public List<string> Members { get; set; } = new List<string>();
    }
} 