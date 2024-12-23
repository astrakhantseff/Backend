using System.Collections.Generic;

namespace FileImport.Models
{
    internal class Field
    {
        internal string Index { get; set; }
        internal string FieldName { get; set; }
        internal string DevelopmentStage { get; set; }
        internal string FieldType { get; set; }
        internal string RegionOfRussia { get; set; }
        internal string Location { get; set; }
        internal string DiscoveryYear { get; set; }
        internal string DevelopmentStartYear { get; set; }
        internal string Area { get; set; }
        internal string ConservationYear { get; set; }
        internal string DecommissioningYear { get; set; }
        internal List<License> Licenses { get; set; } = new List<License>();

        // Добавленное свойство
        internal License License { get; set; }
        internal string ProtocolNumber { get; set; }
    }
}
