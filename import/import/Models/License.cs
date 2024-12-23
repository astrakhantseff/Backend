using System.Text.RegularExpressions;

namespace FileImport.Models
{
    internal class License
    {
        internal string LicenseNumber { get; set; }
        internal string Series { get; set; }
        internal string Number { get; set; }
        internal string Type { get; set; }
        internal string Status { get; set; }
        internal string Condition { get; set; }
        internal string IntendedUse { get; set; }
        internal string RegistrationDate { get; set; }
        internal string ExpirationDate { get; set; }
        internal string CancellationDate { get; set; }
        internal string LicensingAuthority { get; set; }
        internal string AreaAccordingToDocument { get; set; }
        internal string DepthLimitationOfMiningWithdrawal { get; set; }
        internal string Permanent { get; set; }
        internal string JustificationForObtaining { get; set; }

        internal void Parse()
        {
            const string pattern = @"(\D{3})(\d{5})(\D{2})"; // Регулярное выражение для извлечения частей строки

            var match = Regex.Match(LicenseNumber, pattern);

            if (match.Success)
            {
                Series = match.Groups[1].Value;
                Number = match.Groups[2].Value;
                Type = match.Groups[3].Value;
            }
        }
    }
}