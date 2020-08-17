using System;
using System.Collections.Generic;
using System.Text;

namespace Utils.Model
{
    class CountryStructure
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string CountryShortCode { get; set; }
        public string PostalCodeValidationRule { get; set; }
        public string NeedsState { get; set; }
        public string PrefixRegex { get; set; }
    }
}
