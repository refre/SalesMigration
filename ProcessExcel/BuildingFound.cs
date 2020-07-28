using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ista.Migration.Excel
{
    /// <summary>
    /// Records data whether value if some data already match.
    /// </summary>
    public class BuildingFound
    {
        /// <summary>
        /// Gets or sets building.
        /// </summary>
        public string Building { get; set; }
        /// <summary>
        /// Gets or sets the offer number.
        /// </summary>
        public string OffreNumber { get; set; }
        /// <summary>
        /// Gets or set the datetime when it has been sent.
        /// </summary>
        public string DateSent { get; set; }
    }
}
