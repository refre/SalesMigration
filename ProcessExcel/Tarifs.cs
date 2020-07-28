using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ista.Migration.Excel
{
    /// <summary>
    /// This class is an object to any tarifs value in the xml file.
    /// </summary>
    public class Tarifs
    {
        /// <summary>
        /// Gets or sets the Id.
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// Gets or sets the name.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets the Releve value.
        /// </summary>
        public float Releve { get; set; }

        /// <summary>
        /// Gets or sets the Value.
        /// </summary>
        public float Value { get; set; }

        /// <summary>
        /// Gets or sets the Vente value.
        /// </summary>
        public float Vente { get; set; }

        /// <summary>
        /// Gets or sets the LocationMax15 value.
        /// </summary>
        public float LocationMax15 { get; set; }
    }
}
