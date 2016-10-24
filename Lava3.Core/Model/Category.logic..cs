using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lava3.Core.Model
{
    public partial class Category 
    {


        /// <summary>
        /// Build the regex key from the description.
        /// </summary>
        internal void BuildRegex()
        {
            // Set for default
            RegEx = new DataTypes.ColumnString()
            {
                Value = $"^({Description.Value})$"
                        .Replace("^(*", "(")
                         .Replace("*)$", ")")
            };
           
        }
        public override string ToString()
        {
            string retval =  $"{Description.Value} - {AccountingCategory.Value}";
            if (IsDuplicateDescription)
                retval += "  - IsDuplicateDescription";
            if (IsDuplicateNotes)
                retval += $" - IsDuplicateNotes";
            return retval;
        }
    }
}
