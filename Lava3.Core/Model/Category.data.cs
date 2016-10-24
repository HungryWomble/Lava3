using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Lava3.Core.DataTypes;

namespace Lava3.Core.Model
{
    public partial class Category
    {
        public ColumnString AccountingCategory { get; set; }
        public ColumnString Notes { get; set; }
        public Uri NotesHyperLink { get; set; }

        private ColumnString _Description;
        public ColumnString Description
        {
            get
            {
                return _Description;
            }
            set
            {
                _Description = value;
                BuildRegex();
            }
        }
        public ColumnString RegEx { get; set; }
        public virtual bool IsDuplicateDescription { get; set; }
        public virtual bool IsDuplicateNotes { get; set; }
    }
}
