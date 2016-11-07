using System.Collections.Generic;

namespace Lava3.Core.DataTypes
{
    public interface IColumDataType
    {
        List<string> Errors { get; set; }
        int ColumnNumber { get; set; }

    }
}
