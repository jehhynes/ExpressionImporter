using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExpressionImporter
{
    public enum ImportType
    {
        [Description("Update Existing Records")]
        Update = 0,
        [Description("Create New Records")]
        Create = 1,
        [Description("Full Import (Create and Update)")]
        Full = 3
    }
}
