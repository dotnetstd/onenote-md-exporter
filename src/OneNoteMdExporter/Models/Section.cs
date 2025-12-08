using alxnbl.OneNoteMdExporter.Infrastructure;
using static System.Collections.Specialized.BitVector32;

namespace alxnbl.OneNoteMdExporter.Models
{
    public class Section(Node parent) : Node(parent)
    {
        public bool IsSectionGroup { get; set; }
    }
}
