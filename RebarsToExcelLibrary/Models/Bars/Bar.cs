using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

namespace RebarsToExcel.Models
{
    public class Bar
    {
        public ElementId Id { get; set; }
        public List<ElementId> Ids { get; set; } = new List<ElementId>();
        public string IdsAsString { get; set; }
        public string Position { get; set; }
        public string Class { get; set; }
        public double Diameter { get; set; } = 0;
        public double Length { get; set; } = 0;
        public int Count { get; set; } = 0;
        public int CountType { get; set; } = 0;
        public double Mass { get; set; } = 0;
        public string Shape { get; set; }
        public RebarLevel Level { get; set; }
        public string Section { get; set; }
        public string ConstructionType { get; set; }
        public ConstructionType ConstructionTypeEnum { get; set; } = Models.ConstructionType.Unknown;
        public string ConstructionMark { get; set; }
        public int ConstructionCount { get; set; } = 0;
        public int TypicalFloor { get; set; } = 0;
        public int TypicalFloorCount { get; set; } = 0;
        public RebarElementType ElementType { get; set; } = RebarElementType.Real;
        public string DiameterClassInfo { get; set; }
        public string CountTypeInfo { get; set; }

        public Bar(string rebarClass, double diameter, double mass, string shape)
        {
            Class = rebarClass;
            Diameter = diameter;
            Mass = mass;
            Shape = shape;
            DiameterClassInfo = $"⌀{Diameter} {Class}";
        }

        public void AddId(ElementId id)
        {
            if (id != null)
            {
                Ids.Add(id);
                var idsAsIntegerCollection = Ids.Select(x => x.IntegerValue).Distinct().ToList();
                IdsAsString = string.Join(", ", idsAsIntegerCollection);
            }
        }
    }
}