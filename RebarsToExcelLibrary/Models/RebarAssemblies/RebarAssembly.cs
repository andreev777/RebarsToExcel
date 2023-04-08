using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

namespace RebarsToExcel.Models
{
    public class RebarAssembly
    {
        public ElementId Id { get; set; }
        public List<ElementId> Ids { get; set; } = new List<ElementId>();
        public string IdsAsString { get; set; }
        public string Type { get; set; }
        public string Mark { get; set; }
        public string Definition { get; set; }
        public string GroupModel { get; set; }
        public int Count { get; set; }
        public double Mass { get; set; }
        public RebarLevel Level { get; set; }
        public string Section { get; set; }
        public string ConstructionType { get; set; }
        public ConstructionType ConstructionTypeEnum { get; set; } = Models.ConstructionType.Unknown;
        public string ConstructionMark { get; set; }
        public int ConstructionCount { get; set; } = 0;
        public int TypicalFloor { get; set; } = 0;
        public int TypicalFloorCount { get; set; } = 0;
        public RebarElementType ElementType { get; set; } = RebarElementType.Real;

        public List<Rebar> Rebars = new List<Rebar>();

        public RebarAssembly(string type, string mark, string groupModel, double mass)
        {
            Type = type;
            Mark = mark;
            GroupModel = groupModel;
            Mass = mass;
            Count = 1;
        }

        public void AddRebar(Rebar rebar)
        {
            if (Rebars.Count == 0)
            {
                Rebars.Add(rebar);
                return;
            }

            foreach (var existedRebar in Rebars)
            {
                if (rebar.Class == existedRebar.Class && rebar.Diameter == existedRebar.Diameter && rebar.Length == existedRebar.Length)
                {
                    existedRebar.Count += rebar.Count;
                    return;
                }
            }

            Rebars.Add(rebar);
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

        public string GetRebarsInfoString()
        {
            var rebarInfoStringLines = new List<string>();

            foreach (var rebar in Rebars)
            {
                var rebarInfoStringLine = rebar.GetClassDiameterCountInfoString();
                rebarInfoStringLines.Add(rebarInfoStringLine);
            }

            return string.Join("\n", rebarInfoStringLines);
        }
    }
}