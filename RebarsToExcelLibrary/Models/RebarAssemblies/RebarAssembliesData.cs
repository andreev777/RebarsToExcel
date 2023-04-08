using System.Collections.Generic;
using System.Linq;

namespace RebarsToExcel.Models
{
    public static class RebarAssembliesData
    {
        private static List<RebarAssembly> _data = new List<RebarAssembly>();

        public static void AddRebarAssembly(RebarAssembly rebarAssembly)
        {
            foreach (var existedAssembly in _data)
            {
                if (existedAssembly.Type == rebarAssembly.Type
                    && existedAssembly.Mark == rebarAssembly.Mark
                    && existedAssembly.Mass == rebarAssembly.Mass
                    && existedAssembly.ConstructionType == rebarAssembly.ConstructionType
                    && existedAssembly.ConstructionMark == rebarAssembly.ConstructionMark
                    && existedAssembly.Level.Name == rebarAssembly.Level.Name
                    && existedAssembly.ConstructionCount == rebarAssembly.ConstructionCount
                    && existedAssembly.TypicalFloor == rebarAssembly.TypicalFloor
                    && existedAssembly.TypicalFloorCount == rebarAssembly.TypicalFloorCount
                    && existedAssembly.Section == rebarAssembly.Section)
                {
                    existedAssembly.Count++;
                    existedAssembly.AddId(rebarAssembly.Id);
                    return;
                }
            }

            rebarAssembly.AddId(rebarAssembly.Id);
            rebarAssembly.Ids.Distinct();
            _data.Add(rebarAssembly);
        }

        public static List<RebarAssembly> GetData()
        {
            return _data.OrderBy(x => x.Section)
                .ThenBy(x => x.Level.Elevation)
                .ThenBy(x => x.ConstructionType)
                .ThenBy(x => x.ConstructionMark)
                .ThenBy(x => x.Type)
                .ThenBy(x => x.Mark)
                .ToList();
        }

        public static List<RebarLevel> GetLevels()
        {
            return _data.Select(rebarAssembly => rebarAssembly.Level)
                .Distinct(new LevelComparer())
                .OrderBy(level => level.Elevation)
                .ToList();
        }

        public static List<string> GetSections()
        {
            return _data.Select(rebarAssembly => rebarAssembly.Section)
                .Distinct()
                .OrderBy(section => section)
                .ToList();
        }

        public static List<string> GetConstructionTypes()
        {
            return _data.Select(rebarAssembly => rebarAssembly.ConstructionType)
                .Distinct()
                .OrderBy(type => type)
                .ToList();
        }

        public static void AnalyzeDataByConstructionCount()
        {
            var constructionCountRebarAssemblies = new List<RebarAssembly>();

            foreach (var rebarAssembly in _data)
            {
                if (rebarAssembly.ConstructionCount > 1)
                {
                    for (int i = 1; i < rebarAssembly.ConstructionCount; i++)
                    {
                        var constructionCountRebarAssembly = new RebarAssembly(rebarAssembly.Type, rebarAssembly.Mark, rebarAssembly.GroupModel, rebarAssembly.Mass)
                        {
                            Id = rebarAssembly.Id,
                            Ids = rebarAssembly.Ids,
                            IdsAsString = rebarAssembly.IdsAsString,
                            Definition = rebarAssembly.Definition,
                            Count = rebarAssembly.Count,
                            Level = rebarAssembly.Level,
                            Section = rebarAssembly.Section,
                            ConstructionType = rebarAssembly.ConstructionType,
                            ConstructionMark = rebarAssembly.ConstructionMark,
                            ConstructionCount = rebarAssembly.ConstructionCount,
                            TypicalFloor = rebarAssembly.TypicalFloor,
                            TypicalFloorCount = rebarAssembly.TypicalFloorCount,
                            ElementType = RebarElementType.Virtual,
                            Rebars = rebarAssembly.Rebars,
                        };

                        constructionCountRebarAssemblies.Add(constructionCountRebarAssembly);
                    }
                }
            }

            foreach (var rebarAssembly in constructionCountRebarAssemblies)
            {
                AddRebarAssembly(rebarAssembly);
            }
        }

        public static void AnalyzeDataByTypicalFloorCount(List<TypicalFloor> typicalFloors)
        {
            var typicalFloorRebarAssemblies = new List<RebarAssembly>();

            foreach (var rebarAssembly in _data)
            {
                if (rebarAssembly.TypicalFloorCount > 1)
                {
                    var typicalRebarLevels = GetTypicalRebarLevels(rebarAssembly, typicalFloors);

                    if (typicalRebarLevels == null)
                    {
                        return;
                    }

                    foreach (var typicalRebarLevel in typicalRebarLevels)
                    {
                        var typicalFloorRebarAssembly = new RebarAssembly(rebarAssembly.Type, rebarAssembly.Mark, rebarAssembly.GroupModel, rebarAssembly.Mass)
                        {
                            Id = rebarAssembly.Id,
                            Ids = rebarAssembly.Ids,
                            IdsAsString = rebarAssembly.IdsAsString,
                            Definition = rebarAssembly.Definition,
                            Count = rebarAssembly.Count,
                            Level = typicalRebarLevel,
                            Section = rebarAssembly.Section,
                            ConstructionType = rebarAssembly.ConstructionType,
                            ConstructionMark = rebarAssembly.ConstructionMark,
                            ConstructionCount = rebarAssembly.ConstructionCount,
                            TypicalFloor = rebarAssembly.TypicalFloor,
                            TypicalFloorCount = rebarAssembly.TypicalFloorCount,
                            ElementType = RebarElementType.Virtual,
                            Rebars = rebarAssembly.Rebars,
                        };

                        if (typicalFloorRebarAssembly.Level.Name != rebarAssembly.Level.Name)
                        {
                            typicalFloorRebarAssemblies.Add(typicalFloorRebarAssembly);
                        }
                    }
                }
            }

            foreach (var rebarAssembly in typicalFloorRebarAssemblies)
            {
                AddRebarAssembly(rebarAssembly);
            }
        }

        private static List<RebarLevel> GetTypicalRebarLevels(RebarAssembly rebarAssembly, List<TypicalFloor> typicalFloors)
        {
            return typicalFloors.Where(typicalFloor => typicalFloor.ConstructionTypeEnum == rebarAssembly.ConstructionTypeEnum)
                .Where(typicalFloor => typicalFloor.Floor == rebarAssembly.TypicalFloor)
                .SelectMany(typicalFloor => typicalFloor.Levels)
                .ToList();
        }
    }
}
