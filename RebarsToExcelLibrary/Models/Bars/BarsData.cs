using System;
using System.Collections.Generic;
using System.Linq;

namespace RebarsToExcel.Models
{
    public static class BarsData
    {
        private static List<Bar> _data = new List<Bar>();

        public static void AddBar(Bar bar)
        {
            if (bar.CountType == 1)
            {
                foreach (var existedBar in _data)
                {
                    if (existedBar.Position == bar.Position
                        && existedBar.Class == bar.Class
                        && existedBar.Diameter == bar.Diameter
                        && existedBar.Mass == bar.Mass
                        && existedBar.Length == bar.Length
                        && existedBar.CountType == bar.CountType
                        && existedBar.Shape == bar.Shape
                        && existedBar.ConstructionType == bar.ConstructionType
                        && existedBar.ConstructionMark == bar.ConstructionMark
                        && existedBar.Level.Name == bar.Level.Name
                        && existedBar.ConstructionCount == bar.ConstructionCount
                        && existedBar.TypicalFloor == bar.TypicalFloor
                        && existedBar.TypicalFloorCount == bar.TypicalFloorCount
                        && existedBar.Section == bar.Section)
                    {
                        existedBar.Count += bar.Count;
                        existedBar.AddId(bar.Id);
                        return;
                    }
                }

                bar.AddId(bar.Id);
                bar.Ids.Distinct();
                _data.Add(bar);
            }

            else if (bar.CountType == 2)
            {
                foreach (var existedBar in _data)
                {
                    if (existedBar.Position == bar.Position
                        && existedBar.Class == bar.Class
                        && existedBar.Diameter == bar.Diameter
                        && existedBar.CountType == bar.CountType
                        && existedBar.Shape == bar.Shape
                        && existedBar.ConstructionType == bar.ConstructionType
                        && existedBar.ConstructionMark == bar.ConstructionMark
                        && existedBar.Level.Name == bar.Level.Name
                        && existedBar.ConstructionCount == bar.ConstructionCount
                        && existedBar.TypicalFloor == bar.TypicalFloor
                        && existedBar.TypicalFloorCount == bar.TypicalFloorCount
                        && existedBar.Section == bar.Section)
                    {
                        existedBar.Count += bar.Count;
                        existedBar.Length += bar.Length;
                        existedBar.Mass += bar.Mass;
                        existedBar.AddId(bar.Id);
                        return;
                    }
                }

                bar.AddId(bar.Id);
                bar.Ids.Distinct();
                _data.Add(bar);
            }
        }

        public static List<Bar> GetData()
        {
            return _data.OrderBy(x => x.Section)
                .ThenBy(x => x.Level.Elevation)
                .ThenBy(x => x.ConstructionType)
                .ThenBy(x => x.ConstructionMark)
                .ThenBy(x => x.Position)
                .ToList();
        }

        public static List<RebarLevel> GetLevels()
        {
            return _data.Select(bar => bar.Level)
                .Distinct(new LevelComparer())
                .OrderBy(level => level.Elevation)
                .ToList();
        }

        public static List<string> GetSections()
        {
            var sections = _data.Select(bar => bar.Section).Distinct().OrderBy(section => section).ToList();
            sections.Insert(0, "(все)");
            return sections;
        }

        public static List<string> GetConstructionTypes()
        {
            var constructionTypes = _data.Select(bar => bar.ConstructionType).Distinct().OrderBy(type => type).ToList();
            constructionTypes.Insert(0, "(все)");
            return constructionTypes;
        }

        public static void AnalyzeDataByConstructionCount()
        {
            var constructionCountBars = new List<Bar>();

            foreach (var bar in _data)
            {
                if (bar.ConstructionCount > 1)
                {
                    for (int i = 1; i < bar.ConstructionCount; i++)
                    {
                        var counstructionCountBar = new Bar(bar.Class, bar.Diameter, bar.Mass, bar.Shape)
                        {
                            Id = bar.Id,
                            Ids = bar.Ids,
                            IdsAsString = bar.IdsAsString,
                            Position = bar.Position,
                            Length = bar.Length,
                            Count = bar.Count,
                            CountType = bar.CountType,
                            Level = bar.Level,
                            Section = bar.Section,
                            ConstructionType = bar.ConstructionType,
                            ConstructionTypeEnum = bar.ConstructionTypeEnum,
                            ConstructionMark = bar.ConstructionMark,
                            ConstructionCount = bar.ConstructionCount,
                            TypicalFloor = bar.TypicalFloor,
                            TypicalFloorCount = bar.TypicalFloorCount,
                            ElementType = RebarElementType.Virtual,
                            DiameterClassInfo = bar.DiameterClassInfo,
                            CountTypeInfo = bar.CountTypeInfo,
                        };

                        constructionCountBars.Add(counstructionCountBar);
                    }
                }
            }

            foreach (var bar in constructionCountBars)
            {
                AddBar(bar);
            }
        }

        public static void AnalyzeDataByTypicalFloorCount(List<TypicalFloor> typicalFloors)
        {
            var typicalFloorBars = new List<Bar>();

            foreach (var bar in _data)
            {
                if (bar.TypicalFloorCount > 1)
                {
                    var typicalRebarLevels = GetTypicalRebarLevels(bar, typicalFloors);

                    if (typicalRebarLevels == null)
                    {
                        return;
                    }

                    foreach (var typicalRebarLevel in typicalRebarLevels)
                    {
                        var typicalFloorBar = new Bar(bar.Class, bar.Diameter, bar.Mass, bar.Shape)
                        {
                            Id = bar.Id,
                            Ids = bar.Ids,
                            IdsAsString = bar.IdsAsString,
                            Position = bar.Position,
                            Length = bar.Length,
                            Count = bar.Count,
                            CountType = bar.CountType,
                            Level = typicalRebarLevel,
                            Section = bar.Section,
                            ConstructionType = bar.ConstructionType,
                            ConstructionTypeEnum = bar.ConstructionTypeEnum,
                            ConstructionMark = bar.ConstructionMark,
                            ConstructionCount = bar.ConstructionCount,
                            TypicalFloor = bar.TypicalFloor,
                            TypicalFloorCount = bar.TypicalFloorCount,
                            ElementType = RebarElementType.Virtual,
                            DiameterClassInfo = bar.DiameterClassInfo,
                            CountTypeInfo = bar.CountTypeInfo,
                        };

                        if (typicalFloorBar.Level.Name != bar.Level.Name)
                        {
                            typicalFloorBars.Add(typicalFloorBar);
                        }
                    }
                }
            }

            foreach (var typicalFloorBar in typicalFloorBars)
            {
                AddBar(typicalFloorBar);
            }
        }

        private static List<RebarLevel> GetTypicalRebarLevels(Bar bar, List<TypicalFloor> typicalFloors)
        {
            return typicalFloors.Where(typicalFloor => typicalFloor.ConstructionTypeEnum == bar.ConstructionTypeEnum)
                .Where(typicalFloor => typicalFloor.Floor == bar.TypicalFloor)
                .SelectMany(typicalFloor => typicalFloor.Levels)
                .ToList();
        }
    }
}
