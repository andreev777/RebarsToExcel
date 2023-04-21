using System.Collections.Generic;
using System.Linq;

namespace RebarsToExcel.Models
{
    /// <summary>
    /// Хранилище всех деталей.
    /// </summary>
    public static class BarsData
    {
        private static IList<Bar> _data = new List<Bar>();
        /// <summary>
        /// Добавить деталь в хранилище.
        /// </summary>
        public static void AddBar(Bar bar)
        {
            if (bar.CountType == 1) //Поштучно
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
                _data.Add(bar);
            }

            else if (bar.CountType == 2) //Метры погонные
            {
                foreach (var existedBar in _data)
                {
                    if (existedBar.Position == bar.Position
                        && existedBar.Class == bar.Class
                        && existedBar.Diameter == bar.Diameter
                        && existedBar.Mass == bar.Mass
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
                _data.Add(bar);
            }
        }
        /// <summary>
        /// Получить коллекцию всех деталей.
        /// </summary>
        /// <returns>Возвращает коллекцию деталей, отсортированную сначала по секции,
        /// затем по уровню, типу основы, метке основы и позиции.</returns>
        public static IList<Bar> GetData()
        {
            return _data.OrderBy(x => x.Section)
                .ThenBy(x => x.Level.Elevation)
                .ThenBy(x => x.ConstructionType)
                .ThenBy(x => x.ConstructionMark)
                .ThenBy(x => x.Position, new PositionComparer())
                .ToList();
        }
        /// <summary>
        /// Получить коллекцию всех уникальных уровней.
        /// </summary>
        /// <returns>Возращает коллекцию уровней, отсортированную по возвышению над землей.</returns>
        public static IList<RebarLevel> GetLevels()
        {
            return _data.Select(bar => bar.Level)
                .Distinct(new LevelComparer())
                .OrderBy(level => level.Elevation)
                .ToList();
        }
        /// <summary>
        /// Получить коллекцию всех уникальных секций.
        /// </summary>
        /// <returns>Возращает коллекцию секций, отсортированную по алфавиту.</returns>
        public static IList<string> GetSections()
        {
            return _data.Select(bar => bar.Section).Distinct().OrderBy(section => section).ToList();
        }
        /// <summary>
        /// Получить коллекцию всех уникальных типов основ.
        /// </summary>
        /// <returns>Возращает коллекцию типов основ, отсортированную по алфавиту.</returns>
        public static IList<string> GetConstructionTypes()
        {
            return _data.Select(bar => bar.ConstructionType).Distinct().OrderBy(type => type).ToList();
        }
        /// <summary>
        /// Анализировать детали по количеству основ.
        /// </summary>
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
                            Position = bar.Position,
                            PositionWithShapeMark = bar.PositionWithShapeMark,
                            Length = bar.Length,
                            ShapeImagePath = bar.ShapeImagePath,
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
                            DiameterClassLengthInfo = bar.DiameterClassLengthInfo,
                            CountTypeInfo = bar.CountTypeInfo,
                            Sizes = bar.Sizes,
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
        /// <summary>
        /// Анализировать детали по количеству типовых этажей.
        /// </summary>
        public static void AnalyzeDataByTypicalFloorCount(IList<TypicalFloor> typicalFloors)
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
                            Position = bar.Position,
                            PositionWithShapeMark = bar.PositionWithShapeMark,
                            Length = bar.Length,
                            ShapeImagePath = bar.ShapeImagePath,
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
                            DiameterClassLengthInfo = bar.DiameterClassLengthInfo,
                            CountTypeInfo = bar.CountTypeInfo,
                            Sizes = bar.Sizes,
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

        private static IList<RebarLevel> GetTypicalRebarLevels(Bar bar, IList<TypicalFloor> typicalFloors)
        {
            return typicalFloors.Where(typicalFloor => typicalFloor.ConstructionTypeEnum == bar.ConstructionTypeEnum)
                .Where(typicalFloor => typicalFloor.Floor == bar.TypicalFloor)
                .SelectMany(typicalFloor => typicalFloor.Levels)
                .ToList();
        }
    }
}
