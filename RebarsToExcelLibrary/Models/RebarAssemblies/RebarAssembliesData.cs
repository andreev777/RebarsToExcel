using System.Collections.Generic;
using System.Linq;

namespace RebarsToExcel.Models
{
    /// <summary>
    /// Хранилище всех сборочных единиц.
    /// </summary>
    public static class RebarAssembliesData
    {
        private static IList<RebarAssembly> _data = new List<RebarAssembly>();
        /// <summary>
        /// Добавить сборочную единицу в хранилище.
        /// </summary>
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
            _data.Add(rebarAssembly);
        }
        /// <summary>
        /// Получить коллекцию всех сборочных единиц.
        /// </summary>
        /// <returns>Возвращает коллекцию сборочных единиц, отсортированную сначала по секции,
        /// затем по уровню, типу основы, метке основы, типу и позиции.</returns>
        public static IList<RebarAssembly> GetData()
        {
            return _data.OrderBy(x => x.Section)
                .ThenBy(x => x.Level.Elevation)
                .ThenBy(x => x.ConstructionType)
                .ThenBy(x => x.ConstructionMark)
                .ThenBy(x => x.Type)
                .ThenBy(x => x.Mark)
                .ToList();
        }
        /// <summary>
        /// Получить коллекцию всех уникальных уровней.
        /// </summary>
        /// <returns>Возращает коллекцию уровней, отсортированную по возвышению над землей.</returns>
        public static IList<RebarLevel> GetLevels()
        {
            return _data.Select(rebarAssembly => rebarAssembly.Level)
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
            return _data.Select(rebarAssembly => rebarAssembly.Section)
                .Distinct()
                .OrderBy(section => section)
                .ToList();
        }
        /// <summary>
        /// Получить коллекцию всех уникальных типов основ.
        /// </summary>
        /// <returns>Возращает коллекцию типов основ, отсортированную по алфавиту.</returns>
        public static IList<string> GetConstructionTypes()
        {
            return _data.Select(rebarAssembly => rebarAssembly.ConstructionType)
                .Distinct()
                .OrderBy(type => type)
                .ToList();
        }
        /// <summary>
        /// Анализировать сборочные единицы по количеству основ.
        /// </summary>
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
        /// <summary>
        /// Анализировать сборочные единицы по количеству типовых этажей.
        /// </summary>
        public static void AnalyzeDataByTypicalFloorCount(IList<TypicalFloor> typicalFloors)
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

        private static IList<RebarLevel> GetTypicalRebarLevels(RebarAssembly rebarAssembly, IList<TypicalFloor> typicalFloors)
        {
            return typicalFloors.Where(typicalFloor => typicalFloor.ConstructionTypeEnum == rebarAssembly.ConstructionTypeEnum)
                .Where(typicalFloor => typicalFloor.Floor == rebarAssembly.TypicalFloor)
                .SelectMany(typicalFloor => typicalFloor.Levels)
                .ToList();
        }
    }
}
