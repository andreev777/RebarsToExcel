using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;

namespace RebarsToExcel.Models
{
    /// <summary>
    /// Класс сборочной единицы.
    /// </summary>
    public class RebarAssembly
    {
        /// <summary>
        /// ElementId.
        /// </summary>
        public ElementId Id { get; set; }
        /// <summary>
        /// Коллекция уникальных ElementId.
        /// </summary>
        public HashSet<ElementId> Ids { get; set; } = new HashSet<ElementId>();
        /// <summary>
        /// Тип сборочной единицы (параметр Описание в модели).
        /// </summary>
        public string Type { get; set; }
        /// <summary>
        /// Марка.
        /// </summary>
        public string Mark { get; set; }
        /// <summary>
        /// _Наименование.
        /// </summary>
        public string Definition { get; set; }
        /// <summary>
        /// Группа модели.
        /// </summary>
        public string GroupModel { get; set; }
        /// <summary>
        /// Количество.
        /// </summary>
        public int Count { get; set; }
        /// <summary>
        /// _Масса.
        /// </summary>
        public double Mass { get; set; }
        /// <summary>
        /// Уровень.
        /// </summary>
        public RebarLevel Level { get; set; }
        /// <summary>
        /// _Секция.
        /// </summary>
        public string Section { get; set; }
        /// <summary>
        /// _Тип основы.
        /// </summary>
        public string ConstructionType { get; set; }
        /// <summary>
        /// Тип основы (перечисление).
        /// </summary>
        public ConstructionType ConstructionTypeEnum { get; set; }
        /// <summary>
        /// _Метка основы.
        /// </summary>
        public string ConstructionMark { get; set; }
        /// <summary>
        /// _Количество основ.
        /// </summary>
        public int ConstructionCount { get; set; }
        /// <summary>
        /// _Типовой этаж.
        /// </summary>
        public int TypicalFloor { get; set; }
        /// <summary>
        /// _Количество типовых этажей.
        /// </summary>
        public int TypicalFloorCount { get; set; }
        /// <summary>
        /// Тип элемента: Real - реальный элемент модели, Virtual - виртуальный элемент модели, полученный аналитическим способом.
        /// </summary>
        public RebarElementType ElementType { get; set; }
        /// <summary>
        /// Коллекция арматуры, из которой состоит сборочная единица.
        /// </summary>
        public IList<Rebar> Rebars { get; set; } = new List<Rebar>();

        public RebarAssembly(string type, string mark, string groupModel, double mass)
        {
            Type = type;
            Mark = mark;
            GroupModel = groupModel;
            Mass = mass;
            Count = 1;

            ConstructionTypeEnum = Models.ConstructionType.Unknown;
            ElementType = RebarElementType.Real;
        }
        /// <summary>
        /// Добавить арматуру в коллекцию арматуры сборочной единицы.
        /// </summary>
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
        /// <summary>
        /// Добавить экземпляр ElementId к существующей коллекции Ids.
        /// </summary>
        public void AddId(ElementId id)
        {
            if (id != null)
            {
                Ids.Add(id);
            }
        }
        /// <summary>
        /// Получить информацию о детали в виде строки. Например, "Ø12А500, L=3000".
        /// </summary>
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
        /// <summary>
        /// Возвращает список Id в виде строки.
        /// </summary>
        public string GetIdsAsString()
        {
            var idsAsIntegerCollection = Ids.Select(x => x.IntegerValue);
            return string.Join(", ", idsAsIntegerCollection);
        }
    }
}