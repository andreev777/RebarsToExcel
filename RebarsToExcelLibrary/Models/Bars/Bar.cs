using Autodesk.Revit.DB;
using System.Collections.Generic;
using System.Linq;
using RebarsToExcel.Models.Abstractions;

namespace RebarsToExcel.Models
{
    /// <summary>
    /// Класс детали.
    /// </summary>
    public class Bar : RebarAbstraction
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
        /// Позиция (Марка).
        /// </summary>
        public string Position { get; set; }
        /// <summary>
        /// Позиция (Марка) с дополнительным обозначением для гнутой детали.
        /// </summary>
        public string PositionWithShapeMark { get; set; }
        /// <summary>
        /// Путь к файлу с изображением эскиза детали.
        /// </summary>
        public string ShapeImagePath { get; set; }
        /// <summary>
        /// _Тип подсчета количества: 1 - шт., 2 - м.п.
        /// </summary>
        public int CountType { get; set; }
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
        /// Строковое представление информации о детали. Например, "Ø12А500, L=3000".
        /// </summary>
        public string DiameterClassLengthInfo { get; set; }
        /// <summary>
        /// Строковое представление типа подсчета количества. Например, "шт." или "м.п."
        /// </summary>
        public string CountTypeInfo { get; set; }
        /// <summary>
        /// Коллекция размеров детали.
        /// </summary>
        public IDictionary<BarSize, double> Sizes = new Dictionary<BarSize, double>();

        public Bar(string rebarClass, double diameter, double mass, string shape)
        {
            Class = rebarClass;
            Diameter = diameter;
            Mass = mass;
            Shape = shape;

            ConstructionTypeEnum = Models.ConstructionType.Unknown;
            ElementType = RebarElementType.Real;
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
        /// Получить список Id в виде строки.
        /// </summary>
        public string GetIdsAsString()
        {
            var idsAsIntegerCollection = Ids.Select(x => x.IntegerValue);
            return string.Join(", ", idsAsIntegerCollection);
        }
    }
}