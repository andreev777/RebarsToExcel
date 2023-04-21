using RebarsToExcel.Models.Abstractions;

namespace RebarsToExcel.Models
{
    /// <summary>
    /// Класс детали сборочной единицы.
    /// </summary>
    public class Rebar : RebarAbstraction
    {
        /// <summary>
        /// _Тип сборки.
        /// </summary>
        public string TypeOfAssembly { get; set; }
        /// <summary>
        /// _Метка сборки.
        /// </summary>
        public string MarkOfAssembly { get; set; }

        public Rebar(string rebarClass, double diameter, double mass, string shape)
        {
            Class = rebarClass;
            Diameter = diameter;
            Mass = mass;
            Shape = shape;
        }
        /// <summary>
        /// Получить информацию о детали без длины. Например, "Ø12А500".
        /// </summary>
        public string GetClassDiameterInfoString() => $"⌀{Diameter}{Class}";
        /// <summary>
        /// Получить информацию о детали. Например, "Ø12А500, L=3000".
        /// </summary>
        public string GetClassDiameterCountInfoString() => $"⌀{Diameter}{Class}, L={Length}мм - {Count}шт.";
    }
}