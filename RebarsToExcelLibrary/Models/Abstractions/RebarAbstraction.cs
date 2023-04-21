namespace RebarsToExcel.Models.Abstractions
{
    public abstract class RebarAbstraction
    {
        /// <summary>
        /// _Класс арматуры.
        /// </summary>
        public string Class { get; set; }
        /// <summary>
        /// _Диаметр стержня.
        /// </summary>
        public double Diameter { get; set; } = 0;
        /// <summary>
        /// _Длина стержня.
        /// </summary>
        public double Length { get; set; } = 0;
        /// <summary>
        /// Количество.
        /// </summary>
        public double Count { get; set; } = 0;
        /// <summary>
        /// _Масса.
        /// </summary>
        public double Mass { get; set; } = 0;
        /// <summary>
        /// _Форма стержня.
        /// </summary>
        public string Shape { get; set; }
    }
}
