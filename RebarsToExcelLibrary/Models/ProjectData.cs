namespace RebarsToExcel.Models
{
    /// <summary>
    /// Данные проекта.
    /// </summary>
    public static class ProjectData
    {
        /// <summary>
        /// Имя файла с таблицей Excel.
        /// </summary>
        public static string FileName { get; set; }
        /// <summary>
        /// Шифр объекта.
        /// </summary>
        public static string ProjectCode { get; set; }
        /// <summary>
        /// Наименование объекта.
        /// </summary>
        public static string ProjectName { get; set; }
        /// <summary>
        /// Наименование здания.
        /// </summary>
        public static string BuildingName { get; set; }
    }
}
