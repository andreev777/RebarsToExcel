namespace RebarsToExcel.Models
{
    /// <summary>
    /// Класс размера детали.
    /// </summary>
    public class BarSize
    {
        public byte Id { get; set; }
        public string Name { get; set; }

        public BarSize(byte id, string name)
        {
            Id = id;
            Name = name;
        }
    }
}
