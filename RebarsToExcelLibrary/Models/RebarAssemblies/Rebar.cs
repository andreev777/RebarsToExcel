namespace RebarsToExcel.Models
{
    public class Rebar
    {
        public string Class { get; set; }
        public double Diameter { get; set; } = 0;
        public double Length { get; set; } = 0;
        public int Count { get; set; } = 0;
        public double Mass { get; set; } = 0;
        public string Shape { get; set; }
        public string TypeOfAssembly { get; set; }
        public string MarkOfAssembly { get; set; }

        public Rebar(string rebarClass, double diameter, double mass, string shape)
        {
            Class = rebarClass;
            Diameter = diameter;
            Mass = mass;
            Shape = shape;
        }

        public string GetClassDiameterInfoString() => $"⌀{Diameter}{Class}";

        public string GetClassDiameterCountInfoString() => $"⌀{Diameter}{Class}, L={Length}мм - {Count}шт.";
    }
}