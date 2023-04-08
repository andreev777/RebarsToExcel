using Prism.Mvvm;

namespace RebarsToExcel.Models
{
    public class RebarLevel : BindableBase
    {
        public string Name { get; set; }
        public double Elevation { get; set; }

        private bool isSelected;
        public bool IsSelected
        {
            get => isSelected;
            set
            {
                isSelected = value;
                RaisePropertyChanged(nameof(IsSelected));
            }
        }

        public RebarLevel(string name, double elevation)
        {
            Name = name;
            Elevation = elevation;
            IsSelected = false;
        }
    }
}

