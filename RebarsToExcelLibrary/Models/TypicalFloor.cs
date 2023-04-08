using System.Collections.Generic;

namespace RebarsToExcel.Models
{
    public class TypicalFloor
    {
        public ConstructionType ConstructionTypeEnum { get; set; }
        public int Floor;
        public List<RebarLevel> Levels;

        public TypicalFloor(ConstructionType constructionTypeEnum, int floor, List<RebarLevel> levels)
        {
            ConstructionTypeEnum = constructionTypeEnum;
            Floor = floor;
            Levels = levels;
        }
    }
}
