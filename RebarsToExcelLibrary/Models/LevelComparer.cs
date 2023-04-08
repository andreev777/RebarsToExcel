using System.Collections.Generic;

namespace RebarsToExcel.Models
{
    public class LevelComparer : IEqualityComparer<RebarLevel>
    {
		public bool Equals(RebarLevel x, RebarLevel y)
		{
			if (x.Name == y.Name) 
				return true;

			return false;
		}

		public int GetHashCode(RebarLevel obj)
		{
			return GetHashCode();
		}
	}
}
