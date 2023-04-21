using System;
using System.Collections.Generic;

namespace RebarsToExcel.Models
{
    public class PositionComparer : IComparer<string>
    {
		public int Compare(string x, string y)
		{
			if (x == null && y == null)
				return 0;

			if (x == null)
				return -1;

			if (y == null)
				return 1;

			if (IsDigitsOnly(x) && IsDigitsOnly(y))
            {
				var xInt = Convert.ToInt32(x);
				var yInt = Convert.ToInt32(y);

				return xInt.CompareTo(yInt);
            }
			
			return x.CompareTo(y);
		}

		public int GetHashCode(string obj)
		{
			return GetHashCode();
		}

		private bool IsDigitsOnly(string str)
		{
			foreach (char c in str)
			{
				if (c < '0' || c > '9')
					return false;
			}

			return true;
		}
	}
}
