using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NPOIPlusConsoleExample
{
	public class ExampleData
	{
		public readonly int ID;
		public readonly string Name;
		public readonly DateTime DateOfBirth;

		public ExampleData(int id, string name, DateTime dateOfBirth)
		{
			ID = id;
			Name = name;
			DateOfBirth = dateOfBirth;
		}

	}
}
