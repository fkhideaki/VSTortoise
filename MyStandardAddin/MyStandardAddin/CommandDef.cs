using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;


namespace MyStandardAddin
{
	class CommandDef
	{
		public string Name;
		public string Text;
		public int Index;

		public delegate void CommandMethod(DTE2 vs_app);
		public CommandMethod Method = null;
	}
}
