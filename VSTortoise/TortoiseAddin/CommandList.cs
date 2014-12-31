using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EnvDTE;
using EnvDTE80;


namespace TortoiseAddin
{
	class CommandList
	{
		Dictionary<string, CommandDef> NameToCommand = new Dictionary<string,CommandDef>();
		List<CommandDef> Commands = new List<CommandDef>();


		public void AddCommand(string name_text, CommandDef.CommandMethod method)
		{
			AddCommand(name_text, name_text, method);
		}

		public void AddCommand(string name, string text, CommandDef.CommandMethod method)
		{
			CommandDef cmd = new CommandDef();
			cmd.Name = name;
			cmd.Text = text;
			cmd.Index = NumCommands() + 1;
			cmd.Method = method;

			Commands.Add(cmd);
			NameToCommand.Add(name, cmd);
		}

		public CommandDef GetCommand(int idx)
		{
			return Commands[idx];
		}

		public CommandDef GetCommand(string Name)
		{
			return NameToCommand[Name];
		}

		public int NumCommands()
		{
			return Commands.Count();
		}
	}
}
