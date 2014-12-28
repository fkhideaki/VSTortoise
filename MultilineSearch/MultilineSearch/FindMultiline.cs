using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using EnvDTE;
using EnvDTE80;

namespace MultilineSearch
{
	class FindMultiline
	{
		private DTE2 DTE;

		public FindMultiline(DTE2 dte)
		{
			DTE = dte;
		}

		public void Execute()
		{
			String find_str = GetMultilineFindPattern();
			if (find_str == null)
				return;

			Find f = DTE.Find;
			f.FindWhat = find_str;
			f.MatchWholeWord = false;
			f.MatchCase = false;
			f.Backwards = false;
			f.MatchInHiddenText = true;
			f.Target = vsFindTarget.vsFindTargetCurrentDocument;
			f.PatternSyntax = vsFindPatternSyntax.vsFindPatternSyntaxRegExpr;
			f.Action = vsFindAction.vsFindActionFind;
			//DTE.ExecuteCommand("Edit.FindNext");
		}

		string GetMultilineFindPattern()
		{
			TextSelection sel = (TextSelection)DTE.ActiveDocument.Selection;
			string s = (string)sel.Text;
			if (s == null)
				return null;
			if (s == "")
				return null;

			return AvoidRegExprString(s);
		}

		string AvoidRegExprString(string src)
		{
			if (src == null)
				return "";

			string avoid_reg = (string)src.Clone();
			avoid_reg = avoid_reg.Replace("\\", "\\\\");

			avoid_reg = avoid_reg.Replace("@", "\\@");
			avoid_reg = avoid_reg.Replace("!", "\\!");
			avoid_reg = avoid_reg.Replace("\"", "\\\"");
			avoid_reg = avoid_reg.Replace("#", "\\#");
			avoid_reg = avoid_reg.Replace("$", "\\$");
			avoid_reg = avoid_reg.Replace("%", "\\%");
			avoid_reg = avoid_reg.Replace("&", "\\&");
			avoid_reg = avoid_reg.Replace("'", "\\'");
			avoid_reg = avoid_reg.Replace(":", "\\:");
			avoid_reg = avoid_reg.Replace("{", "\\{");
			avoid_reg = avoid_reg.Replace("}", "\\}");
			avoid_reg = avoid_reg.Replace("[", "\\[");
			avoid_reg = avoid_reg.Replace("]", "\\]");
			avoid_reg = avoid_reg.Replace("(", "\\(");
			avoid_reg = avoid_reg.Replace(")", "\\)");
			avoid_reg = avoid_reg.Replace("<", "\\<");
			avoid_reg = avoid_reg.Replace(">", "\\>");
			avoid_reg = avoid_reg.Replace("+", "\\+");
			avoid_reg = avoid_reg.Replace("-", "\\-");
			avoid_reg = avoid_reg.Replace("/", "\\/");
			avoid_reg = avoid_reg.Replace("*", "\\*");
			avoid_reg = avoid_reg.Replace("=", "\\=");
			avoid_reg = avoid_reg.Replace(".", "\\.");
			avoid_reg = avoid_reg.Replace("?", "\\?");
			avoid_reg = avoid_reg.Replace("^", "\\^");
			avoid_reg = avoid_reg.Replace("~", "\\~");
			avoid_reg = avoid_reg.Replace("|", "\\|");
			avoid_reg = avoid_reg.Replace("\r", "\\r");
			avoid_reg = avoid_reg.Replace("\n", "\\r\\n");
			avoid_reg = avoid_reg.Replace("\\r\\r\\n", "\\r\\n");

			return avoid_reg;
		}
	}
}
