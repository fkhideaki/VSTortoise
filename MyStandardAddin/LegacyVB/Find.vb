Imports System
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE100
Imports System.Diagnostics

Public Class DTEFind

	Dim DTE As DTE2

	Public Sub New(dte2_instance As DTE2)
		DTE = dte2_instance
	End Sub

	' 選択文字列を正規表現化し, 削除検索
	Sub RemoveReplaceWithAvoidRegExpr()
		Dim s As String = DTE.ActiveDocument.Selection.text
		If s Is Nothing Then Return
		If s = "" Then Return

		Dim find_str As String = AvoidRegExprString(s)

		DTE.ExecuteCommand("Edit.Replace")
		DTE.Find.FindWhat = find_str
		DTE.Find.ReplaceWith = "\1"
		DTE.Find.Target = vsFindTarget.vsFindTargetCurrentDocument
		DTE.Find.MatchCase = False
		DTE.Find.MatchWholeWord = True
		DTE.Find.Backwards = False
		DTE.Find.MatchInHiddenText = True
		DTE.Find.PatternSyntax = vsFindPatternSyntax.vsFindPatternSyntaxRegExpr
		DTE.Find.Action = vsFindAction.vsFindActionFind
	End Sub

	' オプション簡易指定検索
	Sub FindOptByInputBox()
		DTE.ExecuteCommand("Edit.Find")
		DTE.Find.Target = vsFindTarget.vsFindTargetCurrentDocument
		DTE.Find.MatchCase = False
		DTE.Find.MatchWholeWord = True
		DTE.Find.Backwards = False
		DTE.Find.MatchInHiddenText = True
		DTE.Find.PatternSyntax = vsFindPatternSyntax.vsFindPatternSyntaxRegExpr
		DTE.Find.Action = vsFindAction.vsFindActionFind
	End Sub

	' 現在の行を検索文字列として検索
	Sub SetLineToSearchString()
		DTE.ActiveDocument.Selection.EndOfLine()
		DTE.ActiveDocument.Selection.StartOfLine(vsStartOfLineOptions.vsStartOfLineOptionsFirstText, True)
		DTE.ExecuteCommand("Edit.FindNextSelected")
	End Sub


	' 検索ダイアログ表示
	Sub FindDiag_Solution()
		DTE.ExecuteCommand("Edit.FindinFiles")
		DTE.Find.SearchPath = "ソリューション全体"
	End Sub

	Sub FindDiag_CurrentProject()
		DTE.ExecuteCommand("Edit.FindinFiles")
		DTE.Find.SearchPath = "現在のプロジェクト"
	End Sub

	Sub FindDiag_CurrentDocument()
		DTE.ExecuteCommand("Edit.FindinFiles")
		DTE.Find.SearchPath = "現在のドキュメント"
	End Sub


	' C++ 用。"#define TAG*"という文字列を検索するダイアログを表示する。
	Sub FindDiag_AllTags()
		DTE.ExecuteCommand("Edit.FindinFiles")
		DTE.Find.SearchPath = "ソリューション全体"
		DTE.Find.FindWhat = "#define TAG"
	End Sub

	' インクリメント検索に現在の文字列をセットする
	Sub SetInclFinder()
		If DTE.ActiveDocument.Selection.text = "" Then
			DTE.ExecuteCommand("Edit.SelectCurrentWord")
		End If
		Dim st As String = DTE.ActiveDocument.Selection.Text
		If st.Length > 100 Then Return

		Dim ca() As Char = st.ToCharArray()

		DTE.ActiveDocument.ActiveWindow.Object.ActivePane.IncrementalSearch.StartForward()
		For Each c As Char In ca
			DTE.ActiveDocument.ActiveWindow.Object.ActivePane.IncrementalSearch.AppendCharAndSearch(AscW(c.ToString()))
		Next
		DTE.ActiveDocument.ActiveWindow.Object.ActivePane.IncrementalSearch.Exit()
	End Sub



	Private Sub OpenFile(ByVal path As String)
		DTE.ItemOperations.OpenFile(path)
	End Sub

	Private Function GetCurrentLine() As String
		DTE.ActiveDocument.Selection.EndOfLine()
		DTE.ActiveDocument.Selection.StartOfLine(vsStartOfLineOptions.vsStartOfLineOptionsFirstText)
		DTE.ActiveDocument.Selection.EndOfLine(True)
		Return DTE.ActiveDocument.Selection.Text
	End Function


	Sub GoToFindResultString()
		Dim str As String = GetCurrentLine()

		Dim vs() As String = str.Split("(")
		If vs Is Nothing Then Return
		If vs.Length < 2 Then Return

		Dim file_path As String = vs(0)
		If IO.File.Exists(file_path) = False Then Return

		Dim file_name As String = IO.Path.GetFileName(file_path)

		Dim vs2() As String = vs(1).Split(")")
		If vs2 Is Nothing Then Return
		If vs2.Length < 1 Then Return

		Dim target_line As Integer
		Try
			target_line = Integer.Parse(vs2(0))
		Catch ex As Exception
			Return
		End Try

		OpenFile(file_path)
		DTE.Windows.Item(file_name).Activate()
		DTE.ActiveDocument.Selection.GotoLine(target_line)
	End Sub


	' 検索結果のカスタム
	Sub MakeCustomFindLinks()
		Dim iu As New IDE_Utility(DTE)

		Dim out_msg As String = DTE.ActiveDocument.Selection.text
		If out_msg = "" Then
			iu.GetOutputWindowPane("CustomFind").Activate()
		Else
			iu.GetOutputWindowPane("CustomFind").Clear()
			iu.GetOutputWindowPane("CustomFind").OutputString(out_msg + vbNewLine)
			DTE.ExecuteCommand("Window.ActivateDocumentWindow")
		End If
	End Sub


	' ペースト → 検索を次に進めるをセットで実行
	Sub PasteAndGoNextFind()
		DTE.ActiveDocument.Selection.Paste()
		DTE.ExecuteCommand("Edit.FindNext")
	End Sub


	Sub RemoveRegExpr()
		DTE.ExecuteCommand("Edit.Replace")
		DTE.Find.Action = vsFindAction.vsFindActionReplace
		DTE.Find.FindWhat = "{\@|\!|\""|\#|\$|\%|\&|\'|\{|\}|\(|\)|\<|\>|\+|\-|\/|\*|\=|\!|\<|\.|\?|\^|\~|\||\\}"
		DTE.Find.ReplaceWith = "\\\1"
		DTE.Find.Target = vsFindTarget.vsFindTargetCurrentDocument
		DTE.Find.MatchCase = False
		DTE.Find.MatchWholeWord = True
		DTE.Find.Backwards = False
		DTE.Find.MatchInHiddenText = True
		DTE.Find.PatternSyntax = vsFindPatternSyntax.vsFindPatternSyntaxRegExpr
	End Sub


	Function AvoidRegExprString(ByVal src As String) As String
		Dim s As String = DTE.ActiveDocument.Selection.text
		If s Is Nothing Then Return ""
		If s = "" Then Return ""

		Dim avoid_reg As String = s.Clone()
		avoid_reg = avoid_reg.Replace("\", "\\")

		avoid_reg = avoid_reg.Replace("@", "\@")
		avoid_reg = avoid_reg.Replace("!", "\!")
		avoid_reg = avoid_reg.Replace("""", "\""")
		avoid_reg = avoid_reg.Replace("#", "\#")
		avoid_reg = avoid_reg.Replace("$", "\$")
		avoid_reg = avoid_reg.Replace("%", "\%")
		avoid_reg = avoid_reg.Replace("&", "\&")
		avoid_reg = avoid_reg.Replace("'", "\'")
		avoid_reg = avoid_reg.Replace(":", "\:")
		avoid_reg = avoid_reg.Replace("{", "\{")
		avoid_reg = avoid_reg.Replace("}", "\}")
		avoid_reg = avoid_reg.Replace("[", "\[")
		avoid_reg = avoid_reg.Replace("]", "\]")
		avoid_reg = avoid_reg.Replace("(", "\(")
		avoid_reg = avoid_reg.Replace(")", "\)")
		avoid_reg = avoid_reg.Replace("<", "\<")
		avoid_reg = avoid_reg.Replace(">", "\>")
		avoid_reg = avoid_reg.Replace("+", "\+")
		avoid_reg = avoid_reg.Replace("-", "\-")
		avoid_reg = avoid_reg.Replace("/", "\/")
		avoid_reg = avoid_reg.Replace("*", "\*")
		avoid_reg = avoid_reg.Replace("=", "\=")
		avoid_reg = avoid_reg.Replace(".", "\.")
		avoid_reg = avoid_reg.Replace("?", "\?")
		avoid_reg = avoid_reg.Replace("^", "\^")
		avoid_reg = avoid_reg.Replace("~", "\~")
		avoid_reg = avoid_reg.Replace("|", "\|")
		avoid_reg = avoid_reg.Replace(vbNewLine, "\r\n")

		Return avoid_reg
	End Function

	Private Function GetMultilineFindPattern() As String
		Dim s As String = DTE.ActiveDocument.Selection.text
		If s Is Nothing Then Return Nothing
		If s = "" Then Return Nothing

		Dim find_str As String = AvoidRegExprString(s)
		Return find_str
	End Function

	Sub SetMultilineFindTargetWithRegExpr()
		Dim find_str As String = GetMultilineFindPattern()
		If find_str Is Nothing Then Return

		DTE.Find.FindWhat = find_str
		DTE.Find.MatchWholeWord = True
		DTE.Find.Backwards = False
		DTE.Find.MatchInHiddenText = True
		DTE.Find.Target = vsFindTarget.vsFindTargetCurrentDocument
		DTE.Find.MatchCase = False
		DTE.Find.Action = vsFindAction.vsFindActionFind
		DTE.Find.PatternSyntax = vsFindPatternSyntax.vsFindPatternSyntaxRegExpr
		DTE.ExecuteCommand("Edit.FindNext")
	End Sub

	Sub SetMultilineFindTargetWithRegExprVS2012()
		Dim find_str As String = GetMultilineFindPattern()
		If find_str Is Nothing Then Return

		DTE.Find.FindWhat = find_str
		'DTE.Find.MatchWholeWord = True
		'DTE.Find.MatchCase = False
		DTE.Find.Backwards = False
		DTE.Find.MatchInHiddenText = True
		DTE.Find.Target = vsFindTarget.vsFindTargetCurrentDocument
		DTE.Find.PatternSyntax = vsFindPatternSyntax.vsFindPatternSyntaxRegExpr
		DTE.Find.Action = vsFindAction.vsFindActionFind
		'DTE.ExecuteCommand("Edit.FindNext")
	End Sub

End Class
