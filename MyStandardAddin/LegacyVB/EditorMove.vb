Imports System
Imports EnvDTE
Imports EnvDTE80
Imports System.Diagnostics

Public Module EditorMove

    ' 次の行を作成して先頭に移動
	Public Sub GotoNewNextLine(DTE As DTE2)
		DTE.ActiveDocument.Selection.EndOfLine()
		DTE.ActiveDocument.Selection.NewLine()
	End Sub

End Module
