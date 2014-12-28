Imports System
Imports EnvDTE
Imports EnvDTE80
Imports System.Diagnostics


Public Module tools

    ' inputbox で入力されたディレクトリの全ファイルのリストをカレントキャレットに挿入する
	Sub ListupFilesToCurrentEditor(DTE As DTE2)
		Dim dir As String = InputBox("directory path")

		If IO.Directory.Exists(dir) = False Then Return

		Dim sl() As String = IO.Directory.GetFiles(dir)

		For Each s As String In sl
			DTE.ActiveDocument.Selection.Text = s + vbNewLine
		Next
	End Sub

End Module
