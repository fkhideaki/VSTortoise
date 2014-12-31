Imports System
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE100
Imports System.Diagnostics


Public MustInherit Class Tortoise
	Private DTE As DTE2 = Nothing

	Protected MustOverride Function GetTPName() As String

	Sub New(parent As DTE2)
		DTE = parent
	End Sub

	Private Function CreateCmd(cmd As String, path As String) As String
		Dim sc As String = " /command:" + cmd
		Dim sp As String = " /path:" + """" + path + """"
		Return GetTPName() + sc + sp + " /notempfile"
	End Function

	Private Function GetCurrentFilePath() As String
		Dim filename As String = ""
		If Not get_current_doc(DTE, filename) Then Return Nothing
		Return filename
	End Function

	Private Function GetCurrentDirPath() As String
		Dim dirname As String = ""
		If Not get_current_doc_directory(DTE, dirname) Then Return Nothing
		Return dirname
	End Function

	Private Sub ExecutePathCmd(cmd As String, path As String)
		If path Is Nothing Then Return

		Dim cmdline As String = CreateCmd(cmd, path)
		DTE.ExecuteCommand("Tools.Shell", cmdline)
	End Sub

	' Diffウィンドウを表示
	Public Sub Diff()
		ExecutePathCmd("diff", GetCurrentFilePath())
	End Sub

	' Logウィンドウを表示
	Public Sub Log()
		ExecutePathCmd("log", GetCurrentFilePath())
	End Sub

	' カレントディレクトリに対するDiffを実行
	Public Sub Diff_CurrentDir()
		ExecutePathCmd("diff", GetCurrentDirPath())
	End Sub

	' カレントディレクトリに対するLogを実行
	Public Sub Log_CurrentDir()
		ExecutePathCmd("log", GetCurrentDirPath())
	End Sub

	'現在のドキュメントのフルパスを返す
	Function get_current_doc(DTE As DTE2, ByRef filename As String) As Boolean
		filename = DTE.ActiveDocument.FullName
		If Not System.IO.File.Exists(filename) Then Return False
		Return True
	End Function

	'現在のドキュメントが保存されたディレクトリのフルパスを返す
	Function get_current_doc_directory(DTE As DTE2, ByRef dirname As String) As Boolean
		Dim filename As String = ""
		If Not get_current_doc(DTE, filename) Then Return False
		dirname = System.IO.Path.GetDirectoryName(filename)
		Return True
	End Function
End Class
