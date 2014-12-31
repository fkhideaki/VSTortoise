Imports System
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE100
Imports System.Diagnostics


Public Class IDEPath
	'現在のドキュメントのフルパスを返す
	Public Shared Function GetCurrentDoc(DTE As DTE2, ByRef filename As String) As Boolean
		filename = GetCurrentDoc(DTE)
		Return Not filename Is Nothing
	End Function

	Public Shared Function GetCurrentDoc(DTE As DTE2) As String
		Dim filename As String = DTE.ActiveDocument.FullName
		If Not System.IO.File.Exists(filename) Then Return Nothing
		Return filename
	End Function

	'現在のドキュメントが保存されたディレクトリのフルパスを返す
	Public Shared Function GetCurrentDocDir(DTE As DTE2, ByRef dirname As String) As Boolean
		dirname = GetCurrentDocDir(DTE)
		Return Not dirname Is Nothing
	End Function

	Public Shared Function GetCurrentDocDir(DTE As DTE2) As String
		Dim filename As String = ""
		If Not GetCurrentDoc(DTE, filename) Then Return Nothing
		Return System.IO.Path.GetDirectoryName(filename)
	End Function
End Class
