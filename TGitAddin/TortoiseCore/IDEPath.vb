Imports System
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE100
Imports System.Diagnostics


Public Class IDEPath
	'現在のドキュメントのフルパスを返す
	Public Shared Function GetCurrentDoc(DTE As DTE2, ByRef filename As String) As Boolean
		filename = DTE.ActiveDocument.FullName
		Return System.IO.File.Exists(filename)
	End Function

	'現在のドキュメントが保存されたディレクトリのフルパスを返す
	Public Shared Function GetCurrentDocDir(DTE As DTE2, ByRef dirname As String) As Boolean
		Dim filename As String = ""
		If Not GetCurrentDoc(DTE, filename) Then Return False
		dirname = System.IO.Path.GetDirectoryName(filename)
		Return True
	End Function
End Class
