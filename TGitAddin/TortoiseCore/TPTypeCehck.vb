Imports System
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE100
Imports System.Diagnostics



Public Class TPTypeCehck
	Public Shared Function GetEngine(parent As DTE2) As TProcEngine
		Dim dir As String = IDEPath.GetCurrentDocDir(parent)
		If dir Is Nothing Then Return New TNull(parent)

		Do While (dir.Length >= 0)
			If System.IO.Directory.Exists(dir + "\\" + ".git") Then Return New TGit(parent)
			If System.IO.Directory.Exists(dir + "\\" + ".svn") Then Return New TSvn(parent)
			Dim s As String = System.IO.Directory.GetParent(dir).ToString()
			If s = dir Then Return New TNull(parent)
		Loop

		Return New TNull(parent)
	End Function
End Class
