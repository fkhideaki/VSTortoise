Imports System
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE100
Imports System.Diagnostics



Public Class TPTypeCehck
	Public Enum EngineType
		None
		Git
		Svn
	End Enum

	Public Shared Function GetEngineType(dir As String) As EngineType
		If dir Is Nothing Then Return EngineType.None

		Do While (dir.Length >= 0)
			If System.IO.Directory.Exists(dir + "\\" + ".git") Then Return EngineType.Git
			If System.IO.Directory.Exists(dir + "\\" + ".svn") Then Return EngineType.Svn
			Dim s As String = System.IO.Directory.GetParent(dir).ToString()
			If s = dir Then Return EngineType.None
		Loop

		Return EngineType.None
	End Function

	Public Shared Function GetEngine(parent As DTE2) As TProcEngine
		Dim dir As String = IDEPath.GetCurrentDocDir(parent)
		Dim t As EngineType = GetEngineType(dir)
		Select Case t
			Case EngineType.Git
				Return New TGit(parent)
			Case EngineType.Svn
				Return New TSvn(parent)
			Case Else
				Return New TNull(parent)
		End Select
	End Function
End Class
