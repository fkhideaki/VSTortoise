Imports System
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE100
Imports System.Diagnostics


Public MustInherit Class TProcEngine
	Private DTE As DTE2 = Nothing

	Sub New(parent As DTE2)
		DTE = parent
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

	Protected MustOverride Function GetTPName() As String

	Private Function CreateCmd(cmd As String, path As String) As String
		Dim sc As String = " /command:" + cmd
		Dim sp As String = " /path:" + """" + path + """"
		Return GetTPName() + sc + sp + " /notempfile"
	End Function

	Private Function GetCurrentFilePath() As String
		Dim filename As String = ""
		If Not IDEPath.GetCurrentDoc(DTE, filename) Then Return Nothing
		Return filename
	End Function

	Private Function GetCurrentDirPath() As String
		Dim dirname As String = ""
		If Not IDEPath.GetCurrentDocDir(DTE, dirname) Then Return Nothing
		Return dirname
	End Function

	Private Sub ExecutePathCmd(cmd As String, path As String)
		If path Is Nothing Then Return

		Dim cmdline As String = CreateCmd(cmd, path)
		DTE.ExecuteCommand("Tools.Shell", cmdline)
	End Sub
End Class
