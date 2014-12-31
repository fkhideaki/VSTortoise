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
		ExecutePathCmd("diff", IDEPath.GetCurrentDoc(DTE))
	End Sub

	' Logウィンドウを表示
	Public Sub Log()
		ExecutePathCmd("log", IDEPath.GetCurrentDoc(DTE))
	End Sub

	' カレントディレクトリに対するDiffを実行
	Public Sub Diff_CurrentDir()
		ExecutePathCmd("diff", IDEPath.GetCurrentDocDir(DTE))
	End Sub

	' カレントディレクトリに対するLogを実行
	Public Sub Log_CurrentDir()
		ExecutePathCmd("log", IDEPath.GetCurrentDocDir(DTE))
	End Sub

	Protected MustOverride Function GetTPName() As String

	Private Function CreateCmd(cmd As String, path As String) As String
		Dim tp As String = GetTPName()
		If tp Is Nothing Then Return Nothing

		Dim sc As String = " /command:" + cmd
		Dim sp As String = " /path:" + """" + path + """"
		Return tp + sc + sp + " /notempfile"
	End Function

	Private Sub ExecutePathCmd(cmd As String, path As String)
		If path Is Nothing Then Return

		Dim cmdline As String = CreateCmd(cmd, path)
		If cmdline Is Nothing Then Return

		DTE.ExecuteCommand("Tools.Shell", cmdline)
	End Sub
End Class
