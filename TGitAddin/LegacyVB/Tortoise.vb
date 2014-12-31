Imports System
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE100
Imports System.Diagnostics


Public Class Tortoise
	Private DTE As DTE2 = Nothing
	Private tp_path As String = """TortoiseProc.exe"""

	Sub New(parent As DTE2)
		DTE = parent
	End Sub

	Private Function CreateCmd(cmd As String, path As String) As String
		Dim sc As String = " /command:" + cmd
		Dim sp As String = " /path:" + """" + path + """"
		Return tp_path + sc + sp + " /notempfile"
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
	Sub SVN_Diff()
		ExecutePathCmd("diff", GetCurrentFilePath())
	End Sub

	' Logウィンドウを表示
	Sub SVN_Log()
		ExecutePathCmd("log", GetCurrentFilePath())
	End Sub

	' カレントディレクトリに対するDiffを実行
	Sub SVN_Diff_CurrentDir()
		ExecutePathCmd("diff", GetCurrentDirPath())
	End Sub

	' カレントディレクトリに対するLogを実行
	Sub SVN_Log_CurrentDir()
		ExecutePathCmd("log", GetCurrentDirPath())
	End Sub
End Class
