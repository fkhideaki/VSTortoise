Imports System
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE100
Imports System.Diagnostics


Public Class SVN

	Private Enum SVN_Action
		log
		diff
		nodef
	End Enum

	Private DTE As DTE2 = Nothing

	Sub New(parent As DTE2)
		DTE = parent
	End Sub


	Private Class SVN_RequestForm_Class
		WithEvents chk_diff As New System.Windows.Forms.RadioButton
		WithEvents chk_log As New System.Windows.Forms.RadioButton
		WithEvents f As New System.Windows.Forms.Form

		Sub check(ByVal s As Object, ByVal e As EventArgs) Handles chk_diff.CheckedChanged, chk_log.CheckedChanged
			f.Close()
		End Sub

		Sub f_key(ByVal s As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles f.KeyDown
			If e.KeyCode = System.Windows.Forms.Keys.Escape Then f.Close()
		End Sub

		Function get_request() As SVN_Action
			Dim max_x As Integer = 0
			Dim last_y As Integer = 0

			chk_diff.Parent = f
			chk_diff.Text = "&DIFF"
			chk_diff.Top = last_y
			last_y = chk_diff.Bottom
			max_x = System.Math.Max(max_x, chk_diff.Right)

			chk_log.Parent = f
			chk_log.Text = "&LOG"
			chk_log.Top = last_y
			last_y = chk_log.Bottom
			max_x = System.Math.Max(max_x, chk_log.Right)

			f.KeyPreview = True
			f.WindowState = System.Windows.Forms.FormWindowState.Normal
			f.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
			f.Width = f.Width - f.ClientSize.Width + max_x
			f.Height = f.Height - f.ClientSize.Height + last_y
			f.ShowDialog()

			If (chk_diff.Checked) Then Return SVN_Action.diff
			If (chk_log.Checked) Then Return SVN_Action.log
			Return SVN_Action.nodef
		End Function
	End Class

	Const tp_path = """TortoiseProc.exe"""

	' Diffウィンドウを表示
	Sub SVN_Diff()
		Dim filename As String = ""
		If Not get_current_doc(DTE, filename) Then Return

		Dim cmdline As String = tp_path + " /command:diff /path:" + """" + filename + """" + " /notempfile"
		DTE.ExecuteCommand("Tools.Shell", cmdline)
	End Sub

	' カレントディレクトリに対するDiffを実行
	Sub SVN_Diff_CurrentDir()
		Dim dirname As String = ""
		If Not get_current_doc_directory(DTE, dirname) Then Return

		Dim cmdline As String = tp_path + " /command:diff /path:" + """" + dirname + """" + " /notempfile"
		DTE.ExecuteCommand("Tools.Shell", cmdline)
	End Sub

	' Logウィンドウを表示
	Sub SVN_Log()
		Dim filename As String = ""
		If Not get_current_doc(DTE, filename) Then Return

		Dim cmdline As String = tp_path + " /command:log /path:" + """" + filename + """" + " /notempfile"
		DTE.ExecuteCommand("Tools.Shell", cmdline)
	End Sub

	' カレントディレクトリに対するLogを実行
	Sub SVN_Log_CurrentDir()
		Dim dirname As String = ""
		If Not get_current_doc_directory(DTE, dirname) Then Return

		Dim cmdline As String = tp_path + " /command:log /path:" + """" + dirname + """" + " /notempfile"
		DTE.ExecuteCommand("Tools.Shell", cmdline)
	End Sub

	Private Function SVN_RequestForm() As SVN_Action
		Dim f As New SVN_RequestForm_Class
		Return f.get_request
	End Function

	Sub SVN_Req()
		Select Case SVN_RequestForm()
			Case SVN_Action.diff
				SVN_Diff()
			Case SVN_Action.log
				SVN_Log()
		End Select
	End Sub

End Class