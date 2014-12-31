Imports System
Imports EnvDTE
Imports EnvDTE80
Imports System.Diagnostics


'ランチャ風のコードリスト
Public Module FavoriteList

    ' 指定したパスをシェルから起動する。
	Private Sub LaunchCmd(DTE As DTE2, ByVal cmd As String)
		If System.IO.File.Exists(cmd) = False Then
			If System.IO.Directory.Exists(cmd) = False Then
				MsgBox("指定されたパスにオブジェクトが存在しません : " + vbNewLine + vbNewLine + cmd)
				Return
			End If
		End If

		cmd = """" + cmd + """"
		DTE.ExecuteCommand("Tools.Shell", cmd)
	End Sub


    ' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' よく使うディレクトリなど
    Private Const VisualStudioDocs As String = "C:\Documents and Settings\fukushima\My Documents\VisualStudioDocuments\"


    ' ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' 実行

	Public Sub VisualStudioDocuments(DTE As DTE2)
		LaunchCmd(DTE, VisualStudioDocs)
	End Sub

	Public Sub VisualStudioRegularExpression(DTE As DTE2)
		LaunchCmd(DTE, VisualStudioDocs + "VisualStudio用正規表現.pdf")
	End Sub

	Public Sub VisuaStudioLocalShortcut(DTE As DTE2)
		LaunchCmd(DTE, VisualStudioDocs + "vs_shortcutlist.html")
	End Sub

End Module
