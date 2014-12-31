Imports System
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE100
Imports System.Diagnostics

Public Module ShortcutViewer

    '*** VSコマンド諸元格納用
    Structure VSCommandMemDef
        Public Guid As String
        Public ID As Integer
        Public Name As String
        Public LocalizedName As String
        Public Bindings As String
    End Structure


    '*** XHTMLファイル 出力先フルパス
    Function GetXTHMLPath() As String
        Dim sfd As New System.Windows.Forms.SaveFileDialog
        sfd.Filter = "*.html|*.html"
        If sfd.ShowDialog() <> System.Windows.Forms.DialogResult.OK Then Return ""

        Return sfd.FileName
    End Function


    '*** VSコマンド一覧表の出力 **************************************
    ' 1. [Guid、ID、Name、Localname、Bindings → XHTML]をVS IDEに表示
    '*****************************************************************
	Public Sub GetCommandList_XHTML(DTE As DTE2)

		'***  [Guid、ID、Name、Localname、Bindings]を取得
		'   ※ 取得した時点で[使用する場所]-[コマンド名(英名)]順にソートされている
		Dim VSCommandList As Collections.SortedList = fGetCommandListSort(DTE)

		'*** 取得したデータをXHTMLファイルとして出力
		Call OutputCommandsAsHTML(DTE, VSCommandList)

	End Sub


    '******************************************
    ' VSコマンドとショートカットキー情報の取得
    '******************************************
	Private Function fGetCommandListSort(DTE As DTE2) As Collections.SortedList

		'*** VSコマンド情報(返値用ワーク)
		Dim sl As New Collections.SortedList

		'*** VSコマンドのCSV出力
		For Each cmd As Command In DTE.Commands
			'NameまたはLocalizedNameが空でなければ取得対象
			If cmd.Name <> "" AndAlso cmd.LocalizedName <> "" Then
				'ショートカットが割り当たっていたら取得対象
				If Not cmd.Bindings Is Nothing _
					AndAlso cmd.Bindings.GetLength(0) > 0 Then
					'VSコマンド情報を取得
					Dim VSCommandMem As VSCommandMemDef
					VSCommandMem.Guid = cmd.Guid
					VSCommandMem.ID = cmd.ID
					VSCommandMem.Name = cmd.Name
					VSCommandMem.LocalizedName = cmd.LocalizedName
					For i As Integer = 0 To cmd.Bindings.GetLength(0) - 1
						VSCommandMem.Bindings = cmd.Bindings(i)
						'キー文字列の生成(ここを変更すると任意の順にソートできる)
						Dim vd As String = VSCommandMem.Bindings	'(記述コードを短くするため)
						Dim slKey As String = vd.Substring(0, vd.IndexOf("::")) _
											& New String(" ", 50) & VSCommandMem.Name _
											& vd.Substring(vd.IndexOf("::") + 2)
						'ソートコレクションへの追加
						sl.Add(slKey, VSCommandMem)
					Next
				Else
					'ショートカットキーが割り当てられていないVSコマンドは取得しない
				End If
			End If
		Next

		Return sl

	End Function

    '*************************************************************
    ' VSコマンドとショートカットキー情報→XHTMLファイルとして出力
    '*************************************************************
	Private Sub OutputCommandsAsHTML(DTE As DTE2, ByVal sl As Collections.SortedList)

		'*** 文字定数(記述コードを短くするため)
		Const QT As String = CStr(ControlChars.Quote)
		Const TB As String = CStr(ControlChars.Tab)

		'***[使用する場所]Breakトリガ
		Dim BreakTrriger As String = ""

		'*** XHTMLファイル出力用ストリームライタ
		Dim xhtmlW As IO.StreamWriter = Nothing

		'*** VSコマンドのXHTML出力
		Try
			Dim out_file_path As String = GetXTHMLPath()
			If out_file_path = "" Then Return
			xhtmlW = New IO.StreamWriter(out_file_path, False, System.Text.Encoding.GetEncoding("utf-8"))

			'XHTMLヘッダ
			xhtmlW.WriteLine("<?xml version=" & QT & "1.0" & QT & " encoding=" & QT & "UTF-8" & QT & " ?>")
			xhtmlW.WriteLine("<!DOCTYPE html PUBLIC " & QT & "-//W3C//DTD XHTML 1.0 Strict//EN" & QT & " " & QT & "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd" & QT & ">")
			xhtmlW.WriteLine("<html xmlns=" & QT & "http://www.w3.org/1999/xhtml" & QT & " xml:lang=" & QT & "ja" & QT & " lang=" & QT & "ja" & QT & ">")
			xhtmlW.WriteLine("")
			xhtmlW.WriteLine("<head id=ctl00_Head1>")
			xhtmlW.WriteLine(TB & "<meta http-equiv=Content-Type content=" & QT & "text/html; charset=utf-8" & QT & ">")
			xhtmlW.WriteLine(TB & "<title>現在のショートカット キー (現在の開発環境)</title>")
			xhtmlW.WriteLine(TB & "<style type=text/css>")
			xhtmlW.WriteLine(TB & TB & "p {padding-right: 0px; padding-left: 0px; padding-bottom: 0px; margin: 0px 0px 10px; padding-top: 0px}")
			xhtmlW.WriteLine(TB & TB & ".RightPanel {font-size: 70%; vertical-align: top; font-family: Verdana, Arial, Helvetica, sans-serif}")
			xhtmlW.WriteLine(TB & TB & ".title {margin: 0px 0px 10px; font: 190% Arial, Helvetica, sans-serif; color: #000000}")
			xhtmlW.WriteLine(TB & TB & ".ContentArea .topic table td p {padding-right: 5px; padding-left: 5px; padding-bottom: 5px; margin: 0px; padding-top: 5px}")
			xhtmlW.WriteLine(TB & TB & "div.ContentArea table th {padding-right: 5px; padding-left: 5px; font-size: 70%; padding-bottom: 5px; padding-top: 5px; text-align: left; background: #cccccc; vertical-align: bottom}")
			xhtmlW.WriteLine(TB & TB & "div.ContentArea table th p {font-weight: bold}")
			xhtmlW.WriteLine(TB & TB & "div.ContentArea {margin: 20px; line-height: 140%}")
			xhtmlW.WriteLine(TB & TB & "div.ContentArea table {margin: 10px 0px; width: auto; border-collapse: collapse}")
			xhtmlW.WriteLine(TB & TB & "div.ContentArea table td {border-right: #cccccc 0px solid; border-top: #cccccc 0px solid; background: #ffffff; vertical-align: top; border-left: #cccccc 0px solid; border-bottom: #cccccc 0px solid; padding-right: 5px; padding-left: 5px; font-size: 70%; padding-bottom: 5px; padding-top: 5px; text-align: left}")
			xhtmlW.WriteLine(TB & "	</style>")
			xhtmlW.WriteLine("</head>")
			xhtmlW.WriteLine("")
			xhtmlW.WriteLine("<body id=ctl00_MTPS_Body>")
			xhtmlW.WriteLine(TB & "<div class=RightPanel id=ctl00_LibFrame_MtpsContentPanel>")
			xhtmlW.WriteLine(TB & TB & "<div class=ContentArea>")
			xhtmlW.WriteLine(TB & TB & TB & "<div class=topic>")

			'コマンド明細
			For i As Integer = 0 To sl.Count - 1
				'1コマンド分の情報を取得
				Dim slBf As VSCommandMemDef = sl.GetByIndex(i)
				'「使用する場所」の切り出し
				Dim VSCommandClass As String = slBf.Bindings.Substring(0, slBf.Bindings.IndexOf("::"))
				If VSCommandClass <> BreakTrriger Then
					'初回でなければ、現在の<table>の終了
					If BreakTrriger <> "" Then
						xhtmlW.WriteLine(TB & TB & TB & TB & TB & "</tbody>")
						xhtmlW.WriteLine(TB & TB & TB & TB & "</table>")
						xhtmlW.WriteLine(TB & TB & TB & TB & "<hr />")
					End If
					'次の<table>の開始
					xhtmlW.WriteLine(TB & TB & TB & TB & "<div class=majorTitle>Visual Studio " & DTE.Version & " " & DTE.Edition & "</div>")
					xhtmlW.WriteLine(TB & TB & TB & TB & "<div class=title>[" & VSCommandClass & "] ショートカット キー (現在の開発環境)</div>")
					xhtmlW.WriteLine(TB & TB & TB & TB & "<p></p>")
					xhtmlW.WriteLine(TB & TB & TB & TB & "<table cellspacing=2 cellpadding=5 width=" & QT & "100%" & QT & ">")
					xhtmlW.WriteLine(TB & TB & TB & TB & TB & "<tbody>")
					xhtmlW.WriteLine(TB & TB & TB & TB & TB & "<tr><th>コマンド名</th><th>コマンド名(日本語)</th><th>ショートカット キー</th></tr>")
					BreakTrriger = VSCommandClass
				End If
				'ショートカットキーの切り出し
				Dim sKey As String = slBf.Bindings.Substring(slBf.Bindings.IndexOf("::") + 2)
				Dim strBf As String() = {VSCommandClass, slBf.Name, slBf.LocalizedName, sKey}
				xhtmlW.WriteLine("<tr><td><p>" & slBf.Name & "</p></td><td><p>" & slBf.LocalizedName & "</p></td><td><p>" & sKey & "</p></td></tr>")
			Next

			'XHTMLフッタ
			xhtmlW.WriteLine(TB & TB & TB & TB & TB & "</tbody>")
			xhtmlW.WriteLine(TB & TB & TB & TB & "</table>")
			xhtmlW.WriteLine(TB & TB & TB & "</div>")
			xhtmlW.WriteLine(TB & TB & "</div>")
			xhtmlW.WriteLine(TB & "</div>")
			xhtmlW.WriteLine("</body>")
			xhtmlW.WriteLine("")
			xhtmlW.WriteLine("</html>")

		Finally
			If Not xhtmlW Is Nothing Then
				xhtmlW.Close()
			End If
		End Try

	End Sub

End Module
