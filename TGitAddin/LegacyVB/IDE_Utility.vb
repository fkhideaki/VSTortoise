Imports System
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE100
Imports System.Diagnostics


Public Class IDE_Utility

	Dim DTE As DTE2

	Public Sub New(dte2_instance As DTE2)
		DTE = dte2_instance
	End Sub

	' 選択範囲の外側を非表示にする
	Sub ShowOnlySelected()
		Dim ts As TextSelection = DTE.ActiveDocument.Selection
		Dim CurrentLine As String = ts.ActivePoint.Line.ToString()

		Dim prev_top As Integer = ts.TopPoint.Line
		Dim prev_bot As Integer = ts.BottomPoint.Line

		ts.GotoLine(prev_bot + 1, False)
		ts.EndOfDocument(True)
		ts.OutlineSection()

		ts.GotoLine(prev_top - 1, False)
		ts.StartOfDocument(True)
		ts.OutlineSection()
	End Sub


	Function GetActivePane()
		Dim ActWindow = DTE.ActiveDocument.ActiveWindow
		Return ActWindow.Object.ActivePane
	End Function

	Public Sub SelectNextBrace()
		Dim ActPane = GetActivePane()

		Dim Selection = DTE.ActiveDocument.Selection
		Dim Selstr As String = Selection.Text

		Dim BraceSelecting As Boolean = Selstr.StartsWith("{")

		If BraceSelecting Then
			DTE.ActiveDocument.Selection.CharLeft()
			DTE.ActiveDocument.Selection.CharRight()
		End If

		ActPane.IncrementalSearch.StartForward()
		ActPane.IncrementalSearch.AppendCharAndSearch(AscW("{"))

		ActPane.IncrementalSearch.Exit()
		DTE.ExecuteCommand("Edit.GotoBraceExtend")
	End Sub

	Public Sub SelectPrevBrace()
		Dim ActPane = GetActivePane()

		Dim Selection = DTE.ActiveDocument.Selection
		Dim Selstr As String = Selection.Text

		Dim BraceSelecting As Boolean = Selstr.EndsWith("}")

		If BraceSelecting Then
			DTE.ActiveDocument.Selection.CharRight()
			DTE.ActiveDocument.Selection.CharLeft(False, 2)
		End If

		ActPane.IncrementalSearch.StartBackward()
		ActPane.IncrementalSearch.AppendCharAndSearch(AscW("}"))
		DTE.ActiveDocument.Selection.CharLeft()

		ActPane.IncrementalSearch.Exit()
		DTE.ExecuteCommand("Edit.GotoBraceExtend")
	End Sub

	Public Sub CloseNextBrace()
		SelectNextBrace()
		DTE.ActiveDocument.Selection.OutlineSection()
	End Sub

	Public Sub ClosePrevBrace()
		SelectPrevBrace()
		DTE.ActiveDocument.Selection.OutlineSection()
	End Sub


	Public Sub SelectNextParenthesis()
		Dim ActPane = GetActivePane()

		Dim Selection = DTE.ActiveDocument.Selection
		Dim Selstr As String = Selection.Text

		Dim BraceSelecting As Boolean = Selstr.StartsWith("(")

		If BraceSelecting Then
			DTE.ActiveDocument.Selection.CharLeft()
			DTE.ActiveDocument.Selection.CharRight()
		End If

		ActPane.IncrementalSearch.StartForward()
		ActPane.IncrementalSearch.AppendCharAndSearch(AscW("("))

		ActPane.IncrementalSearch.Exit()
		DTE.ExecuteCommand("Edit.GotoBraceExtend")
	End Sub

	Public Sub SelectPrevParenthesis()
		Dim ActPane = GetActivePane()

		Dim Selection = DTE.ActiveDocument.Selection
		Dim Selstr As String = Selection.Text

		Dim BraceSelecting As Boolean = Selstr.EndsWith(")")

		If BraceSelecting Then
			DTE.ActiveDocument.Selection.CharRight()
			DTE.ActiveDocument.Selection.CharLeft(False, 2)
		End If

		ActPane.IncrementalSearch.StartBackward()
		ActPane.IncrementalSearch.AppendCharAndSearch(AscW(")"))
		DTE.ActiveDocument.Selection.CharLeft()

		ActPane.IncrementalSearch.Exit()
		DTE.ExecuteCommand("Edit.GotoBraceExtend")
	End Sub



	Private Function RemoveEscapeChar_CPP(ByVal s As String) As String
		Return Replace(s, "\", "\\")
	End Function

	' 現在選択中の範囲のエスケープ文字を文字列内の形式に変換する
	Public Sub ConvertEscapeChatToInsideString()
		Dim s_src As String = DTE.ActiveDocument.Selection.Text
		Dim s_dst As String = RemoveEscapeChar_CPP(s_src)
		DTE.ActiveDocument.Selection.Text = s_dst
	End Sub

	' クリップボード内の文字列のエスケープ文字を変換して貼り付け
	Public Sub PasteAndConvertEscapeCharToInsideString()
		Dim ac As New AsyncClip
		Dim s As String = ac.GetString()

		Dim s_dst As String = RemoveEscapeChar_CPP(s)
		DTE.ActiveDocument.Selection.Text = s_dst
	End Sub


	' 目盛りを生成する
	Public Sub MakeScaler()
		DTE.ActiveDocument.Selection.EndOfLine()
		DTE.ActiveDocument.Selection.StartOfLine(vsStartOfLineOptions.vsStartOfLineOptionsFirstColumn)
		DTE.ActiveDocument.Selection.Text = "//--+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8----+----9----+----100"
		DTE.ActiveDocument.Selection.NewLine()
	End Sub

	' 分割線を生成する
	Public Sub MakeSplitter()
		DTE.ActiveDocument.Selection.Text = "// ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
		DTE.ActiveDocument.Selection.NewLine()
	End Sub

	' Allman形式の括弧を生成する
	Sub CreateAllmanBrace()
		DTE.ActiveDocument.Selection.EndOfLine()
		DTE.ActiveDocument.Selection.NewLine()
		DTE.ActiveDocument.Selection.Text = "{"
		DTE.ActiveDocument.Selection.NewLine(2)
		DTE.ActiveDocument.Selection.Text = "}"
		DTE.ActiveDocument.Selection.LineUp()
		DTE.ActiveDocument.Selection.Indent()
	End Sub


	' タスクペイン切り替え : ユーザタスク
	Public Sub TaskPaneSwitch_UserTask()
		DTE.ExecuteCommand("Categories", "ユーザー タスク")
	End Sub

	' タスクペイン切り替え : コメント
	Public Sub TaskPaneSwitch_Comment()
		DTE.ExecuteCommand("Categories", "コメント")
	End Sub

	' イミディエイトに指定した変数を出力する
	Public Sub OutImmediateValue(ByVal valuename As String)
		DTE.ExecuteCommand("Debug.Immediate")
		DTE.ExecuteCommand("Debug.EvaluateStatement", valuename)
	End Sub



	' 現在のソリューションディレクトリを表示する
	Sub ShowCurrentSolutionDir()
		Dim slnfile As String = System.IO.Path.GetFullPath(DTE.Solution.FullName)
		Dim dirpath As String = System.IO.Path.GetDirectoryName(slnfile)

		dirpath = """" + dirpath + """"

		DTE.ExecuteCommand("Tools.Shell", dirpath)
	End Sub

	' 現在のアウトプットディレクトリを表示する
	Sub ShowCurrentOutputDir()
		'Dim outdir As String = ""
		'Dim dirpath As String = System.IO.Path.GetDirectoryName(outdir)

		'For Each d As EnvDTE.BuildDependency In DTE.Solution.SolutionBuild.BuildDependencies

		'Next


		'dirpath = """" + dirpath + """"

		'DTE.ExecuteCommand("Tools.Shell", dirpath)
	End Sub

	' 現在のドキュメントが保存されたディレクトリを表示する
	Sub ShowCurrentDocumentDir()
		Dim filename As String = ""
		If Not get_current_doc(DTE, filename) Then Return
		Dim dir_path As String = System.IO.Path.GetDirectoryName(filename)
		dir_path = """" + dir_path + """"
		DTE.ExecuteCommand("Tools.Shell", dir_path)
	End Sub

	' ワードラップをスイッチする
	Sub SwitchWordWrap()
		DTE.ExecuteCommand("Edit.ToggleWordWrap")
	End Sub

	' デバッグ出力ウィンドウをクリア
	Sub ClearDebugOutputWindow()
		DTE.ExecuteCommand("Edit.ClearOutputWindow")
	End Sub


	' ソリューションの情報をメッセージボックスで表示
	Sub ShowCurrentSolutionInformation()
		Dim slnfile As String = System.IO.Path.GetFullPath(DTE.Solution.FullName)
		Dim dirpath As String = System.IO.Path.GetDirectoryName(slnfile)
		Dim filename As String = System.IO.Path.GetFileName(slnfile)

		Dim subs() As String = dirpath.Split("\")

		Dim msg As String = ""
		msg += "------------------------------------------------------------------" + vbNewLine
		For i As Integer = 0 To subs.Length - 1
			If i <> 0 Then msg += " ↑ " + vbNewLine
			msg += subs(subs.Length - i - 1) + vbNewLine
		Next
		msg += "------------------------------------------------------------------" + vbNewLine
		msg += vbNewLine
		msg += " Dir : " + dirpath + vbNewLine
		msg += " SlnFile : " + filename + vbNewLine
		msg += vbNewLine

		MsgBox(msg)
	End Sub



	'' OneFunctionView は、VS コード モデルおよびエディタのオートメーション モデルを使って
	'' 関数の前後にアウトライン セクションを作成するため、現在カレットを含んでいる
	'' 関数のみが表示されます。
	''
	Sub OneFunctionView()
		Dim textSelection As EnvDTE.TextSelection
		Dim textSelectionPointSaved As EnvDTE.EditPoint
		Dim editPoint As EnvDTE.EditPoint

		textSelection = DTE.ActiveWindow.Selection
		textSelectionPointSaved = textSelection.ActivePoint.CreateEditPoint
		editPoint = textSelection.ActivePoint.CreateEditPoint

		'' ドキュメントを開始するポイントおよびアウトラインを取得します。
		editPoint.MoveToPoint(editPoint.CodeElement(EnvDTE.vsCMElement.vsCMElementFunction).GetStartPoint(vsCMPart.vsCMPartWholeWithAttributes))
		editPoint.LineUp()
		textSelection.MoveToPoint(editPoint, False)
		textSelection.StartOfDocument(True)
		textSelection.OutlineSection()
		'editPoint.LineDown()
		'' editPoint を関数に戻して、ドキュメントの終わりまでアウトラインを指定します。
		editPoint.MoveToPoint(textSelectionPointSaved)
		editPoint.MoveToPoint(editPoint.CodeElement(EnvDTE.vsCMElement.vsCMElementFunction).GetEndPoint(vsCMPart.vsCMPartWholeWithAttributes))
		editPoint.LineDown()
		textSelection.MoveToPoint(editPoint, False)
		textSelection.EndOfDocument(True)
		textSelection.OutlineSection()
		textSelection.MoveToPoint(textSelectionPointSaved)
	End Sub



	' 現在の選択領域から指定した文字列を含む行を抽出する
	' 検索結果は現在のドキュメントの最下行以降に追加される
	Sub FindString()
		Dim str_s As String = DTE.ActiveDocument.Selection.Text
		Dim lines() As String = str_s.Replace(vbCrLf, vbLf).Split(vbLf)

		Dim find_str As String = InputBox("search text")
		If find_str = "" Then Return

		DTE.ActiveDocument.Selection.EndOfDocument()
		DTE.ActiveDocument.Selection.Text = vbNewLine + vbNewLine + " -- find reslut --" + vbNewLine + vbNewLine

		Dim i As Integer = 0
		For Each s As String In lines
			i += 1
			If InStr(s, find_str) <> 0 Then
				DTE.ActiveDocument.Selection.Text += i.ToString("00000") + " : " + s + vbNewLine
			End If
		Next
	End Sub

	Sub MakeBackup(DTE As DTE2)
		Dim filename As String = ""
		If Not get_current_doc(DTE, filename) Then Return

		filename = System.IO.Path.GetFullPath(filename)
		Dim filename_base As String = System.IO.Path.GetDirectoryName(filename) + "\" + System.IO.Path.GetFileNameWithoutExtension(filename)
		Dim file_ext As String = System.IO.Path.GetExtension(filename)

		Dim datestr As String = GetTimestampString()

		Dim i As Integer = 1
		Dim outfilename As String = filename_base + "_" + datestr + file_ext

		If MsgBox("バックアップ作成 : " + vbNewLine + vbNewLine + outfilename, MsgBoxStyle.YesNo) = MsgBoxResult.No Then Return

		If System.IO.File.Exists(outfilename) Or System.IO.Directory.Exists(outfilename) Then
			MsgBox("error")
			Return
		End If

		Try
			System.IO.File.Copy(filename, outfilename)
		Catch ex As System.Exception
			MsgBox("error : " + vbNewLine + vbNewLine + ex.Message)
		End Try
	End Sub


	' 作業領域用ファイルを開く
	Sub OpenFile(ByVal file_path As String)
		If System.IO.File.Exists(file_path) = False Then Return
		DTE.ItemOperations.OpenFile(file_path)
	End Sub

	Sub OpenTempEX(idx As Integer)
		OpenFile("C:\ws\vs\2005\tmp_" + idx.ToString() + ".cpp")
	End Sub

	Sub OpenTempEX_0()
		OpenTempEX(0)
	End Sub


	' c++用。現在のファイルと対になる.h,.cppを開く
	Sub OpenPair_cpp_h()
		Dim filename As String = ""
		If Not get_current_doc(DTE, filename) Then Return

		Dim src_ext As String = System.IO.Path.GetExtension(filename).ToLower()
		Dim dst_ext As String

		If (src_ext = ".cpp") Then
			dst_ext = ".h"
		ElseIf (src_ext = ".h") Then
			dst_ext = ".cpp"
		ElseIf (src_ext = ".hpp") Then
			dst_ext = ".cpp"
		Else
			Return
		End If

		Dim base_path As String = ""
		base_path += System.IO.Path.GetDirectoryName(filename)
		base_path += "\"
		base_path += System.IO.Path.GetFileNameWithoutExtension(filename)

		Dim pair_path As String = base_path + dst_ext

		If System.IO.File.Exists(pair_path) Then
			DTE.ItemOperations.OpenFile(pair_path)
		Else
			If dst_ext = ".h" Then
				dst_ext = ".hpp"
			ElseIf dst_ext = ".cpp" Then
				dst_ext = ".c"
			Else
				Return
			End If

			pair_path = base_path + dst_ext

			If System.IO.File.Exists(pair_path) Then
				DTE.ItemOperations.OpenFile(pair_path)
			End If
		End If
	End Sub


	' 出力ウィンドウの指定した名前を持つペインを開く。
	' 存在しないペイン名が指定された場合は、自動的に生成される。
	Public Function GetOutputWindowPane(ByVal Name As String, Optional ByVal show As Boolean = True) As OutputWindowPane
		Dim win As Window = DTE.Windows.Item(EnvDTE.Constants.vsWindowKindOutput)
		If show Then win.Visible = True
		Dim ow As OutputWindow = win.Object
		Dim owpane As OutputWindowPane
		Try
			owpane = ow.OutputWindowPanes.Item(Name)
		Catch e As System.Exception
			owpane = ow.OutputWindowPanes.Add(Name)
		End Try
		owpane.Activate()
		Return owpane
	End Function


	' テキストエディタのフォントサイズを変更する
	Const MAX_TEXTEDITOR_FONT_SIZE As Integer = 15
	Const MIN_TEXTEDITOR_FONT_SIZE As Integer = 2

	Sub InclTextEditorFont()
		AddTextEditorFontSize(1)
		DTE.ExecuteCommand("Window.ActivateDocumentWindow")
	End Sub

	Sub DeclTextEditorFont()
		AddTextEditorFontSize(-1)
		DTE.ExecuteCommand("Window.ActivateDocumentWindow")
	End Sub

	Sub AddTextEditorFontSize(ByVal add_size As Integer)
		Dim prop_container As EnvDTE.Properties = DTE.Properties("FontsAndColors", "TextEditor")
		Dim prop As EnvDTE.Property = prop_container.Item("FontSize")

		Dim current_font_size As Integer = CInt(prop.Value)
		Dim font_size As Integer = current_font_size
		font_size += add_size
		font_size = System.Math.Min(MAX_TEXTEDITOR_FONT_SIZE, font_size)
		font_size = System.Math.Max(MIN_TEXTEDITOR_FONT_SIZE, font_size)
		If (font_size <> current_font_size) Then prop.Value = font_size

		Dim out_msg As String = current_font_size.ToString + " -> " + font_size.ToString
		GetOutputWindowPane("font size").OutputString(out_msg + vbNewLine)
		Return
	End Sub


	' C/C++ のインデントスタイルの切り替え
	Sub SwitchIndentStyle()
		Dim s_doctype As String = GetCurrentDocPropertyPageName(DTE)

		Dim props_CPPEditor As EnvDTE.Properties = DTE.Properties("TextEditor", s_doctype)
		Dim prop_tabstyle As EnvDTE.Property = props_CPPEditor.Item("InsertTabs")

		prop_tabstyle.Value = Not prop_tabstyle.Value

		' 出力ウィンドウにタブの状態を出力
		Dim out_msg As String = "InsertTabs : " + prop_tabstyle.Value.ToString
		GetOutputWindowPane(s_doctype + " setting").OutputString(out_msg + vbNewLine)
	End Sub


	' タブサイズの変更
	Sub InclTabSize_CPP()
		AddTabSize_CPP(2)
	End Sub

	Sub DeclTabSize_CPP()
		AddTabSize_CPP(-2)
	End Sub

	Sub AddTabSize_CPP(ByVal add_size As Integer)
		Dim s_doctype As String = GetCurrentDocPropertyPageName(DTE)

		Dim props_CPPEditor As EnvDTE.Properties = DTE.Properties("TextEditor", s_doctype)
		Dim prop_tabsize As EnvDTE.Property = props_CPPEditor.Item("TabSize")

		Dim prev_size = prop_tabsize.Value
		Dim i As Integer = prev_size + add_size

		i = System.Math.Min(10, i)
		i = System.Math.Max(2, i)
		If (prev_size <> i) Then prop_tabsize.Value = i

		' 出力ウィンドウにタブの状態を出力
		Dim out_msg As String = "TabSize : " + prop_tabsize.Value.ToString
		GetOutputWindowPane(s_doctype + " setting").OutputString(out_msg + vbNewLine)
	End Sub


	' インデントサイズの変更
	Sub InclIndentSize_CPP()
		AddIndentSize_CPP(2)
	End Sub

	Sub DeclIndentSize_CPP()
		AddIndentSize_CPP(-2)
	End Sub

	Sub AddIndentSize_CPP(ByVal add_size As Integer)
		Dim s_doctype As String = GetCurrentDocPropertyPageName(DTE)

		Dim props_CPPEditor As EnvDTE.Properties = DTE.Properties("TextEditor", s_doctype)
		Dim prop_indentsize As EnvDTE.Property = props_CPPEditor.Item("IndentSize")

		Dim prev_size = prop_indentsize.Value
		Dim i As Integer = prev_size + add_size

		i = System.Math.Min(10, i)
		i = System.Math.Max(2, i)
		If (prev_size <> i) Then prop_indentsize.Value = i

		' 出力ウィンドウにタブの状態を出力
		Dim out_msg As String = "IndentSize : " + prop_indentsize.Value.ToString
		GetOutputWindowPane(s_doctype + " setting").OutputString(out_msg + vbNewLine)
	End Sub


	' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

	Sub ClipCurrentFile()
		Dim filename As String = ""
		If Not get_current_doc(DTE, filename) Then Return
		Dim ac As New AsyncClip
		ac.Paste(filename)

		Dim out_msg As String = "clipboard <- " + filename
		GetOutputWindowPane("clipboard").OutputString(out_msg + vbNewLine)
	End Sub

	Sub ClipCurrentDir()
		Dim filename As String = ""
		If Not get_current_doc(DTE, filename) Then Return
		Dim dirname As String = System.IO.Path.GetDirectoryName(filename)
		Dim ac As New AsyncClip
		ac.Paste(dirname)

		Dim out_msg As String = "clipboard <- " + dirname
		GetOutputWindowPane("clipboard").OutputString(out_msg + vbNewLine)
	End Sub

	' クリップボード内の文字列がファイルのパスであればそれを開く
	Sub OpenClip()
		Dim ac As New AsyncClip
		Dim filename As String = ac.GetString()
		If System.IO.File.Exists(filename) = False Then Return
		DTE.ItemOperations.OpenFile(filename)
	End Sub


	' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

	' 仮想スペースを切り替える
	Sub SwitchVirtualSpace()
		Dim s_doctype As String = GetCurrentDocPropertyPageName(DTE)

		Dim props_CPPEditor As EnvDTE.Properties = DTE.Properties("TextEditor", GetCurrentDocPropertyPageName(DTE))
		Dim prop_virtualspace As EnvDTE.Property = props_CPPEditor.Item("VirtualSpace")
		prop_virtualspace.Value = Not prop_virtualspace.Value

		Dim out_msg As String = "VirtualSpace : " + prop_virtualspace.Value.ToString
		GetOutputWindowPane(s_doctype + " setting").OutputString(out_msg + vbNewLine)
	End Sub


	' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

	' 環境パラメータ名の列挙のサンプル
	'Dim props_CPPEditor As EnvDTE.Properties = DTE.Properties("TextEditor", GetCurrentDocPropertyPageName())
	'Dim s As String = ""
	'Dim ii As Integer
	'For Each p As EnvDTE.Property In props_CPPEditor
	'    s += p.Name + vbNewLine
	'Next
	'MsgBox(s)



	' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

	Sub SelectCurrentBlock()
		TopOfBlock()
		DTE.ExecuteCommand("Edit.GotoBraceExtend")
	End Sub

	' カレントブロックの先頭まで移動
	Sub TopOfBlock()
		TopOfBlock_Core(False)
	End Sub

	Sub TopOfBlock_AndGoUpWhenRoot()
		TopOfBlock_Core(True)
	End Sub

	Sub TopOfBlock_Core(ByVal GoUpWhenRoot As Boolean)
		Dim TDoc As TextDocument = DTE.ActiveDocument.Object("TextDocument")
		Dim EP As EditPoint = TDoc.CreateEditPoint(Nothing)
		Dim Sel As TextSelection = TDoc.Selection

		EP.LineDown(Sel.ActivePoint.Line - 1)

		Dim Indent As Long = IndentLevel(EP)

		If Indent <= 1 Then
			If GoUpWhenRoot Then
				DTE.ActiveDocument.Selection.LineUp()
			End If
			Exit Sub
		End If

		Dim CurrentIndent As Long = Indent

		Do While CurrentIndent >= Indent
			CurrentIndent = IndentLevel(EP)
			If CurrentIndent = -1 Then
				Exit Sub
			Else
				If CurrentIndent >= Indent Then
					If EP.Line = 1 Then
						Exit Sub
					Else
						EP.LineUp()
					End If
				End If
			End If
		Loop

		Sel.MoveToLineAndOffset(EP.Line, EP.LineCharOffset)
	End Sub

	Function IndentLevel(ByVal EP As EditPoint)
		Dim LastLine As Long

		Do While True
			LastLine = EP.Line
			EP.StartOfLine()
			SkipSpaceRight(EP)

			If EP.Line <> LastLine Then
				' empty line
				EP.LineUp(2)
			ElseIf EP.LineCharOffset = EP.LineLength + 1 Then
				' line with only spaces and/or tabs
				EP.LineUp()
			Else
				Return EP.DisplayColumn
				Exit Do
			End If

			If EP.Line <> LastLine - 1 Then
				Return -1
			End If
		Loop

		Return -1
	End Function

	Sub SkipSpaceRight(ByVal EP As EditPoint)
		Dim Line As Long
		Dim LastPos As Long

		Line = EP.Line
		'' Don't call IsWhitespace here due to consing overhead.
		Do While EP.Line = Line And (EP.GetText(1) = " " Or EP.GetText(1) = CStr(Microsoft.VisualBasic.ControlChars.Tab))
			LastPos = EP.LineCharOffset
			EP.CharRight()
			If EP.LineCharOffset = LastPos Then
				'' end of document
				Exit Sub
			End If
		Loop
	End Sub


	' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

	' スコープ解決を排除してペーストする
	Public Sub PasteOnlyLastScopeStr()
		Dim ac As New AsyncClip
		Dim s As String = ac.GetString().Clone()

		Dim last_idx As Integer = s.LastIndexOf("::")
		If (last_idx >= 0) Then
			s = s.Substring(last_idx + 2)
		End If

		DTE.ActiveDocument.Selection.Text = s
	End Sub

End Class
