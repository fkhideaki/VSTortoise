Imports System
Imports EnvDTE
Imports EnvDTE80
Imports System.Diagnostics

Public Module utilities


    '現在のドキュメントのフルパスを返す
    Function get_current_doc(DTE As DTE2, ByRef filename As String) As Boolean
        filename = DTE.ActiveDocument.FullName
        If Not System.IO.File.Exists(filename) Then Return False
        Return True
    End Function

    '現在のドキュメントが保存されたディレクトリのフルパスを返す
    Function get_current_doc_directory(DTE As DTE2, ByRef dirname As String) As Boolean
        Dim filename As String = ""
        If Not get_current_doc(DTE, filename) Then Return False
        dirname = System.IO.Path.GetDirectoryName(filename)
        Return True
    End Function


    ' 現在の日時をあらわす文字列を生成する。ファイル名に挿入可能。
    Function GetTimestampString() As String
        Dim s As String = ""
        s += System.DateTime.Now.Year.ToString("0000")
        s += System.DateTime.Now.Month.ToString("00")
        s += System.DateTime.Now.Day.ToString("00")
        s += "_"
        s += System.DateTime.Now.Hour.ToString("00")
        s += System.DateTime.Now.Minute.ToString("00")
        s += System.DateTime.Now.Second.ToString("00")
        Return s
    End Function


    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
    ' -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

    ' マクロ内でクリップボードを操作するための機能
    Class AsyncClip
        Public Sub Paste(ByVal s As String)
            m_str = s
            Dim t As New System.Threading.Thread(AddressOf core_func_paste)
            t.ApartmentState = System.Threading.ApartmentState.STA
            t.Start()
            t.Join()
        End Sub

        Public Function GetString() As String
            m_str = ""
            Dim t As New System.Threading.Thread(AddressOf core_func_get)
            t.ApartmentState = System.Threading.ApartmentState.STA
            t.Start()
            t.Join()
            Return m_str
        End Function

        Dim m_str As String

        Private Sub core_func_paste()
            System.Windows.Forms.Clipboard.SetDataObject(m_str, True)
        End Sub

        Private Sub core_func_get()
            Try
                Dim iData As System.Windows.Forms.IDataObject = System.Windows.Forms.Clipboard.GetDataObject()
                If iData.GetDataPresent(System.Windows.Forms.DataFormats.Text) Then
                    m_str = CType(iData.GetData(System.Windows.Forms.DataFormats.Text), String)
                End If
            Catch ex As System.Exception
                m_str = ""
            End Try
        End Sub
    End Class

End Module
