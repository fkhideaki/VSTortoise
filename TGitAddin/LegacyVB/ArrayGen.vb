Imports System
Imports EnvDTE
Imports EnvDTE80
Imports System.Diagnostics

Public Module ArrayGen

    Function EvalArraGenerateScript(ByVal script As String) As String()
        Dim vs() As String = script.Split("#")

        If vs.Length <> 2 Then Return Nothing

        Dim StrTemplate As String = vs(0).Trim()
        Dim StrArg As String = vs(1).Trim()

        Dim TemplateAry() As String = StrTemplate.Split("@")
        Dim ArgAry() As String = StrArg.Split(" ")

        If TemplateAry.Length <> ArgAry.Length + 1 Then Return Nothing

        Dim NumLayer As Integer = ArgAry.Length

        Dim LoopAry(ArgAry.Length - 1) As Integer
        For i As Integer = 0 To LoopAry.Length - 1
            Try
                LoopAry(i) = Integer.Parse(ArgAry(i))
            Catch ex As Exception
                Return Nothing
            End Try
        Next

        Dim TotalLines As Integer = 1
        For Each i As Integer In LoopAry
            TotalLines *= i
        Next

        Dim RetAry(TotalLines - 1) As String

        Dim CurrentLayer As Integer = 0
        Dim CurrentLayerLength As Integer = TotalLines
        Do
            CurrentLayerLength \= LoopAry(CurrentLayer)
            For i As Integer = 0 To RetAry.Length - 1
                RetAry(i) += TemplateAry(CurrentLayer)

                Dim idx As Integer = i \ CurrentLayerLength
                idx = idx Mod LoopAry(CurrentLayer)
                RetAry(i) += idx.ToString()

                If CurrentLayer = ArgAry.Length - 1 Then
                    RetAry(i) += TemplateAry(CurrentLayer + 1)
                End If
            Next

            CurrentLayer += 1
            If CurrentLayer = ArgAry.Length Then Exit Do
        Loop

        Return RetAry
    End Function

End Module
