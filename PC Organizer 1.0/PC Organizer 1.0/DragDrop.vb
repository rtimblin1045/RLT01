Imports System.IO

Public Class DragDrop

    Private Function IsString(ByVal obj As Object) As Boolean
        Dim b As Boolean
        Try
            b = Not Integer.Parse(obj, 1) > 0
        Catch ex As Exception
            b = True
        End Try
        Return b
    End Function

    Public Function GetFileName(ByRef strFileName As String, ByVal e As DragEventArgs) As Boolean
        Dim b As Boolean
        strFileName = ""
        If e.AllowedEffect And DragDropEffects.Copy = DragDropEffects.Copy Then
            Dim data As Array = e.Data.GetData("FileName")
            If data IsNot Nothing Then
                If data.Length = 1 AndAlso IsString(data.GetValue(0)) Then
                    strFileName = data(0)
                    Dim ext As String = System.IO.Path.GetExtension(strFileName).ToLower
                    Dim t$ = ext.Replace(".", "")
                    Dim s$ = "jpgpngbmp"
                    If s.Contains(t) Then b = True
                End If
            End If
        End If
        Return b
    End Function

End Class
