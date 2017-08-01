'   Upon selection of a range of cells, input your desired row separation number.

Sub Select_Every_Nth_Separated_Row()
    Application.ScreenUpdating = False
    Dim C As Range, u As Range
    Dim trig As Double
    Dim iTxt As Variant
    '-------------------
    iTxt = InputBox("For every separation : ", Default:=1) + 1
    '-------------------
    trig = 0
    For Each r In Selection.Rows
        trig = IIf(trig = 0, r.Row, trig)
        If ((r.Row - trig) Mod iTxt) = 0 Then
            If (u Is Nothing) Then
                Set u = r
            Else
                Set u = Union(u, r)
            End If
        End If
    Next
    u.Select
    '-------------------
End Sub
