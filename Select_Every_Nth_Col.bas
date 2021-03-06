'   Upon selection of a range of cells, input your desired column separation number.

Sub Select_Every_Nth_Column()
    Application.ScreenUpdating = False
    Dim C As Range, u As Range
    Dim trig As Double
    Dim iTxt As Variant
    '-------------------
    iTxt = InputBox("For every separation : ", Default:=1) + 1
    '-------------------
    trig = 0
    For Each C In Selection.Columns
        trig = IIf(trig = 0, C.Column, trig)
        If ((C.Column - trig) Mod iTxt) = 0 Then
            If (u Is Nothing) Then
                Set u = C
            Else
                Set u = Union(u, C)
            End If
        End If
    Next
    u.Select
    '-------------------
End Sub
