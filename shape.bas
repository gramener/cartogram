Private Sub Worksheet_Change(ByVal Target As Range)
    Dim KeyCells As Range

    ' The variable KeyCells contains the cells that will
    ' cause an alert when they are changed.
    Set KeyCells = Range("A:A")

    If Not Application.Intersect(KeyCells, Range(Target.Address)) _
           Is Nothing Then

        For Each Cell In Target.Cells
            ShapeName = Cell.Offset(0, 2).value & ":" & Cell.Offset(0, 4).value
            Set Shape = ActiveSheet.Shapes(ShapeName)
            With Shape.Fill
                .ForeColor.RGB = Gradient(Cell.value)
                .Visible = msoTrue
                .Transparency = 0
                .Solid
            End With
        Next

    End If
End Sub

Public Function Gradient(value)
    ' We'll always use a 3 point scale: 0, .5, 1
    g = Array(Array(1, 0, 0), Array(1, 1, 0), Array(0, 1, 0))
    Dim result As Variant

    If value < 0 Then
        result = g(0)
    ElseIf value < 0.5 Then
        a = g(0)
        b = g(1)
        q = 2 * value
        p = 1 - q
        result = Array(a(0) * p + b(0) * q, a(1) * p + b(1) * q, a(2) * p * b(2) * q)
    ElseIf value <= 1# Then
        a = g(1)
        b = g(2)
        q = 2 * (value - 0.5)
        p = 1 - q
        result = Array(a(0) * p + b(0) * q, a(1) * p + b(1) * q, a(2) * p * b(2) * q)
    Else
        result = g(2)
    End If

    Gradient = RGB(result(0) * 255, result(1) * 255, result(2) * 255)

End Function
