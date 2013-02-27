Private Sub Worksheet_Change(ByVal Target As Range)
    Dim KeyCells As Range

    ' The variable KeyCells contains the cells that will
    ' cause an alert when they are changed.
    Set KeyCells = Range("A:A")

    If Not Application.Intersect(KeyCells, Range(Target.Address)) _
           Is Nothing Then

        For Each Cell In Target.Cells
            ShapeName = Cell.Offset(0, 2).value & ":" & Cell.Offset(0, 4).value
            If ShapeName <> ":" Then
                Set Shape = ActiveSheet.Shapes(ShapeName)
                With Shape.Fill
                    .ForeColor.RGB = Gradient(Cell.value)
                    .Visible = msoTrue
                    .Transparency = 0
                    .Solid
                End With
            End If
        Next

    End If
End Sub


Public Function Gradient(value)
    ' We'll always use a 3 value scale
    g = Array(RGBval(Range("B1")), RGBval(Range("C1")), RGBval(Range("D1")))
    v = Array(Range("B1").Value, Range("C1").Value, Range("D1").Value)
    Dim result As Variant

    If value < v(0) Then
        result = g(0)
    ElseIf value < v(1) Then
        a = g(0)
        b = g(1)
        q = (value - v(0)) / (v(1) - v(0))
        p = 1 - q
        result = Array(a(0) * p + b(0) * q, a(1) * p + b(1) * q, a(2) * p + b(2) * q)
    ElseIf value <= v(2) Then
        a = g(1)
        b = g(2)
        q = (value - v(1)) / (v(2) - v(1))
        p = 1 - q
        result = Array(a(0) * p + b(0) * q, a(1) * p + b(1) * q, a(2) * p + b(2) * q)
    Else
        result = g(2)
    End If

    If result(0) < 0 Then result(0) = 0
    If result(1) < 0 Then result(1) = 0
    If result(2) < 0 Then result(2) = 0

    Gradient = RGB(result(0) * 255, result(1) * 255, result(2) * 255)
End Function


Public Function RGBval(Cell)
    ' Convert the background color of the cell into (r, g, b)
    c = Cell.Interior.Color
    b = c \ 65536
    c = c - b * 65536
    g = c \ 256
    r = c - g * 256
    RGBval = Array(r / 255, g / 255, b / 255)
End Function
