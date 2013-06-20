Public Function UNID()
    Set WMI = GetObject("winmgmts:!\\.\root\cimv2")
    Set Board = WMI.ExecQuery("Select * from Win32_BaseBoard")
    For Each b In Board
        UNID = b.SerialNumber
    Next b
End Function

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim KeyCells As Range

    ' The variable KeyCells contains the cells that will
    ' cause an alert when they are changed.
    Set KeyCells = Range("A:A")

    If Not Application.Intersect(KeyCells, Range(Target.Address)) _
           Is Nothing Then

        ' Ensure file expires
        Dim Expires As Date
        Expires = #01/01/2013#
        If Now > Expires Then
            MsgBox "This license expired on " & Expires
            Exit Sub
        End If

        ' Validate the UNID
        LicenseKey = "LICENSEKEY"
        ID = UNID()
        If ID <> LicenseKey Then
            InputBox "This file is only valid for the machine " & LicenseKey & _
                ". Your machine is:", "Invalid License", ID
            Exit Sub
        End If

        ' Get the range of colors for up to 255 cells (B-IV)
        Dim g(0 To 255)
        Dim v(0 To 255)
        For Each Cell In Range("B1:IV1")
            If Cell.value <> "" Then
                v(Cell.Column - 2) = Cell.value
                g(Cell.Column - 2) = RGBval(Cell)
                n = Cell.Column - 1
            End If
        Next

        For Each Cell In Target.Cells
            ShapeName = Cell.Offset(0, 1).value
            If ShapeName <> "" Then
                Set Shape = ActiveSheet.Shapes(ShapeName)
                With Shape.Fill
                    .ForeColor.RGB = Gradient(Cell.value, g, v, n)
                    .Visible = msoTrue
                    .Transparency = 0
                    .Solid
                End With
            End If
        Next

    End If
End Sub


Public Function Gradient(value, g, v, n)
    ' We'll always use a 3 value scale
    ' g = Array(RGBval(Range("B1")), RGBval(Range("C1")), RGBval(Range("D1")))
    ' v = Array(Range("B1").value, Range("C1").value, Range("D1").value)
    ' n = 3
    Dim result As Variant

    If value < v(0) Then
        result = g(0)
    ElseIf value >= v(n - 1) Then
        result = g(n - 1)
    Else
        i = 1
        While value >= v(i)
            i = i + 1
        Wend
        a = g(i - 1)
        b = g(i)
        If value = v(i) Then
            q = 1
        ElseIf value = v(i - 1) Then
            q = 0
        Else
            q = (value - v(i - 1)) / (v(i) - v(i - 1))
        End If
        p = 1 - q
        result = Array(a(0) * p + b(0) * q, a(1) * p + b(1) * q, a(2) * p + b(2) * q)
    End If

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

Sub Filter()
    s = LCase(InputBox("Type in the regions to select", "Filter"))
    If Len(s) = 0 Then
        Exit Sub
    End If
    For Each shp In ActiveSheet.Shapes
        ' Type 5 = msoFreeform. Type 6 = msoGroup. Ignore buttons, etc.
        If (shp.Type = 5 Or shp.Type = 6) And InStr(LCase(shp.Name), s) > 0 Then
            shp.Select (False)
        End If
    Next
End Sub
