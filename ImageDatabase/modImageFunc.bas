Attribute VB_Name = "modImageFunc"
'''''''''''''''''''''''''''''''''''''''''''
' Image Function Module
'
'
'''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''

Public Sub CreateThumb(PicBox As Object, ByVal ActualPic As StdPicture, ByVal MaxHeight As Integer, ByVal MaxWidth As Integer, Center As Boolean, Optional ByVal PicTop As Integer, Optional ByVal PicLeft As Integer)
    'MaxHeight is max. image height allowed
    'MaxWidth is max. picture width allowed
    Dim NewH As Integer 'New Height
    Dim NewW As Integer 'New Width
    'set starting var.
    NewH = ActualPic.Height 'actual image height
    NewW = ActualPic.Width 'actual image width
    'do logic


    If NewH > MaxHeight Or NewW > MaxWidth Then 'picture is too large


        If NewH > NewW Then 'height is greater than width
            NewW = Fix((NewW / NewH) * MaxHeight) 'rescale height
            NewH = MaxHeight 'set max height
        ElseIf NewW > NewH Then 'width is greater than height
            NewH = Fix((NewH / NewW) * MaxWidth) 'rescale width
            NewW = MaxHeight 'set max width
            Debug.Print "Width>"
        Else 'image is perfect square
            NewH = MaxHeight
            NewW = MaxWidth
        End If
    End If
    'check if centered


    If Center = True Then 'center picture
        PicTop = (PicBox.Height / 2) - (NewH / 2)
        PicLeft = (PicBox.Width / 2) - (NewW / 2)
    Else 'if Optional variables are missing Then and center=false


        If IsMissing(PicTop) = True Or IsMissing(PicLeft) = True Then
            PicTop = 0 'Default top position
            PicLeft = 0 'Default left position
        End If
    End If
    'Draw newly scaled picture


    With PicBox
        .AutoRedraw = True 'set needed properties
        .Cls 'clear picture box
        .PaintPicture ActualPic, PicLeft, PicTop, NewW, NewH 'paint new picture size in picturebox
    End With
End Sub

