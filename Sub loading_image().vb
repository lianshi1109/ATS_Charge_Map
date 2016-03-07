Sub loading_image()

Dim FD
Dim Image_location_y
Dim Image_location_x
Dim f
Dim location_y
Dim location_x
Dim filename
Dim myname



    Set fso = CreateObject("Scripting.FileSystemObject")
    Set OAK = ActiveWorkbook
    myPath = "\\whfile.csot.tcl.com\11制造中心\整合厂\工艺整合部\2.阵列工艺整合科\12 personal\李安石\ATS图片拼接\Image"    '把文件路径定义给变量
    Application.ScreenUpdating = False
    Set FD = fso.GetFolder(myPath)
     
    For Each f In FD.Files
    filename = VBA.Right(f.Name, 3)
    myname = f.Name
    If filename = "jpg" Then
    
    Image_location_x = VBA.Left(f.Name, 1)
    Image_location_y = Right(VBA.Left(f.Name, 2), 1)
    Sheets("lookup").Select
    Cells(1, 2) = Image_location_x
    Cells(3, 2) = Image_location_y
    location_x = Cells(2, 2)
    location_y = 0 - Cells(4, 2)
    Sheets("mapping").Select

    Cells(17, 6).Offset(location_y, location_x).Select
    
    ActiveSheet.Pictures.Insert("\\whfile.csot.tcl.com\11制造中心\整合厂\工艺整合部\2.阵列工艺整合科\12 personal\李安石\ATS图片拼接\Image\" & myname).Select
       
    Selection.ShapeRange.LockAspectRatio = msoFalse
    Selection.Placement = xlMoveAndSize
    Selection.ShapeRange.Left = Cells(17, 6).Offset(location_y, location_x).Left
    Selection.ShapeRange.Top = Cells(17, 6).Offset(location_y, location_x).Top
    Selection.ShapeRange.Height = Cells(17, 6).Offset(location_y, location_x).Height
    Selection.ShapeRange.Width = Cells(17, 6).Offset(location_y, location_x).Width
 
    End If
    Next
    
End Sub
