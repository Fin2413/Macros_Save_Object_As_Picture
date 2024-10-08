Attribute VB_Name = "Module1"
Sub Save_Object_As_Picture()
    ' ���������� ����������
    Dim li As Long, oObj As Shape, wsSh As Worksheet, wsTmpSh As Worksheet
    Dim sImagesPath As String, sName As String, productName As String
    Dim topLeftCell As Range
 
    ' ��������� ���� ��� ���������� �����������
    sImagesPath = ActiveWorkbook.Path & "\images\" ' ����� ��� ���������� ����������� � ������� �������� �����
    
    ' �������� �����, ���� ��� �� ����������
    If Dir(sImagesPath, 16) = "" Then
        MkDir sImagesPath ' ������� ����� ��� �����������, ���� � ���
    End If
    
    ' ���������� ���������� ������ � ��������������, ����� �������� ����������
    On Error Resume Next ' ������������ ������
    Application.ScreenUpdating = False ' ��������� ���������� ������
    Application.DisplayAlerts = False ' ��������� ��������������

    ' ��������� �������� ����� � �������� ���������� �����
    Set wsSh = ActiveSheet ' ��������� ��������� �����
    Set wsTmpSh = ActiveWorkbook.Sheets.Add ' ���������� ���������� ����� ��� ������ � ��������

    ' ������� ���� �������� �� �������� �����
    For Each oObj In wsSh.Shapes
        ' ��������, �������� �� ������ ������������
        If oObj.Type = 13 Then ' ��� 13 � ��� �����������
            li = li + 1 ' ������� ��� ���� �����������
            
            ' �������� ������, ��� ��������� ������� ����� ���� �������
            Set topLeftCell = oObj.topLeftCell
            
            ' �������� ��� �������� �� ������� ������� (������� A) ��� �� ������, ��� � �����������
            productName = wsSh.Cells(topLeftCell.Row, 1).Value ' ������������ �� ������� A
            
            ' ������� ������������ ������� �� ����� �����
            productName = Replace(productName, "/", "_") ' �������� �����
            productName = Replace(productName, "\", "_") ' �������� �������� �����
            productName = Replace(productName, ":", "_") ' �������� ���������
            productName = Replace(productName, "*", "_") ' �������� ���������
            productName = Replace(productName, "?", "_") ' �������� �������������� �����
            productName = Replace(productName, """", "_") ' �������� �������
            productName = Replace(productName, "<", "_") ' �������� ������� �����
            productName = Replace(productName, ">", "_") ' �������� ������� �����
            productName = Replace(productName, "|", "_") ' �������� ������������ �����
            
            ' ���� ��� �������� ������, ���������� ����������� ���
            If productName = "" Then
                productName = "img" & li
            End If
            
            ' �������� �����������
            oObj.Copy

            ' ������������� ���������� ������� ��� �������� �����������
            With wsTmpSh.ChartObjects.Add(0, 0, oObj.Width, oObj.Height).Chart
                .ChartArea.Border.LineStyle = 0 ' ������� ������� �������
                .Parent.Select ' �������� ������
                .Paste ' ��������� ����������� � ������
                .Export Filename:=sImagesPath & productName & ".jpg", FilterName:="JPG" ' ������������ ����������� ��� ���� JPG
                .Parent.Delete ' ������� ��������� ������ ����� ���������� �����������
            End With
            
            ' ���������� ��� ����� � ������, ��� ���������� �����������
            oObj.topLeftCell.Value = productName ' ���������� ��� ����� � ������
        End If
    Next oObj

    ' ������������ ������
    Set oObj = Nothing
    Set wsSh = Nothing
    wsTmpSh.Delete ' �������� ���������� �����

    ' �������� ������� ���������� ������ � ��������������
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ' ��������� � ���������� ��������
    MsgBox "������� ��������� � �����: " & sImagesPath, vbInformation, "�����"
End Sub

