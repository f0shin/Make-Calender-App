Attribute VB_Name = "z_all"
Option Explicit

Public Sign_MsgBox As Integer

Public Sign_MgMode As Integer

'������ ���<->����� ��� ��ȯ (�ɼ�: 1-shtMain���� �����α���, 0-���α׷� ���� �߿� �ڵ��α���)
Public Sub Change_Mode(Optional automode As Integer = 0)
    With shtMain.btnMgMode
        Select Case .Caption
            Case "����� ���"
                If automode = 1 Then
                    frmMgMode.Show
                    If Sign_MgMode = -1 Then: Exit Sub
                End If
                .Caption = "������ ���"
'                Visible_Sheets
                With shtMain
                    .Unprotect Password:="44775485520"
                    .EnableSelection = xlNoRestrictions
                End With
                Sign_MgMode = 0
                
            Case "������ ���"
                .Caption = "����� ���"
'                VeryHide_Sheets
                With shtMain
                    .Protect Password:="44775485520", DrawingObjects:=True, Contents:=True, Scenarios:=True
                    .EnableSelection = xlNoSelection
                End With
        End Select
    End With
End Sub

''��Ʈ �ʱ�ȭ
'Public Sub Init_Sheet(sht As Worksheet)
'    With sht
'
'    End With
'End Sub

'�����ִ� ��� �������� ���� �� �ݱ�
Public Function SaveClose_AllExcel() As Integer
    GetString1 = "�����ִ� ��� ���������� ���� �� �����ϴ�." & vbCrLf & "�������ø� [Ȯ��]�� ���� ���̶�� [���]�� �������ּ���."
    frmMsgBox_OkCancel.Show
    If Sign_MsgBox = -1 Then: SaveClose_AllExcel = -1: Exit Function
    
    Dim ms As Workbook
    For Each ms In Application.Workbooks
        If ms.Name <> ThisWorkbook.Name Then
            ms.Save
            ms.Close
        End If
    Next
    
    GetString1 = "�����ִ� ��� ���������� ���� �� �ݱ⸦ �Ϸ� �Ͽ����ϴ�."
    frmMsgBox_OkOnly.Show
End Function
