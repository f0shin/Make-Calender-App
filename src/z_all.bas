Attribute VB_Name = "z_all"
Option Explicit

Public Sign_MsgBox As Integer

Public Sign_MgMode As Integer

'관리자 모드<->사용자 모드 전환 (옵션: 1-shtMain에서 수동로그인, 0-프로그램 실행 중에 자동로그인)
Public Sub Change_Mode(Optional automode As Integer = 0)
    With shtMain.btnMgMode
        Select Case .Caption
            Case "사용자 모드"
                If automode = 1 Then
                    frmMgMode.Show
                    If Sign_MgMode = -1 Then: Exit Sub
                End If
                .Caption = "관리자 모드"
'                Visible_Sheets
                With shtMain
                    .Unprotect Password:="44775485520"
                    .EnableSelection = xlNoRestrictions
                End With
                Sign_MgMode = 0
                
            Case "관리자 모드"
                .Caption = "사용자 모드"
'                VeryHide_Sheets
                With shtMain
                    .Protect Password:="44775485520", DrawingObjects:=True, Contents:=True, Scenarios:=True
                    .EnableSelection = xlNoSelection
                End With
        End Select
    End With
End Sub

''시트 초기화
'Public Sub Init_Sheet(sht As Worksheet)
'    With sht
'
'    End With
'End Sub

'열려있는 모든 엑셀파일 저장 후 닫기
Public Function SaveClose_AllExcel() As Integer
    GetString1 = "열려있는 모든 엑셀파일이 저장 후 닫힙니다." & vbCrLf & "괜찮으시면 [확인]을 편집 중이라면 [취소]를 선택해주세요."
    frmMsgBox_OkCancel.Show
    If Sign_MsgBox = -1 Then: SaveClose_AllExcel = -1: Exit Function
    
    Dim ms As Workbook
    For Each ms In Application.Workbooks
        If ms.Name <> ThisWorkbook.Name Then
            ms.Save
            ms.Close
        End If
    Next
    
    GetString1 = "열려있던 모든 엑셀파일을 저장 후 닫기를 완료 하였습니다."
    frmMsgBox_OkOnly.Show
End Function
