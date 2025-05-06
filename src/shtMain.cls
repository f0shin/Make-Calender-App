VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'https://www.timeanddate.com/holidays/ 기준
Private Const HOLLIDAY_URL As String = "https://www.timeanddate.com/holidays/"
Private Const MAX_YEAR As Integer = 2040
Private Const MIN_YEAR As Integer = 2020

Dim requestYear As String
Dim requestMonth As String
Dim requestCountry As String

Dim arrHoliday() As Date '휴일 목록

'***************************************************
'이벤트
'***************************************************
'시작
Private Sub btnRun_Click()
    '세팅
    requestYear = shtMain.txtYear
    requestMonth = shtMain.txtMonth
    If shtMain.opCountryKR.Value Then
        requestCountry = "south-korea"
    Else
        requestCountry = "japan"
    End If
    
    '휴일 가져오기
    Get_Holiday
    
    '달력 만들기
    Make_Calender

    '종료
    'MsgBox "Main : Completed!"
    MsgBox "Completed!"
End Sub

Private Sub btnYearUp_Click()
    Dim y As Integer
    y = CInt(shtMain.txtYear.Value)
    y = y + 1
    If y > MAX_YEAR Then
        y = year(Now())
    End If
    shtMain.txtYear.Value = CStr(y)
End Sub

Private Sub btnYearDown_Click()
    Dim y As Integer
    y = CInt(shtMain.txtYear.Value)
    y = y - 1
    If y < MIN_YEAR Then
        y = year(Now())
    End If
    shtMain.txtYear.Value = CStr(y)
End Sub

Private Sub btnMonthUp_Click()
    Dim m As Integer
    m = CInt(shtMain.txtMonth.Value)
    m = m + 1
    If m > 12 Then
        m = 1
    End If
    shtMain.txtMonth.Value = CStr(m)
End Sub

Private Sub btnMonthDown_Click()
    Dim m As Integer
    m = CInt(shtMain.txtMonth.Value)
    m = m - 1
    If m < 1 Then
        m = 12
    End If
    shtMain.txtMonth.Value = CStr(m)
End Sub

Private Sub btnAdmin_Click()
    Static isAdminMode As Boolean
    
    '모드 토글
    isAdminMode = Not isAdminMode
    
    '관리자 모드 활성화
    If isAdminMode Then
        LockCellsAndShapes
        shtMain.btnAdmin.Caption = "User Mode"
        MsgBox "Admin Mode : ON", vbInformation
    Else
        UnlockCellsAndShapes
        shtMain.btnAdmin.Caption = "Admin Mode"
        MsgBox "Admin Mode : OFF", vbInformation
    End If
End Sub

Private Sub LockCellsAndShapes()
    Dim shp As Shape

    '시트 보호: 셀 수정 금지, 행/열 삽입 금지, 도형 이동 제한
    shtMain.Protect Password:="1234", _
                    UserInterfaceOnly:=True, _
                    AllowFormattingCells:=False, _
                    AllowFormattingColumns:=False, _
                    AllowFormattingRows:=False, _
                    AllowInsertingColumns:=False, _
                    AllowInsertingRows:=False
    
    '모든 도형 이동 제한
    For Each shp In shtMain.Shapes
        shp.Locked = True '도형 잠금
        shp.Placement = xlFreeFloating '도형 이동 불가
    Next shp
End Sub

Private Sub UnlockCellsAndShapes()
    Dim shp As Shape

    '시트 보호 해제
    shtMain.Unprotect Password:="1234"

    '도형 이동 허용
    For Each shp In shtMain.Shapes
        shp.Locked = False
        shp.Placement = xlMoveAndSize '이동 가능하게 설정
    Next shp
End Sub


'*********************************************
'프로시저 모음
'*********************************************

Public Function Get_Holiday()
    Dim http As Object
    Dim html As Object
    Dim url As String
    Dim table As Object
    Dim tr As Object
    Dim r As Integer: r = 0
    Dim unixDate As Long
    Dim arrResult() As Date
    
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP")
    url = HOLLIDAY_URL & requestCountry & "/" & requestYear & "?hol=9"
    http.Open "GET", url, False
    http.send
    
    '오류 발생 시 스킵
    If Err.Number <> 0 Or http.Status <> 200 Then
        MsgBox "웹 요청 실패! 인터넷 연결을 확인하세요.", vbExclamation
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    
    Set html = CreateObject("HTMLFile")
    html.body.innerHTML = http.responseText

    Set table = html.getElementById("holidays-table")
    If table Is Nothing Then Exit Function ' 페이지 구조가 변경되었을 경우 스킵

    For Each tr In table.getElementsByTagName("tr")
        If tr.className = "showrow" Then
            unixDate = tr.getAttribute("data-date") / 1000
            ReDim Preserve arrResult(0 To r)
            arrResult(r) = (((unixDate / 60) / 60) / 24) + DateSerial(1970, 1, 1)
            r = r + 1
        End If
    Next
    
    '공휴일 데이터를 설정 (비어 있는 경우에도 문제 없음)
    If r > 0 Then
        ReDim arrHoliday(0 To UBound(arrResult))
        For r = 0 To UBound(arrResult)
            arrHoliday(r) = arrResult(r)
        Next r
    Else
        ReDim arrHoliday(0) ' 공휴일 없음 처리
    End If

    'MsgBox "Get_Holiday : Completed!"
    
    '객체 해제
    Set http = Nothing
    Set html = Nothing
    Set table = Nothing
    Set tr = Nothing
End Function


Private Sub Make_Calender()
    Dim currentRow As Integer
    Dim currentCol As Integer
    
    Dim newWorkbook As Workbook
    Dim wsBase As Worksheet
    Dim wsNew As Worksheet
    
    Dim firstDate As Date
    Dim lastDate As Date
    
    Dim holiday As Variant
    Dim currentYear As Integer
    Dim currentMonth As Integer
    
    ' 새 파일 생성 및 시트 복사
    Set newWorkbook = Workbooks.Add
    Set wsBase = ThisWorkbook.Sheets("export")
    wsBase.Copy Before:=newWorkbook.Sheets(1)
    
    ' 연월 설정
    currentYear = CInt(requestYear)
    currentMonth = CInt(requestMonth)
    
    Set wsNew = ActiveSheet
    wsNew.Name = currentYear & "_" & currentMonth
    wsNew.Cells(1, 2) = currentYear & " " & GetMonthName(currentMonth)
    
    ' 날짜 설정
    firstDate = DateSerial(currentYear, currentMonth, 1)
    lastDate = DateSerial(currentYear, currentMonth + 1, 0)
    
    ' 첫 번째 열 설정
    currentCol = GetCurrentCol(firstDate)
    currentRow = 3
    
    ' 날짜 채우기
    Dim currentDay As Integer
    For currentDay = 0 To lastDate - firstDate
        Dim currentDate As Date
        currentDate = firstDate + currentDay
        wsNew.Cells(currentRow, currentCol) = currentDay + 1
        
        ' 공휴일 데이터가 존재할 경우에만 검사
        If UBound(arrHoliday) >= 0 Then
            For Each holiday In arrHoliday
                If holiday = currentDate Then
                    wsNew.Cells(currentRow, currentCol).Font.Color = RGB(255, 0, 0) ' 빨간색
                    Exit For
                End If
            Next holiday
        End If
        
        ' 다음 열로 이동
        currentCol = currentCol + 2
        If currentCol > 14 Then
            currentCol = 2
            currentRow = currentRow + 2
            If currentRow > 11 Then
                '6주차가 된 경우 : 11,12 행을 복사해서 13,14행에 붙여넣고 값은 지움.
                wsNew.Rows("11:12").Copy
                wsNew.Rows("13:14").PasteSpecial xlPasteFormats ' 서식만 복사
                wsNew.Rows("13:14").ClearContents
            End If
        End If
    Next currentDay
    
    ' 포커스를 A1에 두기
    wsNew.Activate
    wsNew.Range("A1").Select
    
    ' 빈 시트 삭제 (Sheet1이 존재하면 제거)
    Application.DisplayAlerts = False ' 경고 메시지 방지
    On Error Resume Next
    newWorkbook.Sheets("Sheet1").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True ' 경고 메시지 다시 활성화

    'MsgBox "Make_Calender : Completed! " & wsNew.Name, vbInformation
    
End Sub


'첫번재 날이 무슨 요일인지에 따른 첫번째 열 설정
Function GetCurrentCol(firstDate As Date) As Integer
    Dim weekdayValue As Integer
    weekdayValue = Weekday(firstDate, vbSunday) ' 일요일 시작 기준

    Select Case weekdayValue
        Case 1: GetCurrentCol = 2  ' Sunday
        Case 2: GetCurrentCol = 4  ' Monday
        Case 3: GetCurrentCol = 6  ' Tuesday
        Case 4: GetCurrentCol = 8  ' Wednesday
        Case 5: GetCurrentCol = 10 ' Thursday
        Case 6: GetCurrentCol = 12 ' Friday
        Case 7: GetCurrentCol = 14 ' Saturday
        Case Else: GetCurrentCol = 2 ' 기본값 (오류 방지)
    End Select
End Function

'월 이름
Function GetMonthName(currentMonth As Integer) As String
    Select Case currentMonth
        Case 1: GetMonthName = "January"
        Case 2: GetMonthName = "February"
        Case 3: GetMonthName = "March"
        Case 4: GetMonthName = "April"
        Case 5: GetMonthName = "May"
        Case 6: GetMonthName = "June"
        Case 7: GetMonthName = "July"
        Case 8: GetMonthName = "August"
        Case 9: GetMonthName = "September"
        Case 10: GetMonthName = "October"
        Case 11: GetMonthName = "November"
        Case 12: GetMonthName = "December"
        Case Else: GetMonthName = "Invalid Month"
    End Select
End Function

