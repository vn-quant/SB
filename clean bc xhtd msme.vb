Sub msme()
    Dim ws_xhtd As Worksheet, ws_msme As Worksheet, ws_clean As Worksheet
    Dim lastRow As Long, dataRows As Long, lastCol As Long
    Dim rng As Range
    Dim headers As Variant
    Dim dstLastRow As Long
    
    Set ws_xhtd = ThisWorkbook.Worksheets("Báo cáo XHTD")
    
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("MSME").Delete   ' delete if present
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set ws_msme = ThisWorkbook.Worksheets.Add
    ws_msme.Name = "MSME"
    
    
    Set ws_msme = ThisWorkbook.Worksheets("MSME")
    
    ' find last row in col A of source
    lastRow = ws_xhtd.Cells(ws_xhtd.Rows.Count, "A").End(xlUp).Row
    
    ' destination range in MSME (from B2 down)
    
    '--- Paste provided headers into row 1 ---
    headers = Array( _
        "STT", "NGAY_GUI_DUYET", "USER_NHAP_LIEU", "MA_DVKD", "NGAY_CHAM_DIEM", "USER_CHAM_DIEM", _
        "TEN_KH_CONG_TY", "ID_HO_SO", "ID_KHACH_HANG", "MA_LOS", "SAN_PHAM_CAP_TIN_DUNG", "DOI_TUONG_KH", _
        "VON_DIEU_LE", "DOANH_THU", "NAM_BCTC", "DIEM_XHTD_TONG", "PD", "KET_QUA_XH", _
        "HANG_SAU_DIEU_CHINH", "NGAY_DIEU_CHINH_HANG", "USER_DIEU_CHINH_HANG", "TRANG_THAI_XHTD", "TRANG_THAI_HIEU_LUC", _
        "GHI_CHU", "SC_SCORE_13", "SC_SCORE_32", "SC_SCORE_20", "SC_VALUE_01", "NGANH_CAP1", "TIME_SINCE_ESTABLISHMENT", _
        "DIEM_MO_HINH_PTC", "CROSS_FINANCIAL1_0014", "CROSS_FINANCIAL1_0023", "CROSS_FINANCIAL1_0187", "PERF7", "R_LIQ19", _
        "R_LIQ22", "R_SIZE7", "SIZE16", "DIEM_MO_HINH_TC", "ACCOUNT_FEE3_00031M_9M", "ACCOUNT3_0017_1M_6M", _
        "ACCOUNT3_0059_1M_9M", "CROSS1_00241M_6M", "CROSS1_0027", "LOAN2_00021M_6M", "LOAN2_00281M_3M", _
        "LOAN2_00351M_9M", "DIEM_MO_HINH_GD", "CIC_OTHERS2_00041M_1M_3M_3M", "CIC2_00121M_12M", "CIC1_0005", _
        "CIC2_00061M_1M_3M_3M", "CIC_OTHERS2_00031M_3M", "CIC_OTHERS2_00031M_6M", "DIEM_MO_HINH_CIC")
    
    ws_msme.Range("A1").Resize(1, UBound(headers) + 1).Value = headers
    
    '==================================================
    '                   TT XHTD
    '==================================================
    
    ' STT A:A
    ws_msme.Range("A2:A" & (lastRow - 7)).Value = ws_xhtd.Range("A9:A" & (lastRow)).Value
    
    ' NGAY_GUI_DUYET B:B, NGAY_CHAM_DIEM E:E
    ws_msme.Range("B2:B" & (lastRow - 7)).Formula = "=IF('Báo cáo XHTD'!B9="""","""",LEFT('Báo cáo XHTD'!B9,SEARCH("" "",'Báo cáo XHTD'!B9,1)))"
    ws_msme.Range("E2:E" & (lastRow - 7)).Formula = "=IF('Báo cáo XHTD'!E9="""","""",LEFT('Báo cáo XHTD'!E9,SEARCH("" "",'Báo cáo XHTD'!E9,1)))"
    
    ' USER_NHAP_LIEU C:C, MA_DVKD D:D
    ws_msme.Range("C2:D" & (lastRow - 7)).Value = ws_xhtd.Range("C9:D" & (lastRow)).Value
    
    ' "USER_CHAM_DIEM", "TEN_KH_CONG_TY", "ID_HO_SO", "ID_KHACH_HANG", "MA_LOS", "SAN_PHAM_CAP_TIN_DUNG", "DOI_TUONG_KH", "VON_DIEU_LE"
    ws_msme.Range("F2:M" & (lastRow - 7)).Value = ws_xhtd.Range("F9:M" & (lastRow)).Value
    
    '"DOANH_THU" N:N
    ws_msme.Range("N2:N" & (lastRow - 7)).Formula = "=IF('Báo cáo XHTD'!N9="""","""",SUBSTITUTE('Báo cáo XHTD'!N9,""."",""""))"

    '"NAM_BCTC" O:O
    'ws_msme.Range("O2:O" & (lastRow - 7)).Formula = "=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(Báo cáo XHTD'!o9,"Ngày",""),".",""),"Thang","/"),"Nam","/")"
    ws_msme.Range("O2:O" & (lastRow - 7)).Value = ws_xhtd.Range("O9:O" & (lastRow)).Value

    'DIEM XHTD TONG P:P
    ws_msme.Range("P2:P" & (lastRow - 7)).Formula = "=SUBSTITUTE('Báo cáo XHTD'!P9,"","",""."")"

    'PD -> "TRANG_THAI_XHTD", "TRANG_THAI_HIEU_LUC" , Q:W
    ws_msme.Range("Q2:Q" & (lastRow - 7)).Formula = "=SUBSTITUTE('Báo cáo XHTD'!Q9,"","",""."")"
    ws_msme.Range("R2:W" & (lastRow - 7)).Value = ws_xhtd.Range("R9:W" & (lastRow)).Value
    
    '==================================================
    '                  MH PTC
    '==================================================
    
    '  "SC_SCORE_13", "SC_SCORE_32", "SC_SCORE_20", "SC_VALUE_01", "NGANH_CAP1", "TIME_SINCE_ESTABLISHMENT": Y:AD
    ws_msme.Range("Y2:AD" & (lastRow - 7)).Value = ws_xhtd.Range("Y9:AD" & (lastRow)).Value
    
    'DIEM_MO_HINH_PTC: AE
    ws_msme.Range("AE2:AE" & (lastRow - 7)).Formula = "=SUBSTITUTE('Báo cáo XHTD'!AE9,"","",""."")"
    
    '==================================================
    '                  MH TC
    '==================================================
    
    'AF : AM
    ws_msme.Range("AF2:AM" & (lastRow - 7)).Value = ws_xhtd.Range("AF9:AM" & (lastRow)).Value
    
    'DIEM_MO_HINH_PC: AN
    ws_msme.Range("AN2:AN" & (lastRow - 7)).Formula = "=SUBSTITUTE('Báo cáo XHTD'!AN9,"","",""."")"
    
    '==================================================
    '                  MH GD
    '==================================================
    ' replace "." = "" then replace "," = "." ex: 68.521,099 -> 68521,099 -> 68521.099
    'AO:AW
    With ws_msme.Range("AO2:AW" & (lastRow - 7))
        .Formula = "=SUBSTITUTE(SUBSTITUTE('Báo cáo XHTD'!AO9,""."",""""),"","",""."")"
    End With

    '==================================================
    '                  MH cic
    '==================================================
    With ws_msme.Range("AX2:BD" & (lastRow - 7))
        .Formula = "=SUBSTITUTE(SUBSTITUTE('Báo cáo XHTD'!AX9,""."",""""),"","",""."")"
    End With
    
    '--- Replace "NA/NA" with blank across whole pasted block (A1:BD[dstLastRow]) ---
    ws_msme.Cells.Replace What:="NA/NA", Replacement:="", LookAt:=xlPart, MatchCase:=False
    
    '==================================================
    '                  data_clean
    '==================================================
    
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets("data_clean").Delete   ' delete if present
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws_clean = ThisWorkbook.Worksheets.Add
    ws_clean.Name = "data_clean"
    ws_clean.Range("A1").Resize(ws_msme.UsedRange.Rows.Count, _
                                ws_msme.UsedRange.Columns.Count).Value = ws_msme.UsedRange.Value

    
End Sub


