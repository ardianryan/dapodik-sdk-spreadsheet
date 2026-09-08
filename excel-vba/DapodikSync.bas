Attribute VB_Name = "DapodikSync"
'===============================================================================
' Module: DapodikSync
' Deskripsi: Macro dan fungsi sinkronisasi Dapodik ke Excel Worksheet
' Pengembang: Ryan Ardian <inisaya@ardianryan.com>
' Publisher: SMA Negeri 1 Gedeg (@smansagewithai)
'===============================================================================

Option Explicit

'-------------------------------------------------------------------------------
' Inisialisasi Klien dari Sheet 'Config'
'-------------------------------------------------------------------------------
Public Function GetDapodikClientFromSheet() As DapodikClient
    Dim wsConfig As Worksheet
    Dim npsn As String, token As String, host As String, port As String
    Dim client As DapodikClient
    
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Config")
    On Error GoTo 0
    
    If wsConfig Is Nothing Then
        MsgBox "Sheet 'Config' tidak ditemukan. Jalankan macro SetupConfigSheet terlebih dahulu.", vbCritical, "Dapodik SDK"
        Exit Function
    End If
    
    npsn = Trim(wsConfig.Range("B2").Value)
    token = Trim(wsConfig.Range("B3").Value)
    host = Trim(wsConfig.Range("B4").Value)
    port = Trim(wsConfig.Range("B5").Value)
    
    If npsn = "" Or token = "" Then
        MsgBox "NPSN dan Token pada Sheet 'Config' tidak boleh kosong.", vbExclamation, "Dapodik SDK"
        Exit Function
    End If
    
    If host = "" Then host = "127.0.0.1"
    If port = "" Then port = "5774"
    
    Set client = New DapodikClient
    client.Init npsn, token, host, port
    Set GetDapodikClientFromSheet = client
End Function

'-------------------------------------------------------------------------------
' Macro Setup Sheet Config Otomatis
'-------------------------------------------------------------------------------
Public Sub SetupConfigSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Config")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = "Config"
    Else
        ws.Cells.Clear
    End If
    
    ws.Range("A1:C1").Value = Array("Parameter", "Nilai", "Keterangan")
    ws.Range("A2:C2").Value = Array("NPSN", "20300001", "8 Digit NPSN Sekolah")
    ws.Range("A3:C3").Value = Array("TOKEN", "PASTE_TOKEN_WEBSERVICE_DISINI", "Token dari WebService Dapodik")
    ws.Range("A4:C4").Value = Array("HOST", "127.0.0.1", "IP Komputer Server Dapodik")
    ws.Range("A5:C5").Value = Array("PORT", "5774", "Port WebService Dapodik (default 5774)")
    
    With ws.Range("A1:C1")
        .Interior.Color = RGB(27, 94, 32)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    ws.Columns("A:C").AutoFit
    MsgBox "Sheet 'Config' berhasil disiapkan! Silakan masukkan Token dan NPSN sekolah Anda.", vbInformation, "Dapodik SDK"
End Sub

'-------------------------------------------------------------------------------
' Macro Tarik Siswa (Peserta Didik)
'-------------------------------------------------------------------------------
Public Sub TarikDataSiswa()
    Dim client As DapodikClient
    Dim rawJson As String
    Dim ws As Worksheet
    
    Set client = GetDapodikClientFromSheet()
    If client Is Nothing Then Exit Sub
    
    Application.StatusBar = "Sedang menarik data siswa dari Dapodik..."
    Application.ScreenUpdating = False
    
    On Error GoTo ErrHandler
    rawJson = client.GetPesertaDidik(1, 100)
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Data Siswa")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Data Siswa"
    End If
    
    ' Parse JSON dan tulis ke worksheet (memerlukan JsonConverter.bas)
    PopulateJsonToSheet ws, rawJson
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Data Siswa berhasil ditarik ke sheet 'Data Siswa'!", vbInformation, "Dapodik SDK"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Gagal menarik data siswa: " & Err.Description, vbCritical, "Dapodik SDK"
End Sub

'-------------------------------------------------------------------------------
' Macro Tarik Prasarana & Bangunan
'-------------------------------------------------------------------------------
Public Sub TarikDataPrasarana()
    Dim client As DapodikClient
    Dim rawJson As String
    Dim ws As Worksheet
    
    Set client = GetDapodikClientFromSheet()
    If client Is Nothing Then Exit Sub
    
    Application.StatusBar = "Sedang menarik data prasarana dari Dapodik..."
    Application.ScreenUpdating = False
    
    On Error GoTo ErrHandler
    rawJson = client.GetPrasarana(1, 100)
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Data Prasarana")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Data Prasarana"
    End If
    
    PopulateJsonToSheet ws, rawJson
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Data Prasarana berhasil ditarik ke sheet 'Data Prasarana'!", vbInformation, "Dapodik SDK"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Gagal menarik data prasarana: " & Err.Description, vbCritical, "Dapodik SDK"
End Sub

'-------------------------------------------------------------------------------
' Helper untuk mengisi JSON ke Worksheet
'-------------------------------------------------------------------------------
Private Sub PopulateJsonToSheet(ws As Worksheet, jsonText As String)
    Dim parsed As Object
    Dim rows As Object
    Dim item As Object
    Dim key As Variant
    Dim r As Long, c As Long
    Dim colIndex As Object
    
    Set colIndex = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set parsed = JsonConverter.ParseJson(jsonText)
    On Error GoTo 0
    
    If parsed Is Nothing Then
        ws.Range("A1").Value = "Raw JSON Response:"
        ws.Range("A2").Value = jsonText
        Exit Sub
    End If
    
    Set rows = parsed("rows")
    If rows Is Nothing Then Set rows = parsed("data")
    If rows Is Nothing Then Exit Sub
    
    ws.Cells.Clear
    
    ' Cari semua header unik
    c = 1
    For Each item In rows
        For Each key In item.Keys
            If Not colIndex.Exists(CStr(key)) Then
                colIndex.Add CStr(key), c
                ws.Cells(1, c).Value = key
                c = c + 1
            End If
        Next key
    Next item
    
    ' Tulis data baris
    r = 2
    For Each item In rows
        For Each key In item.Keys
            ws.Cells(r, colIndex(CStr(key))).Value = item(key)
        Next key
        r = r + 1
    Next item
    
    ' Format Header
    If c > 1 Then
        With ws.Range(ws.Cells(1, 1), ws.Cells(1, c - 1))
            .Interior.Color = RGB(13, 71, 161)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
        End With
        ws.Columns.AutoFit
    End If
End Sub
