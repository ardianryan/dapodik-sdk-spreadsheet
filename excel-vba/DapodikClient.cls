VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DapodikClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Class: DapodikClient
' Deskripsi: Klien WebService Dapodik Kemendikdasmen untuk Excel VBA
' Pengembang: Ryan Ardian <inisaya@ardianryan.com>
' Publisher: SMA Negeri 1 Gedeg (@smansagewithai)
' Lisensi: MIT
'===============================================================================

Option Explicit

Private pNpsn As String
Private pToken As String
Private pHost As String
Private pPort As String
Private pBaseUrl As String
Private pTimeout As Long

Private Sub Class_Initialize()
    pHost = "127.0.0.1"
    pPort = "5774"
    pTimeout = 30000 ' 30 detik
    UpdateBaseUrl
End Sub

Public Sub Init(ByVal npsn As String, ByVal token As String, Optional ByVal host As String = "127.0.0.1", Optional ByVal port As String = "5774")
    If InStr(npsn, vbCr) > 0 Or InStr(npsn, vbLf) > 0 Then
        Err.Raise vbObjectError + 511, "DapodikClient", "NPSN tidak boleh mengandung karakter newline."
    End If
    If InStr(token, vbCr) > 0 Or InStr(token, vbLf) > 0 Then
        Err.Raise vbObjectError + 512, "DapodikClient", "Token tidak boleh mengandung karakter newline (CRLF injection prevention)."
    End If
    pNpsn = Trim(npsn)
    pToken = Trim(token)
    pHost = Trim(host)
    pPort = Trim(port)
    UpdateBaseUrl
End Sub

Private Sub UpdateBaseUrl()
    Dim cleanHost As String
    cleanHost = pHost
    If InStr(1, cleanHost, "http://", vbTextCompare) = 0 And InStr(1, cleanHost, "https://", vbTextCompare) = 0 Then
        cleanHost = "http://" + cleanHost
    End If
    pBaseUrl = cleanHost & ":" & pPort & "/WebService"
End Sub

Public Property Get BaseUrl() As String
    BaseUrl = pBaseUrl
End Property

Public Property Get npsn() As String
    npsn = pNpsn
End Property

'-------------------------------------------------------------------------------
' Request HTTP ke WebService Dapodik
'-------------------------------------------------------------------------------
Public Function Request(ByVal method As String, ByVal endpoint As String, Optional ByVal queryParams As String = "", Optional ByVal bodyJson As String = "") As String
    Dim http As Object
    Dim fullUrl As String
    Dim cleanEndpoint As String
    
    If pNpsn = "" Or pToken = "" Then
        Err.Raise vbObjectError + 513, "DapodikClient", "NPSN dan Token Dapodik wajib diisi terlebih dahulu."
    End If

    cleanEndpoint = endpoint
    If Left(cleanEndpoint, 1) = "/" Then cleanEndpoint = Mid(cleanEndpoint, 2)
    If InStr(cleanEndpoint, "..") > 0 Or InStr(cleanEndpoint, "\") > 0 Then
        Err.Raise vbObjectError + 515, "DapodikClient", "Endpoint tidak valid (path traversal detected)."
    End If
    
    fullUrl = pBaseUrl & "/" & cleanEndpoint & "?npsn=" & pNpsn
    If queryParams <> "" Then
        If Left(queryParams, 1) <> "&" Then queryParams = "&" & queryParams
        fullUrl = fullUrl & queryParams
    End If

    On Error Resume Next
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    If http Is Nothing Then Set http = CreateObject("MSXML2.XMLHTTP")
    If http Is Nothing Then Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    On Error GoTo 0

    If http Is Nothing Then
        Err.Raise vbObjectError + 514, "DapodikClient", "Gagal menginisialisasi komponen HTTP di Windows."
    End If

    http.Open UCase(method), fullUrl, False
    http.setRequestHeader "Authorization", "Bearer " & pToken
    http.setRequestHeader "Accept", "application/json, text/plain, */*"
    http.setRequestHeader "User-Agent", "dapodik-sdk-vba/1.0.0"

    If bodyJson <> "" Then
        http.setRequestHeader "Content-Type", "application/json"
        http.send bodyJson
    Else
        http.send
    End If

    If http.Status = 401 Or http.Status = 403 Then
        Err.Raise vbObjectError + 401, "DapodikClient", "Autentikasi Gagal (" & http.Status & "): Periksa token atau whitelist IP client di WebService Dapodik."
    End If

    If http.Status < 200 Or http.Status >= 300 Then
        Err.Raise vbObjectError + http.Status, "DapodikClient", "Dapodik HTTP Error " & http.Status & ": " & Left(http.responseText, 200)
    End If

    Request = http.responseText
End Function

'-------------------------------------------------------------------------------
' Endpoint Methods
'-------------------------------------------------------------------------------
Public Function GetSekolah(Optional ByVal semesterId As String = "") As String
    Dim q As String
    If semesterId <> "" Then q = "&semester_id=" & semesterId
    GetSekolah = Request("GET", "getSekolah", q)
End Function

Public Function GetPengguna() As String
    GetPengguna = Request("GET", "getPengguna")
End Function

Public Function GetGtk(Optional ByVal page As Long = 1, Optional ByVal limit As Long = 50) As String
    Dim q As String
    q = "&page=" & page & "&limit=" & limit
    GetGtk = Request("GET", "getGtk", q)
End Function

Public Function GetRombonganBelajar(Optional ByVal semesterId As String = "") As String
    Dim q As String
    If semesterId <> "" Then q = "&semester_id=" & semesterId
    GetRombonganBelajar = Request("GET", "getRombonganBelajar", q)
End Function

Public Function GetPesertaDidik(Optional ByVal page As Long = 1, Optional ByVal limit As Long = 50) As String
    Dim q As String
    q = "&page=" & page & "&limit=" & limit
    GetPesertaDidik = Request("GET", "getPesertaDidik", q)
End Function

Public Function GetMataPelajaran(Optional ByVal semesterId As String = "") As String
    Dim q As String
    If semesterId <> "" Then q = "&semester_id=" & semesterId
    GetMataPelajaran = Request("GET", "getMataPelajaran", q)
End Function

Public Function GetMatevNilai(Optional ByVal semesterId As String = "") As String
    Dim q As String
    If semesterId <> "" Then q = "&semester_id=" & semesterId
    GetMatevNilai = Request("GET", "getMatevNilai", q)
End Function

Public Function PostNilai(ByVal bodyJson As String, Optional ByVal semesterId As String = "") As String
    Dim q As String
    q = "&table=rapor"
    If semesterId <> "" Then q = q & "&semester_id=" & semesterId
    PostNilai = Request("POST", "postNilai", q, bodyJson)
End Function
