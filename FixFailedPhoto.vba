Public Sub TambalFotoSlideAktif()
    Dim pptApp As Object, pptPres As Object, pptSlide As Object
    Dim ws As Worksheet
    Dim barisAktif As Long
    Dim linkFoto As String, directLink As String, fileID As String, tempPath As String
    Dim newPic As Object, shp As Object
    Dim hasilDownload As Boolean

    Set ws = ActiveSheet
    
    ' Ambil baris yang sedang diklik/dipilih di Excel
    barisAktif = ActiveCell.Row
    linkFoto = Trim(ws.Cells(barisAktif, 5).Value) ' Kolom E (Link Foto)

    If InStr(linkFoto, "/d/") = 0 Then
        MsgBox "Link foto di baris Excel yang dipilih kurang valid Bli!", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    Set pptSlide = pptApp.ActiveWindow.View.Slide ' Ambil slide yang lagi dibuka di layar PPT
    On Error GoTo 0

    If pptSlide Is Nothing Then
        MsgBox "Buka slide PPT yang ada FOTO GAGAL-nya dulu Bli!", vbExclamation
        Exit Sub
    End If

    ' Download Foto Baru
    fileID = Split(Split(linkFoto, "/d/")(1), "/")(0)
    directLink = "https://drive.google.com/thumbnail?id=" & fileID & "&sz=w800"
    tempPath = Environ("TEMP") & "\" & fileID & ".jpg"
    
    hasilDownload = DownloadGambarModern(directLink, tempPath)

    If hasilDownload = True Then
        On Error Resume Next
        
        ' Hapus Kotak Teks "FOTO GAGAL!" di Slide Tersebut
        For Each shp In pptSlide.Shapes
            If shp.HasTextFrame Then
                If InStr(shp.TextFrame.TextRange.Text, "FOTO GAGAL") > 0 Then
                    shp.Delete
                End If
            End If
        Next shp
        
        ' Tempel Foto Baru dengan Bingkai Emas Presisi
        Set newPic = pptSlide.Shapes.AddShape(5, 370, 800, 1200, 1800)
        newPic.Adjustments.Item(1) = 0.08
        newPic.Fill.UserPicture tempPath
        
        With newPic.Line
            .Visible = True
            .Weight = 16
            .ForeColor.RGB = RGB(255, 215, 0)
            .BackColor.RGB = RGB(160, 110, 15)
            .TwoColorGradient 3, 1
        End With
        
        On Error GoTo 0
        If Dir(tempPath) <> "" Then Kill tempPath
        
        MsgBox "BERES BLI! Foto berhasil ditambal tanpa re-generate!", vbInformation
    Else
        MsgBox "Foto masih gagal didownload, cek akses link Drive-nya lagi Bli!", vbCritical
    End If
End Sub

' =======================================================
' MESIN DOWNLOAD WINHTTP (WAJIB ADA DI DALAM MODULE)
' =======================================================
Private Function DownloadGambarModern(ByVal url As String, ByVal savePath As String) As Boolean
    Dim http As Object
    Dim stream As Object

    On Error GoTo ErrHandler

    ' Hapus file temp lama jika ada
    If Dir(savePath) <> "" Then
        On Error Resume Next
        Kill savePath
        On Error GoTo ErrHandler
    End If

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    http.send

    If http.Status = 200 Then
        Set stream = CreateObject("ADODB.Stream")
        stream.Open
        stream.Type = 1
        stream.Write http.responseBody
        stream.SaveToFile savePath, 2
        stream.Close
        DownloadGambarModern = True
    Else
        DownloadGambarModern = False
    End If

    Exit Function

ErrHandler:
    On Error Resume Next
    If Not stream Is Nothing Then
        If stream.State = 1 Then stream.Close
    End If
    DownloadGambarModern = False
End Function
