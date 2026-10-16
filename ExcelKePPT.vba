Sub ExcelKePPT_OTCBali_Fix()
    Dim pptApp As Object, pptPres As Object, pptSlideAsli As Object
    Dim slideRange As Object, pptSlide As Object
    Dim ws As Worksheet
    Dim BarisTerakhir As Long, i As Long
    Dim strNama As String, strProgram As String, strCabang As String, linkFoto As String
    Dim directLink As String, fileID As String, tempPath As String
    Dim newPic As Object, errBox As Object
    Dim hasilDownload As Boolean

    Set ws = ActiveSheet
    
    ' Cari baris terakhir berdasarkan Kolom B (Nama)
    BarisTerakhir = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    If BarisTerakhir < 2 Then Exit Sub

    On Error Resume Next
    Set pptApp = GetObject(, "PowerPoint.Application")
    Set pptPres = pptApp.ActivePresentation
    On Error GoTo 0

    If pptPres Is Nothing Then
        MsgBox "Buka PowerPoint template OTC-nya dulu Bli!", vbExclamation
        Exit Sub
    End If

    Set pptSlideAsli = pptPres.Slides(1)

    For i = 2 To BarisTerakhir
        ' Kolom B = Nama (2), C = Program (3), D = Cabang (4), E = Link Foto (5)
        strNama = Trim(ws.Cells(i, 2).Value)
        
        If strNama <> "" Then
            strProgram = Trim(ws.Cells(i, 3).Value)
            strCabang = Trim(ws.Cells(i, 4).Value)
            linkFoto = Trim(ws.Cells(i, 5).Value)
            
            ' Duplikat Slide
            Set slideRange = pptSlideAsli.Duplicate
            Set pptSlide = slideRange.Item(1)
            pptSlide.MoveTo pptPres.Slides.Count
            
            ' Ganti Teks
            GantiTeksLugas pptSlide, "[NAMA]", strNama
            GantiTeksLugas pptSlide, "[PROGRAM]", strProgram
            GantiTeksLugas pptSlide, "[CABANG]", strCabang
            
            ' Download & Tempel Foto Melengkung Berbingkai Gradasi Emas
            If InStr(linkFoto, "/d/") > 0 Then
                fileID = Split(Split(linkFoto, "/d/")(1), "/")(0)
                
                directLink = "https://drive.google.com/thumbnail?id=" & fileID & "&sz=w800"
                tempPath = Environ("TEMP") & "\" & fileID & ".jpg"
                
                hasilDownload = DownloadGambarModern(directLink, tempPath)
                
                If hasilDownload = True Then
                    On Error Resume Next
                    
                    ' ---------------------------------------------------------------
                    ' FOTO MELENGKUNG & BINGKAI GRADASI EMAS MEWAH
                    ' ---------------------------------------------------------------
                    
                    ' 1. Buat Bentuk Kotak Sudut Melengkung (Rounded Rectangle = 5)
                    Set newPic = pptSlide.Shapes.AddShape(5, 370, 800, 1200, 1800)
                    
                    ' 2. Atur Tingkat Kelengkungan Sudut (Ubah angka 0.08 ini jika ingin lebih/kurang melengkung)
                    newPic.Adjustments.Item(1) = 0.08
                    
                    ' 3. Masukkan Foto ke Dalam Bentuk
                    newPic.Fill.UserPicture tempPath
                    
                    ' 4. BINGKAI TEBAL & GRADASI EMAS
                    With newPic.Line
                        .Visible = True
                        .Weight = 16 ' Ketebalan Bingkai (Ubah jika kurang/terlalu tebal)
                        .ForeColor.RGB = RGB(255, 215, 0) ' Warna Emas Cerah (Gold)
                        .BackColor.RGB = RGB(160, 110, 15) ' Warna Emas Tua (Dark Gold)
                        .TwoColorGradient 3, 1 ' 3 = msoGradientDiagonalUp (Memberi efek kilauan gradasi miring)
                    End With
                    
                    On Error GoTo 0
                    
                    If Dir(tempPath) <> "" Then Kill tempPath
                Else
                    Set errBox = pptSlide.Shapes.AddTextbox(1, 370, 800, 1200, 1800)
                    errBox.TextFrame.TextRange.Text = "FOTO GAGAL!" & vbCrLf & "Link Rusak/Private"
                    errBox.TextFrame.TextRange.Font.Color.RGB = RGB(255, 0, 0)
                    errBox.TextFrame.TextRange.Font.Size = 14
                    errBox.TextFrame.TextRange.Font.Bold = msoTrue
                End If
            End If
        End If
    Next i

    MsgBox "BERES BLI! Semua kartu kelulusan OTC Bali berhasil digenerate dengan foto melengkung berbingkai gradasi emas!", vbInformation
End Sub

' =======================================================
' FUNGSI GANTI TEKS
' =======================================================
Sub GantiTeksLugas(sld As Object, cari As String, ganti As String)
    Dim shp As Object, j As Integer
    For Each shp In sld.Shapes
        If shp.Type = 6 Then
            For j = 1 To shp.GroupItems.Count
                ProsesReplace shp.GroupItems(j), cari, ganti
            Next j
        Else
            ProsesReplace shp, cari, ganti
        End If
    Next shp
End Sub

Sub ProsesReplace(shp As Object, cari As String, ganti As String)
    On Error Resume Next
    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            shp.TextFrame.TextRange.Replace cari, ganti
        End If
    End If
    On Error GoTo 0
End Sub

' =======================================================
' MESIN DOWNLOAD MODERN (WinHTTP)
' =======================================================
Function DownloadGambarModern(ByVal url As String, ByVal savePath As String) As Boolean
    Dim http As Object
    Dim stream As Object

    On Error GoTo ErrHandler

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
    DownloadGambarModern = False
End Function
