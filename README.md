Attribute VB_Name = "BOM_MACRO"
Option Explicit

' ==============================================================================
' BOM_MACRO — SW 2023 x64 + Windows 10
' - Top-level bileşenleri, Unsur ağacında göründüğü sırayla Excel’e yazar.
' - A: Görsel (Windows Shell thumb), B..F: Parça Adı, Malzeme, Adet, Birim kg, Toplam kg
' - Kütle: kg garantili (GetMassProperties2 → IMassProperty(MKS) → SW-Mass fallback)
' - Hücre içi görsel: dinamik boyutlandırma; geçici dosyalar otomatik temizlenir.
' ==============================================================================

' Excel ve SW sabitleri (late binding)
Private Const xlCalculationManual    As Long = -4135
Private Const xlCalculationAutomatic As Long = -4105
Private Const xlMoveAndSize          As Long = 1

' Hizalama sabitleri (Center = ortala)
Private Const xlHAlignCenter         As Long = -4108
Private Const xlVAlignCenter         As Long = -4108

Private Const swDocPART              As Long = 1
Private Const swDocASSEMBLY          As Long = 2
Private Const swIsometricView        As Long = 7
Private Const swOpenDocOptions_Silent As Long = 1
Private Const swOpenDocOptions_ReadOnly As Long = 2

' Görsel yerleşim ayarları
Private Const THUMB_ROW_HEIGHT       As Double = 70
Private Const THUMB_MARGIN_PT        As Double = 4
Private Const COL_WIDTH_THUMB        As Double = 12

Private m_PS1Path As String

Public Sub Main()
    On Error GoTo FATAL

    Dim swApp As Object, swAssy As Object
    Set swApp = Application.SldWorks
    If swApp Is Nothing Then MsgBox "SolidWorks bulunamadı.", vbCritical: Exit Sub

    Set swAssy = swApp.ActiveDoc
    If swAssy Is Nothing Then MsgBox "Aktif doküman yok.", vbExclamation: Exit Sub
    If swAssy.GetType <> swDocASSEMBLY Then MsgBox "Montaj (SLDASM) açın.", vbExclamation: Exit Sub

    Dim orderedComps As Variant
    orderedComps = GetTopLevelComponentsInTreeOrder(swAssy)
    If IsEmpty(orderedComps) Then
        MsgBox "Top-level bileşen bulunamadı.", vbInformation: Exit Sub
    End If

    m_PS1Path = WriteThumbnailScript()

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim orderList As Collection: Set orderList = New Collection

    Dim i As Long, swComp As Object, swPDoc As Object
    Dim sPath As String, sCfg As String, key As String
    Dim idx As Long: idx = 0

    For i = LBound(orderedComps) To UBound(orderedComps)
        Set swComp = orderedComps(i)
        If swComp Is Nothing Then GoTo NextComp
        If SafeIsSuppressed(swComp) Or SafeIsEnvelope(swComp) Or SafeExcludeFromBOM(swComp) Then GoTo NextComp

        Set swPDoc = swComp.GetModelDoc2
        If swPDoc Is Nothing Then GoTo NextComp
        If swPDoc.GetType <> swDocPART Then GoTo NextComp

        sPath = swPDoc.GetPathName: If Len(sPath) = 0 Then sPath = swComp.Name2
        sCfg  = swComp.ReferencedConfiguration
        If Len(sCfg) = 0 Then sCfg = swPDoc.GetActiveConfigurationName

        key = LCase$(sPath) & "::" & LCase$(sCfg)

        If dict.Exists(key) Then
            Dim ar As Variant: ar = dict(key): ar(1) = ar(1) + 1: dict(key) = ar
        Else
            idx = idx + 1
            Dim tp As String: tp = GetShellThumbnail(sPath, idx)
            dict.Add key, Array(GetCleanName(swComp.Name2), 1, _
                                GetMaterialSafe(swPDoc, sCfg), _
                                GetMass_KG(swComp, swPDoc, sCfg), tp)
            orderList.Add key
        End If
NextComp:
    Next i

    On Error Resume Next
    If Len(m_PS1Path) > 0 Then If Dir(m_PS1Path) <> "" Then Kill m_PS1Path
    On Error GoTo FATAL

    If dict.Count = 0 Then MsgBox "Parça yok.", vbInformation: Exit Sub

    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    On Error Resume Next
    Set xlApp = GetObject(, "Excel.Application")
    If xlApp Is Nothing Then Set xlApp = CreateObject("Excel.Application")
    On Error GoTo FATAL

    xlApp.ScreenUpdating = False
    SetExcelCalc xlApp, xlCalculationManual

    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Worksheets(1)
    xlSheet.Name = "BOM Listesi"

    xlSheet.Cells(1, 1).Value = "GORSEL"
    xlSheet.Cells(1, 2).Value = "PARCA ADI"
    xlSheet.Cells(1, 3).Value = "MALZEME"
    xlSheet.Cells(1, 4).Value = "ADET"
    xlSheet.Cells(1, 5).Value = "BIRIM AGIRLIK (kg)"
    xlSheet.Cells(1, 6).Value = "TOPLAM AGIRLIK (kg)"

    With xlSheet.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(30, 30, 30)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .WrapText = False
    End With
    xlSheet.Rows(1).RowHeight = 24

    xlSheet.Columns("A:A").ColumnWidth = COL_WIDTH_THUMB
    xlSheet.Columns("B:B").ColumnWidth = 36
    xlSheet.Columns("C:C").ColumnWidth = 22
    xlSheet.Columns("D:D").ColumnWidth = 8
    xlSheet.Columns("D:D").HorizontalAlignment = xlHAlignCenter
    xlSheet.Columns("E:F").ColumnWidth = 20

    Dim r As Long: r = 2
    Dim ii As Long, k As Variant, d As Variant, tf As String

    For ii = 1 To orderList.Count
        k = orderList(ii): d = dict(k): tf = CStr(d(4))

        xlSheet.Rows(r).RowHeight = THUMB_ROW_HEIGHT

        If Len(tf) > 0 Then
            On Error Resume Next
            If Dir(tf) <> "" Then
                Dim cellW As Double, cellH As Double, imgSz As Double, imgL As Double, imgT As Double
                cellW = xlSheet.Cells(r, 1).Width
                cellH = xlSheet.Rows(r).Height
                imgSz = cellW - (2 * THUMB_MARGIN_PT)
                If (cellH - 2 * THUMB_MARGIN_PT) < imgSz Then imgSz = cellH - (2 * THUMB_MARGIN_PT)
                If imgSz < 8 Then imgSz = 8
                imgL = xlSheet.Cells(r, 1).Left + (cellW - imgSz) / 2
                imgT = xlSheet.Cells(r, 1).Top + (cellH - imgSz) / 2

                Dim shp As Object
                Set shp = xlSheet.Shapes.AddPicture(tf, False, True, imgL, imgT, imgSz, imgSz)
                If Not shp Is Nothing Then
                    shp.LockAspectRatio = False
                    shp.Width = imgSz: shp.Height = imgSz
                    shp.Left = imgL:   shp.Top = imgT
                    shp.Placement = xlMoveAndSize
                    Set shp = Nothing
                End If
            End If
            On Error GoTo FATAL
        End If

        xlSheet.Cells(r, 2).Value = d(0)
        xlSheet.Cells(r, 3).Value = d(2)
        xlSheet.Cells(r, 4).Value = d(1)
        xlSheet.Cells(r, 5).Value = d(3)
        xlSheet.Cells(r, 6).Value = CDbl(d(3)) * CLng(d(1))

        ' Her veri satırını yatay ve dikey olarak ortala
        With xlSheet.Range("A" & r & ":F" & r)
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
        End With

        r = r + 1
    Next ii

    xlSheet.Cells(r, 5).Value = "TOPLAM:"
    xlSheet.Cells(r, 5).Font.Bold = True
    xlSheet.Cells(r, 6).Formula = "=SUM(F2:F" & (r - 1) & ")"
    xlSheet.Cells(r, 6).Font.Bold = True
    xlSheet.Cells(r, 6).Interior.Color = RGB(255, 220, 0)

    ' TOPLAM satırını ortala
    With xlSheet.Range("A" & r & ":F" & r)
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
    End With

    xlSheet.Range("E2:F" & r).NumberFormat = "0.00"
    xlSheet.Range("A1:F" & r).Borders.LineStyle = 1

    For ii = 1 To orderList.Count
        k = orderList(ii): d = dict(k): tf = CStr(d(4))
        If Len(tf) > 0 Then
            On Error Resume Next: If Dir(tf) <> "" Then Kill tf: On Error GoTo 0
        End If
    Next ii

    SetExcelCalc xlApp, xlCalculationAutomatic
    xlApp.ScreenUpdating = True
    xlApp.Visible = True
    MsgBox "OK", vbInformation
    Exit Sub

FATAL:
    On Error Resume Next
    SetExcelCalc xlApp, xlCalculationAutomatic
    If Not xlApp Is Nothing Then xlApp.ScreenUpdating = True
    MsgBox "Hata (" & Err.Number & "): " & Err.Description, vbCritical
End Sub

' --- Unsur ağacındaki TOP-LEVEL bileşenleri UI sırasıyla getir ---
Private Function GetTopLevelComponentsInTreeOrder(ByVal swModel As Object) As Variant
    Dim comps As Collection: Set comps = New Collection
    Dim swFeat As Object, typ As String, comp As Object

    Set swFeat = swModel.FirstFeature
    Do While Not swFeat Is Nothing
        typ = ""
        On Error Resume Next
        typ = swFeat.GetTypeName2
        On Error GoTo 0

        If StrComp(typ, "Reference", vbTextCompare) = 0 Then
            Set comp = Nothing
            On Error Resume Next
            Set comp = swFeat.GetSpecificFeature2
            On Error GoTo 0
            If Not comp Is Nothing Then comps.Add comp
        End If

        Set swFeat = swFeat.GetNextFeature
    Loop

    Dim arr() As Variant, i As Long
    If comps.Count > 0 Then
        ReDim arr(0 To comps.Count - 1)
        For i = 1 To comps.Count
            Set arr(i - 1) = comps(i)
        Next
    Else
        arr = Empty
    End If
    GetTopLevelComponentsInTreeOrder = arr
End Function

' --- Windows Shell thumbnail üretimi için PS1 yaz ---
Private Function WriteThumbnailScript() As String
    WriteThumbnailScript = ""
    On Error Resume Next
    Dim tmpDir As String: tmpDir = Environ("TEMP")
    If Right$(tmpDir, 1) <> "\" Then tmpDir = tmpDir & "\"
    Dim ps1Path As String: ps1Path = tmpDir & "sw_bom_thumb_extractor.ps1"

    Dim ps As String
    ps = "param([string]$InputFile, [string]$OutputFile, [int]$Size = 256)" & vbCrLf
    ps = ps & "Add-Type @'" & vbCrLf
    ps = ps & "using System;" & vbCrLf
    ps = ps & "using System.Drawing;" & vbCrLf
    ps = ps & "using System.Drawing.Imaging;" & vbCrLf
    ps = ps & "using System.Runtime.InteropServices;" & vbCrLf
    ps = ps & "public class SwThumb {" & vbCrLf
    ps = ps & "    [DllImport(""shell32.dll"",CharSet=CharSet.Unicode,PreserveSig=false)]" & vbCrLf
    ps = ps & "    private static extern void SHCreateItemFromParsingName(string p,IntPtr b,ref Guid r,out IntPtr v);" & vbCrLf
    ps = ps & "    [ComImport,Guid(""bcc18b79-ba16-442f-80c4-8a59c30c463b""),InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]" & vbCrLf
    ps = ps & "    private interface IFactory{void GetImage(SIZE s,int f,out IntPtr h);}" & vbCrLf
    ps = ps & "    [StructLayout(LayoutKind.Sequential)]" & vbCrLf
    ps = ps & "    private struct SIZE{public int cx,cy;public SIZE(int x,int y){cx=x;cy=y;}}" & vbCrLf
    ps = ps & "    public static bool Save(string path,string output,int size){" & vbCrLf
    ps = ps & "        try{" & vbCrLf
    ps = ps & "            Guid g=typeof(IFactory).GUID;IntPtr p;" & vbCrLf
    ps = ps & "            SHCreateItemFromParsingName(path,IntPtr.Zero,ref g,out p);" & vbCrLf
    ps = ps & "            var f=(IFactory)Marshal.GetObjectForIUnknown(p);" & vbCrLf
    ps = ps & "            IntPtr h;f.GetImage(new SIZE(size,size),0,out h);" & vbCrLf
    ps = ps & "            using(var bmp=Image.FromHbitmap(h)){" & vbCrLf
    ps = ps & "                using(var wb=new Bitmap(size,size)){" & vbCrLf
    ps = ps & "                    using(var g2=Graphics.FromImage(wb)){" & vbCrLf
    ps = ps & "                        g2.Clear(Color.White);" & vbCrLf
    ps = ps & "                        g2.DrawImage(bmp,0,0,size,size);" & vbCrLf
    ps = ps & "                    }wb.Save(output,ImageFormat.Bmp);}" & vbCrLf
    ps = ps & "            }return true;" & vbCrLf
    ps = ps & "        }catch{return false;}}" & vbCrLf
    ps = ps & "}" & vbCrLf
    ps = ps & "'@  -ReferencedAssemblies System.Drawing" & vbCrLf
    ps = ps & "[SwThumb]::Save($InputFile, $OutputFile, $Size) | Out-Null" & vbCrLf

    Dim fNum As Integer: fNum = FreeFile()
    Open ps1Path For Output As #fNum
    Print #fNum, ps
    Close #fNum

    If Dir(ps1Path) <> "" Then WriteThumbnailScript = ps1Path
    On Error GoTo 0
End Function

Private Function GetShellThumbnail(ByVal partPath As String, ByVal uid As Long) As String
    GetShellThumbnail = ""
    If Len(partPath) = 0 Then Exit Function
    If Len(m_PS1Path) = 0 Then Exit Function

    On Error Resume Next
    Dim tmpDir As String: tmpDir = Environ("TEMP")
    If Right$(tmpDir, 1) <> "\" Then tmpDir = tmpDir & "\"
    Dim outBMP As String: outBMP = tmpDir & "sw_bom_v12_" & uid & ".bmp"
    If Dir(outBMP) <> "" Then Kill outBMP

    Dim cmd As String
    cmd = "powershell.exe -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & _
          m_PS1Path & """ """ & partPath & """ """ & outBMP & """ 256"

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 0, True

    If Dir(outBMP) <> "" Then
        If FileLen(outBMP) > 1000 Then GetShellThumbnail = outBMP _
        Else Kill outBMP
    End If
    On Error GoTo 0
End Function

' --- Kütle (kg): montaj → parça → SW‑Mass evaluated ---
Private Function GetMass_KG(ByVal swComp As Object, _
                            ByVal swModelDoc As Object, _
                            ByVal cfg As String) As Double
    Dim m As Double, v As Variant, prevCfg As String
    GetMass_KG = 0#

    On Error Resume Next
    v = swComp.GetMassProperties2(True)
    On Error GoTo 0
    If IsArray(v) Then
        If UBound(v) >= 5 Then
            m = 0#: On Error Resume Next: m = CDbl(v(5)): On Error GoTo 0
            If m > 0# Then GetMass_KG = Round(m, 2): Exit Function
        End If
    End If

    If swModelDoc Is Nothing Then GoTo try_sw_mass
    If swModelDoc.GetType <> swDocPART Then GoTo try_sw_mass

    prevCfg = ""
    On Error Resume Next
    prevCfg = swModelDoc.GetActiveConfigurationName
    If Len(cfg) > 0 Then If StrComp(prevCfg, cfg, vbTextCompare) <> 0 Then swModelDoc.ShowConfiguration2 cfg
    On Error GoTo 0
    swModelDoc.EditRebuild3

    Dim mp As Object
    Set mp = swModelDoc.Extension.CreateMassProperty
    If Not mp Is Nothing Then
        mp.UseSystemUnits = True
        On Error Resume Next: mp.IncludeHiddenBodies = True: On Error GoTo 0
        m = 0#: On Error Resume Next: m = CDbl(mp.Mass): On Error GoTo 0
        If m > 0# Then GetMass_KG = Round(m, 2): GoTo restore_cfg
    End If

    swModelDoc.ForceRebuild3 True
    Set mp = swModelDoc.Extension.CreateMassProperty
    If Not mp Is Nothing Then
        mp.UseSystemUnits = True
        On Error Resume Next: mp.IncludeHiddenBodies = True: On Error GoTo 0
        m = 0#: On Error Resume Next: m = CDbl(mp.Mass): On Error GoTo 0
        If m > 0# Then GetMass_KG = Round(m, 2): GoTo restore_cfg
    End If

try_sw_mass:
    m = GetMassFromSWMassPropertyKG(swModelDoc, cfg)
    If m > 0# Then GetMass_KG = m

restore_cfg:
    On Error Resume Next
    If Len(prevCfg) > 0 And Len(cfg) > 0 Then
        If StrComp(prevCfg, cfg, vbTextCompare) <> 0 Then swModelDoc.ShowConfiguration2 prevCfg
    End If
    On Error GoTo 0
End Function

Private Function GetMassFromSWMassPropertyKG(ByVal modelDoc As Object, ByVal cfg As String) As Double
    Dim mgr As Object, raw As String, res As String
    GetMassFromSWMassPropertyKG = 0#
    On Error Resume Next
    If modelDoc Is Nothing Then Exit Function
    Set mgr = modelDoc.Extension.CustomPropertyManager(cfg)
    If mgr Is Nothing Then Exit Function
    raw = "": res = ""
    mgr.Get4 "SW-Mass", False, raw, res
    GetMassFromSWMassPropertyKG = ParseMassTextToKg(res)
End Function

Private Function ParseMassTextToKg(ByVal txt As String) As Double
    Dim s As String, n As Double, u As String, ci As Long, i As Long, j As Long
    ParseMassTextToKg = 0#
    s = LCase$(Trim$(txt))
    If Len(s) = 0 Then Exit Function
    s = Replace(s, ",", ".")
    i = 0: j = 0
    For ci = 1 To Len(s)
        If Mid$(s, ci, 1) Like "[0-9.]" Then
            i = ci: j = ci
            Do While j <= Len(s) And Mid$(s, j, 1) Like "[0-9.]": j = j + 1: Loop
            Exit For
        End If
    Next ci
    If i = 0 Then Exit Function
    n = Val(Mid$(s, i, j - i))
    If n <= 0# Then Exit Function
    u = Trim$(Mid$(s, j))
    If InStr(u, "kg") > 0 Then ParseMassTextToKg = n _
    ElseIf InStr(u, "g") > 0 Then ParseMassTextToKg = n / 1000# _
    ElseIf InStr(u, "lb") > 0 Then ParseMassTextToKg = n * 0.45359237 _
    Else: ParseMassTextToKg = n
    End If
End Function

' --- Malzeme ---
Private Function GetMaterialSafe(ByVal modelDoc As Object, ByVal cfg As String) As String
    Dim mat As String, db As String
    On Error Resume Next
    mat = modelDoc.GetMaterialPropertyName2(cfg, db)
    If Len(Trim$(mat)) > 0 Then GetMaterialSafe = Trim$(mat): Exit Function
    If modelDoc.GetType = swDocPART Then
        Dim vB As Variant, j As Long, swB As Object, bM As String, bD As String
        vB = modelDoc.GetBodies2(0, True)
        If IsArray(vB) Then
            For j = LBound(vB) To UBound(vB)
                Set swB = vB(j)
                If Not swB Is Nothing Then
                    bM = swB.GetMaterialPropertyName2(cfg, bD)
                    If Len(Trim$(bM)) > 0 Then GetMaterialSafe = Trim$(bM): Exit Function
                End If
            Next j
        End If
    End If
    GetMaterialSafe = "-"
End Function

' --- Yardımcılar ---
Private Function SafeIsSuppressed(ByVal c As Object) As Boolean
    SafeIsSuppressed = False: If c Is Nothing Then Exit Function
    On Error Resume Next: SafeIsSuppressed = CBool(c.IsSuppressed)
End Function

Private Function SafeIsEnvelope(ByVal c As Object) As Boolean
    SafeIsEnvelope = False: If c Is Nothing Then Exit Function
    On Error Resume Next: SafeIsEnvelope = CBool(c.IsEnvelope)
End Function

Private Function SafeExcludeFromBOM(ByVal c As Object) As Boolean
    SafeExcludeFromBOM = False: If c Is Nothing Then Exit Function
    On Error Resume Next: SafeExcludeFromBOM = CBool(c.ExcludeFromBOM)
End Function

Private Function GetCleanName(ByVal n As String) As String
    Dim s As String: s = Trim$(n)
    Dim p As Long: p = InStr(s, "/"): If p > 0 Then s = Left$(s, p - 1)
    Dim d As Long: d = InStrRev(s, "-")
    If d > 0 Then If IsNumeric(Mid$(s, d + 1)) Then s = Left$(s, d - 1)
    GetCleanName = Trim$(s)
End Function

Private Function GetExcelCalc(ByVal xlApp As Object) As Long
    On Error Resume Next: GetExcelCalc = xlApp.Calculation
    If Err.Number <> 0 Then GetExcelCalc = xlCalculationAutomatic
End Function

Private Sub SetExcelCalc(ByVal xlApp As Object, ByVal modeVal As Long)
    On Error Resume Next: xlApp.Calculation = modeVal
End Sub
