Option Explicit On
Option Strict Off
Imports ESRI.ArcGIS.esriSystem 'Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.Geodatabase



Public Class Statistikos
    Inherits System.Windows.Forms.Form
    Private Sub Statistikos_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim pAoInitialize As IAoInitialize
        Dim eLicenseStatus As esriLicenseStatus
        Dim ss As String = Today()
        pAoInitialize = New AoInitialize
        eLicenseStatus = pAoInitialize.Initialize(esriLicenseProductCode.esriLicenseProductCodeStandard)
        If eLicenseStatus <> esriLicenseStatus.esriLicenseCheckedOut Then
            MsgBox("License init failure...")
            End
        End If

        ' All of your code goes here

        ' When the application shuts down...
        pAoInitialize.Shutdown()

        ComboBox2.Text = ""
        RadioButton4.Enabled = True
        sMenuo = Kaire(ss, 7)
        TextBox1.Text = sMenuo

        If RadioButton1.Checked And RadioButton3.Checked Then
            GroupBox4.Enabled = True
        End If


        Me.Top = CType((System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height - Me.Height) / 2, Integer)
        Me.Left = CType((System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width - Me.Width) / 2, Integer)
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim i As Integer
        Dim err As String = ""
        ComboBox1.Items.Clear()
        ComboBox3.Items.Clear()
        cc = Rajonai(Kaire(ComboBox2.Text, 2))
        sFilialas = Kaire(ComboBox2.Text, 2)
        If sFilialas = "  " Then
            Button1.Enabled = False
            ComboBox1.Enabled = False
            ComboBox3.Enabled = False
        Else
            Button1.Enabled = True
            If cc(0).sKodas = "  " Then
                Button1.Enabled = True
                ComboBox1.Text = ""
                ComboBox3.Items.Clear()
                ComboBox1.Items.Clear()
                ComboBox1.Enabled = True
                ComboBox1.Items.Add("   Visi operatoriai")
                Operatoriai = regCreate.get_vart(0, 0, err, iOperatoriuSk, "")
                For i = 0 To iOperatoriuSk - 1
                    ComboBox1.Items.Add(Operatoriai(i))
                Next
                ''Operatoriai = regCreate.get_vart(CType(Kaire(ComboBox2.Text, 2), Integer), 0, err, iOperatoriuSk, "")

                ''If CType(Kaire(ComboBox2.Text, 2), Integer) = 80 Then
                ''    Dim iOperatoriuSk2 As Integer
                ''    Operatoriai2 = regCreate.get_vart(97, 0, err, iOperatoriuSk2, "")
                ''    For i = 0 To iOperatoriuSk2 - 1
                ''        ComboBox1.Items.Add(Operatoriai2(i))
                ''    Next
                ''End If
                ''For i = 0 To iOperatoriuSk - 1
                ''    ReDim Preserve Operatoriai(UBound(Operatoriai) + 1)
                ''    Operatoriai(UBound(Operatoriai)) = Operatoriai2(i)
                ''Next






            Else
                If RadioButton2.Checked Then
                    ComboBox3.Text = "   Visi rajonai"
                    ComboBox3.Items.Clear()
                    ComboBox3.Enabled = True
                    ComboBox3.Items.Add("   Visi rajonai")
                Else
                    ComboBox3.Text = "   Visi rajonai"
                    ComboBox3.Items.Clear()
                    ComboBox3.Enabled = True
                    ComboBox3.Items.Add("   Visi rajonai")
                    For i = 0 To UBound(cc) - 1
                        ComboBox3.Items.Add(cc(i).sKodas & "  " & cc(i).sRajonas)
                    Next i
                    sRajonas = Kaire(ComboBox3.Text, 2)
                End If
                ComboBox1.Enabled = True
                'pagal filiala dedu operatorius
                'If RadioButton1.Checked Then
                ComboBox1.Text = ""
                ComboBox1.Items.Clear()
                ComboBox1.Enabled = True
                ComboBox1.Items.Add("   Visi operatoriai")
                Operatoriai = regCreate.get_vart(CType(Kaire(ComboBox2.Text, 2), Integer), 0, err, iOperatoriuSk, "")
                For i = 0 To iOperatoriuSk - 1
                    ComboBox1.Items.Add(Operatoriai(i))
                Next
                If CType(Kaire(ComboBox2.Text, 2), Integer) = 80 Then
                    Dim iOperatoriuSk2 As Integer
                    Operatoriai2 = regCreate.get_vart(97, 0, err, iOperatoriuSk2, "")
                    For i = 0 To iOperatoriuSk2 - 1
                        ComboBox1.Items.Add(Operatoriai2(i))
                    Next
                    For i = 0 To iOperatoriuSk2 - 1
                        ReDim Preserve Operatoriai(UBound(Operatoriai) + 1)
                        Operatoriai(UBound(Operatoriai)) = Operatoriai2(i)
                    Next
                    iOperatoriuSk = iOperatoriuSk + iOperatoriuSk2
                End If
                
                'End If
            End If
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        sRajonas = Kaire(ComboBox3.Text, 2)
        sRajonasPavad = Vidurys(ComboBox3.Text, 4, Len(Trim(ComboBox3.Text)))
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        sOperatorius = ComboBox1.Text
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim i, ii As Integer
        '   Dim sOperator As String

        Dim iSumaGeod As Integer = 0 'parcels
        Dim iSumaGeodGeoM As Integer = 0 'parcels Geomatininkas
        Dim iSumaGeodKadaGIS As Integer = 0 'parcels KadaGIS
        Dim iSumaPrel As Integer = 0 'parcels

        Dim iSumaPatikra_2 As Integer = 0 'temp KadaGIS
        Dim iSumaIsvada_1 As Integer = 0 'temp KadaGIS
        Dim iSumaLaikini_4 As Integer = 0 'temp KadaGIS
        Dim iSumaIsvadaPatikra_3 As Integer = 0 'temp KadaGIS
        Dim iSumaPatikraNZT_8 As Integer = 0 'temp KadaGIS
        Dim iSumaIsvadaNZT_9 As Integer = 0 'temp KadaGIS
        Dim iSumaPatikraKons_12 As Integer = 0 'temp KadaGIS
        Dim iSumaIsvadaKons_13 As Integer = 0 'temp KadaGIS

        Dim iSumaLaikiniGeom_5 As Integer = 0 'temp Geomatininkas
        Dim iSumaPatikraGeoM_6 As Integer = 0 'temp Geomatininkas
        Dim iSumaIsvadaGeoM_7 As Integer = 0 'temp Geomatininkas
        Dim iSumaPatikraNZTGeoM_10 As Integer = 0 'temp Geomatininkas
        Dim iSumaIsvadaNZTGeoM_11 As Integer = 0 'temp Geomatininkas
        Dim iSumaPatikraKonsGeoM_14 As Integer = 0 'temp Geomatininkas
        Dim iSumaIsvadaKonsGeoM_15 As Integer = 0 'temp Geomatininkas

        Dim iSumaVisi As Long = 0 'visi Sklypai
        Dim iSumaTaskai As Long = 0 'visi geodeziniai taskai

        Dim iSumaSuma As Integer

        Dim iSumaInz1, iSumaInz2, iSumaInz1Laik, iSumaInz2Laik As Integer
        Dim pTable As ITable
        Dim sKelias As String = ""
        Dim sVardas As String = ""
        Dim pCursor As ICursor
        Dim pRowBuffer As IRowBuffer
        Dim sName As String

        If ComboBox2.Text = "" Then : MsgBox("Parinkite filialà") : Exit Sub : End If
        sDataNuo = Replace(DateTimePicker1.Text, ".", "-")
        sDataIki = Replace(DateTimePicker2.Text, ".", "-")

        bOperator = RadioButton3.Checked
        If bOperator Then : sName = "Operator" : Else : sName = "Rajonas" : End If

        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Windows.Forms.Cursor.Current = Cursors.WaitCursor

        If RadioButton1.Checked Then 'sklypai
            If bOperator Then
                If RadioButton5.Checked Then 'paieska pagal varda ir pavarde
                    Paieska_sklypai_oper(sFilialas, sRajonas, sDataNuo, sDataIki, sOperatorius, True)
                Else 'paieska pagal pavarde mazosiomis ir didziosiomis raidemis
                    Paieska_sklypai_oper(sFilialas, sRajonas, sDataNuo, sDataIki, sOperatorius, False)
                End If
            Else
                '  RadioButton4(sFilialas, sRajonas, sDataNuo, sDataIki, sOperatorius)
                Paieska_sklypai_raj(sFilialas, sRajonas, sDataNuo, sDataIki, sOperatorius)
            End If
            'sklypu statistika
            For i = 0 To UBound(SklypuStatistika) - 1
                iSumaGeod = iSumaGeod + SklypuStatistika(i).iGeod
                iSumaGeodGeoM = iSumaGeod + SklypuStatistika(i).iGeodGeoM
                iSumaGeodKadaGIS = iSumaGeod + SklypuStatistika(i).iGeodKadaGIS
                iSumaPrel = iSumaPrel + SklypuStatistika(i).iPrel
                iSumaPatikra_2 = iSumaPatikra_2 + SklypuStatistika(i).iPatikra_2
                iSumaIsvada_1 = iSumaIsvada_1 + SklypuStatistika(i).iIsvada_1
                iSumaLaikini_4 = iSumaLaikini_4 + SklypuStatistika(i).iLaikini_4
                iSumaIsvadaPatikra_3 = iSumaIsvadaPatikra_3 + SklypuStatistika(i).iIsvadaPatikra_3
                iSumaPatikraNZT_8 = iSumaPatikraNZT_8 + SklypuStatistika(i).iPatikraNZT_8
                iSumaIsvadaNZT_9 = iSumaIsvadaNZT_9 + SklypuStatistika(i).iIsvadaNZT_9
                iSumaPatikraKons_12 = iSumaPatikraKons_12 + SklypuStatistika(i).iPatikraKons_12
                iSumaIsvadaKons_13 = iSumaIsvadaKons_13 + SklypuStatistika(i).iIsvadaKons_13
                iSumaLaikiniGeom_5 = iSumaLaikiniGeom_5 + SklypuStatistika(i).iLaikiniGeom_5
                iSumaPatikraGeoM_6 = iSumaPatikraGeoM_6 + SklypuStatistika(i).iPatikraGeoM_6
                iSumaIsvadaGeoM_7 = iSumaIsvadaGeoM_7 + SklypuStatistika(i).iIsvadaGeoM_7
                iSumaPatikraNZTGeoM_10 = iSumaPatikraNZTGeoM_10 + SklypuStatistika(i).iPatikraNZTGeoM_10
                iSumaIsvadaNZTGeoM_11 = iSumaIsvadaNZTGeoM_11 + SklypuStatistika(i).iIsvadaNZTGeoM_11
                iSumaPatikraKonsGeoM_14 = iSumaPatikraKonsGeoM_14 + SklypuStatistika(i).iPatikraKonsGeoM_14
                iSumaIsvadaKonsGeoM_15 = iSumaIsvadaKonsGeoM_15 + SklypuStatistika(i).iIsvadaKonsGeoM_15
                iSumaVisi = iSumaVisi + SklypuStatistika(i).iSuma
                iSumaTaskai = iSumaTaskai + SklypuStatistika(i).iTaskai
            Next i
            Dim saveFileDialog11 As New SaveFileDialog
            saveFileDialog11.Title = "Iðsaugoti sklypø statistikà *.dbf faile"
            saveFileDialog11.Filter = "*.dbf|*.dbf"
            If saveFileDialog11.ShowDialog() = Windows.Forms.DialogResult.OK Then
                If Dir(saveFileDialog11.FileName) <> "" Then : Kill(saveFileDialog11.FileName) : End If
                ii = InStrRev(saveFileDialog11.FileName, "\", -1, CompareMethod.Binary)
                sKelias = Kaire(saveFileDialog11.FileName, ii - 1)
                sVardas = Vidurys(saveFileDialog11.FileName, ii + 1, Len(saveFileDialog11.FileName) - 4 - ii)
                pTable = createDBF(sVardas, sKelias, "S", sName, False)
                If pTable Is Nothing Then
                    MsgBox("Nepavyko sukurti " & sKelias & "/" & sVardas) : Exit Sub
                End If
                For i = 0 To UBound(SklypuStatistika) - 1
                    pCursor = pTable.Insert(True)
                    pRowBuffer = pTable.CreateRowBuffer
                    pRowBuffer.Value(pRowBuffer.Fields.FindField(sName)) = SklypuStatistika(i).sOperator
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("R_Geo_Visi")) = SklypuStatistika(i).iGeod
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("-GeoM")) = SklypuStatistika(i).iGeodGeoM
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("-KadaGIS")) = SklypuStatistika(i).iGeodKadaGIS

                    pRowBuffer.Value(pRowBuffer.Fields.FindField("R_Prelim")) = SklypuStatistika(i).iPrel
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_2")) = SklypuStatistika(i).iPatikra_2
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_1")) = SklypuStatistika(i).iIsvada_1
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_3")) = SklypuStatistika(i).iIsvadaPatikra_3
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_4")) = SklypuStatistika(i).iLaikini_4
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_8")) = SklypuStatistika(i).iPatikraNZT_8
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_9")) = SklypuStatistika(i).iIsvadaNZT_9
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_12")) = SklypuStatistika(i).iPatikraKons_12
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_13")) = SklypuStatistika(i).iIsvadaKons_13
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_5")) = SklypuStatistika(i).iLaikiniGeom_5
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_6")) = SklypuStatistika(i).iPatikraGeoM_6
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_7")) = SklypuStatistika(i).iIsvadaGeoM_7
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_10")) = SklypuStatistika(i).iPatikraNZTGeoM_10
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_11")) = SklypuStatistika(i).iIsvadaNZTGeoM_11
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_14")) = SklypuStatistika(i).iPatikraKonsGeoM_14
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_15")) = SklypuStatistika(i).iIsvadaKonsGeoM_15
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Visi_skl")) = SklypuStatistika(i).iSuma
                    'pRowBuffer.Value(pRowBuffer.Fields.FindField("Taskai_geo")) = SklypuStatistika(i).iTaskai
                    pCursor.InsertRow(pRowBuffer)
                    pCursor.Flush()
                Next i
                pCursor = pTable.Insert(True)
                pRowBuffer = pTable.CreateRowBuffer
                pRowBuffer.Value(pRowBuffer.Fields.FindField(sName)) = "VISO:"
                '  pRowBuffer.Value(pRowBuffer.Fields.FindField("R_Geodez")) = iSumaGeod
                ' pRowBuffer.Value(pRowBuffer.Fields.FindField(sName)) = SklypuStatistika(i).sOperator
                pRowBuffer.Value(pRowBuffer.Fields.FindField("R_Geo_Visi")) = iSumaGeod
                pRowBuffer.Value(pRowBuffer.Fields.FindField("-GeoM")) = 0
                pRowBuffer.Value(pRowBuffer.Fields.FindField("-KadaGIS")) = 0
                pRowBuffer.Value(pRowBuffer.Fields.FindField("R_Prelim")) = iSumaPrel
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_2")) = iSumaPatikra_2
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_1")) = iSumaIsvada_1
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_3")) = iSumaIsvadaPatikra_3
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_4")) = iSumaLaikini_4
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_8")) = iSumaPatikraNZT_8
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_9")) = iSumaIsvadaNZT_9
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_12")) = iSumaPatikraKons_12
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_13")) = iSumaIsvadaKons_13
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_5")) = iSumaLaikiniGeom_5
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_6")) = iSumaPatikraGeoM_6
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_7")) = iSumaIsvadaGeoM_7
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_10")) = iSumaPatikraNZTGeoM_10
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_11")) = iSumaIsvadaNZTGeoM_11
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_14")) = iSumaPatikraKonsGeoM_14
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_15")) = iSumaIsvadaKonsGeoM_15
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Visi_skl")) = iSumaVisi
                '  pRowBuffer.Value(pRowBuffer.Fields.FindField("Taskai_geo")) = iSumaTaskai
                pCursor.InsertRow(pRowBuffer)
                pCursor.Flush()
            End If
            If CheckBox1.Checked = True Then
                iSumaGeod = 0 : iSumaPrel = 0
                iSumaGeodGeoM = 0 : iSumaGeodKadaGIS = 0
                iSumaPatikra_2 = 0 : iSumaIsvada_1 = 0
                iSumaLaikini_4 = 0 : iSumaIsvadaPatikra_3 = 0
                iSumaPatikraNZT_8 = 0 : iSumaIsvadaNZT_9 = 0
                iSumaPatikraKons_12 = 0 : iSumaIsvadaKons_13 = 0
                iSumaLaikiniGeom_5 = 0 : iSumaPatikraGeoM_6 = 0
                iSumaIsvadaGeoM_7 = 0 : iSumaPatikraNZTGeoM_10 = 0
                iSumaIsvadaNZTGeoM_11 = 0 : iSumaPatikraKonsGeoM_14 = 0
                iSumaIsvadaKonsGeoM_15 = 0 : iSumaVisi = 0
                iSumaTaskai = 0
                sMenuo = TextBox1.Text
                If bOperator Then
                    If RadioButton5.Checked Then 'paieska pagal varda ir pavarde
                        Paieska_sklypai_oper_T_P(sFilialas, sRajonas, sMenuo, sDataNuo, sDataIki, sOperatorius, True)
                        '  Paieska_sklypai_oper(sFilialas, sRajonas, sDataNuo, sDataIki, sOperatorius) ', True)
                    Else 'paieska pagal pavarde mazosiomis ir didziosiomis raidemis
                        Paieska_sklypai_oper_T_P(sFilialas, sRajonas, sMenuo, sDataNuo, sDataIki, sOperatorius, False)
                        '  Paieska_sklypai_oper(sFilialas, sRajonas, sDataNuo, sDataIki, sOperatorius) ', False)
                    End If
                Else
                    ' Paieska_sklypai_raj(sFilialas, sRajonas, sDataNuo, sDataIki, sOperatorius)
                    Exit Sub
                End If
                'sklypu statistika
                For i = 0 To UBound(SklypuStatistika) - 1
                    iSumaGeod = iSumaGeod + SklypuStatistika(i).iGeod
                    iSumaGeodGeoM = iSumaGeodGeoM + SklypuStatistika(i).iGeodGeoM
                    iSumaGeodKadaGIS = iSumaGeodKadaGIS + SklypuStatistika(i).iGeodKadaGIS
                    iSumaPrel = iSumaPrel + SklypuStatistika(i).iPrel

                    iSumaPatikra_2 = iSumaPatikra_2 + SklypuStatistika(i).iPatikra_2
                    iSumaIsvada_1 = iSumaIsvada_1 + SklypuStatistika(i).iIsvada_1
                    iSumaLaikini_4 = iSumaLaikini_4 + SklypuStatistika(i).iLaikini_4
                    iSumaIsvadaPatikra_3 = iSumaIsvadaPatikra_3 + SklypuStatistika(i).iIsvadaPatikra_3
                    iSumaPatikraNZT_8 = iSumaPatikraNZT_8 + SklypuStatistika(i).iPatikraNZT_8
                    iSumaIsvadaNZT_9 = iSumaIsvadaNZT_9 + SklypuStatistika(i).iIsvadaNZT_9
                    iSumaPatikraKons_12 = iSumaPatikraKons_12 + SklypuStatistika(i).iPatikraKons_12
                    iSumaIsvadaKons_13 = iSumaIsvadaKons_13 + SklypuStatistika(i).iIsvadaKons_13
                    iSumaLaikiniGeom_5 = iSumaLaikiniGeom_5 + SklypuStatistika(i).iLaikiniGeom_5
                    iSumaPatikraGeoM_6 = iSumaPatikraGeoM_6 + SklypuStatistika(i).iPatikraGeoM_6
                    iSumaIsvadaGeoM_7 = iSumaIsvadaGeoM_7 + SklypuStatistika(i).iIsvadaGeoM_7
                    iSumaPatikraNZTGeoM_10 = iSumaPatikraNZTGeoM_10 + SklypuStatistika(i).iPatikraNZTGeoM_10
                    iSumaIsvadaNZTGeoM_11 = iSumaIsvadaNZTGeoM_11 + SklypuStatistika(i).iIsvadaNZTGeoM_11
                    iSumaPatikraKonsGeoM_14 = iSumaPatikraKonsGeoM_14 + SklypuStatistika(i).iPatikraKonsGeoM_14
                    iSumaIsvadaKonsGeoM_15 = iSumaIsvadaKonsGeoM_15 + SklypuStatistika(i).iIsvadaKonsGeoM_15
                    iSumaVisi = iSumaVisi + SklypuStatistika(i).iSuma
                    iSumaTaskai = iSumaTaskai + SklypuStatistika(i).iTaskai
                Next i
                pTable = createDBF(sVardas, sKelias, "S", sName, True)
                If pTable Is Nothing Then
                    MsgBox("Nepavyko sukurti " & sKelias & "/" & sVardas & "T_P") : Exit Sub
                End If
                For i = 0 To UBound(SklypuStatistika) - 1
                    pCursor = pTable.Insert(True)
                    pRowBuffer = pTable.CreateRowBuffer
                    pRowBuffer.Value(pRowBuffer.Fields.FindField(sName)) = SklypuStatistika(i).sOperator
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("A_Parcels")) = SklypuStatistika(i).iGeod
                    ''pRowBuffer.Value(pRowBuffer.Fields.FindField("R_Geo_Visi")) = SklypuStatistika(i).iGeod
                    ''pRowBuffer.Value(pRowBuffer.Fields.FindField("R_GeoM")) = SklypuStatistika(i).iGeodGeoM
                    ''pRowBuffer.Value(pRowBuffer.Fields.FindField("R_KadaGIS")) = SklypuStatistika(i).iGeodKadaGIS
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("A_Temp")) = SklypuStatistika(i).iPrel
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_2")) = SklypuStatistika(i).iPatikra_2
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_1")) = SklypuStatistika(i).iIsvada_1
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_3")) = SklypuStatistika(i).iIsvadaPatikra_3
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_4")) = SklypuStatistika(i).iLaikini_4
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_8")) = SklypuStatistika(i).iPatikraNZT_8
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_9")) = SklypuStatistika(i).iIsvadaNZT_9
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_12")) = SklypuStatistika(i).iPatikraKons_12
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_13")) = SklypuStatistika(i).iIsvadaKons_13
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_5")) = SklypuStatistika(i).iLaikiniGeom_5
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_6")) = SklypuStatistika(i).iPatikraGeoM_6
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_7")) = SklypuStatistika(i).iIsvadaGeoM_7
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_10")) = SklypuStatistika(i).iPatikraNZTGeoM_10
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_11")) = SklypuStatistika(i).iIsvadaNZTGeoM_11
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_14")) = SklypuStatistika(i).iPatikraKonsGeoM_14
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_15")) = SklypuStatistika(i).iIsvadaKonsGeoM_15
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Visi_skl")) = SklypuStatistika(i).iSuma
                    'pRowBuffer.Value(pRowBuffer.Fields.FindField("Taskai_geo")) = SklypuStatistika(i).iTaskai
                    pCursor.InsertRow(pRowBuffer)
                    pCursor.Flush()
                Next i
                pCursor = pTable.Insert(True)
                pRowBuffer = pTable.CreateRowBuffer
                pRowBuffer.Value(pRowBuffer.Fields.FindField(sName)) = "VISO:"
                pRowBuffer.Value(pRowBuffer.Fields.FindField("A_Parcels")) = iSumaGeod
                ''pRowBuffer.Value(pRowBuffer.Fields.FindField("R_Geo_Visi")) = SklypuStatistika(i).iGeod
                ''pRowBuffer.Value(pRowBuffer.Fields.FindField("R_GeoM")) = 0
                ''pRowBuffer.Value(pRowBuffer.Fields.FindField("R_KadaGIS")) = 0
                pRowBuffer.Value(pRowBuffer.Fields.FindField("A_Temp")) = iSumaPrel
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_2")) = iSumaPatikra_2
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_1")) = iSumaIsvada_1
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_3")) = iSumaIsvadaPatikra_3
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_4")) = iSumaLaikini_4
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_8")) = iSumaPatikraNZT_8
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_9")) = iSumaIsvadaNZT_9
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_12")) = iSumaPatikraKons_12
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_13")) = iSumaIsvadaKons_13
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_5")) = iSumaLaikiniGeom_5
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_6")) = iSumaPatikraGeoM_6
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_7")) = iSumaIsvadaGeoM_7
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_10")) = iSumaPatikraNZTGeoM_10
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_11")) = iSumaIsvadaNZTGeoM_11
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_14")) = iSumaPatikraKonsGeoM_14
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Dubl_15")) = iSumaIsvadaKonsGeoM_15
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Visi_skl")) = iSumaVisi
                '  pRowBuffer.Value(pRowBuffer.Fields.FindField("Taskai_geo")) = iSumaTaskai
                pCursor.InsertRow(pRowBuffer)
                pCursor.Flush()
            End If
        Else
            If bOperator Then
                Paieska_inzineriniai_oper(sFilialas, sRajonas, sDataNuo, sDataIki, sOperatorius)
            Else
                'Paieska_inzineriniai_raj(sFilialas, sRajonas, sDataNuo, sDataIki, sOperatorius)
            End If
            'sumuosiu visas eilutes
            For i = 0 To UBound(InzStatistika) - 1
                iSumaInz1 = iSumaInz1 + InzStatistika(i).iInz1
                iSumaInz2 = iSumaInz2 + InzStatistika(i).iInz2
                iSumaInz1Laik = iSumaInz1Laik + InzStatistika(i).iInz1Laikini
                iSumaInz2Laik = iSumaInz2Laik + InzStatistika(i).iInz2Laikini
                iSumaSuma = iSumaSuma + InzStatistika(i).iSuma
            Next i

            Dim saveFileDialog22 As New SaveFileDialog
            saveFileDialog22.Title = "Iðsaugoti inþineriniø statistikà *.dbf faile"
            saveFileDialog22.Filter = "*.dbf|*.dbf"
            If saveFileDialog22.ShowDialog() = Windows.Forms.DialogResult.OK Then
                If Dir(saveFileDialog22.FileName) <> "" Then : Kill(saveFileDialog22.FileName) : End If
                ii = InStrRev(saveFileDialog22.FileName, "\", -1, CompareMethod.Binary)
                sKelias = Kaire(saveFileDialog22.FileName, ii - 1)
                sVardas = Vidurys(saveFileDialog22.FileName, ii + 1, Len(saveFileDialog22.FileName) - 4 - ii)
                pTable = createDBF(sVardas, sKelias, "I", sName, False)
                For i = 0 To UBound(InzStatistika) - 1
                    pCursor = pTable.Insert(True)
                    pRowBuffer = pTable.CreateRowBuffer
                    pRowBuffer.Value(pRowBuffer.Fields.FindField(sName)) = InzStatistika(i).sOperator
                    '  MsgBox(InzStatistika(i).sOperator & "  " & pRowBuffer.Value(pRowBuffer.Fields.FindField(sName)))


                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Inz1")) = InzStatistika(i).iInz1
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Inz2")) = InzStatistika(i).iInz2
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Inz1L")) = InzStatistika(i).iInz1Laikini
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Inz2L")) = InzStatistika(i).iInz2Laikini
                    pRowBuffer.Value(pRowBuffer.Fields.FindField("Visi_inz")) = InzStatistika(i).iSuma
                    ' pRowBuffer.Value(pRowBuffer.Fields.FindField("Taskai_geo")) = InzStatistika(i).iTaskai
                    pCursor.InsertRow(pRowBuffer)
                    pCursor.Flush()
                Next i
                pCursor = pTable.Insert(True)
                pRowBuffer = pTable.CreateRowBuffer
                pRowBuffer.Value(pRowBuffer.Fields.FindField(sName)) = "VISO:"
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Inz1")) = iSumaInz1
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Inz2")) = iSumaInz2
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Inz1L")) = iSumaInz1Laik
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Inz2L")) = iSumaInz2Laik
                pRowBuffer.Value(pRowBuffer.Fields.FindField("Visi_inz")) = iSumaSuma
                '  pRowBuffer.Value(pRowBuffer.Fields.FindField("Taskai_geo")) = iSumaTaskai
                pCursor.InsertRow(pRowBuffer)
                pCursor.Flush()
            End If
        End If
        MsgBox("Statistikos suskaièiuotos")
        'Me.Close()
        Windows.Forms.Cursor.Current = Cursors.Default

    End Sub

    Private Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged
        If RadioButton4.Checked = True Then
            ComboBox1.Enabled = False
            ComboBox1.Text = ""
            GroupBox3.Enabled = False
        Else

            ComboBox1.Enabled = True
            GroupBox3.Enabled = True
        End If
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked Then

            GroupBox3.Enabled = False
            RadioButton5.Checked = False
            RadioButton6.Checked = False
            CheckBox1.Checked = False
            ComboBox3.Text = "   Visi rajonai"
            ComboBox3.Items.Clear()
            RadioButton4.Enabled = False
            ComboBox1.Items.Clear()
            ComboBox3.Items.Clear()
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            ComboBox3.Text = ""
        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged

        If RadioButton1.Checked = True Then
            GroupBox3.Enabled = True
            RadioButton4.Enabled = True
            ComboBox1.Items.Clear()
            ComboBox3.Items.Clear()
            ComboBox1.Text = ""
            ComboBox2.Text = ""
            ComboBox3.Text = ""
            If RadioButton3.Checked Then
                GroupBox4.Enabled = True
            End If
        Else
            GroupBox4.Enabled = False
        End If
    End Sub

    

    Private Sub RadioButton3_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked = False Then
            GroupBox4.Enabled = False
        Else
            If RadioButton1.Checked Then
                GroupBox4.Enabled = True
            End If
        End If
    End Sub

    Private Sub RadioButton6_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton6.CheckedChanged
        If RadioButton6.Checked Then
            RadioButton5.Checked = False
        Else
            RadioButton5.Checked = True
        End If
    End Sub

    Private Sub RadioButton5_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RadioButton5.CheckedChanged
        If RadioButton5.Checked Then
            RadioButton6.Checked = False
        Else
            RadioButton6.Checked = True
        End If
    End Sub
End Class
