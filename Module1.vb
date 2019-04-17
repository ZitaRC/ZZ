Option Strict Off
Option Explicit On 

Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.DataSourcesGDB
Imports System.Runtime.InteropServices
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.DataSourcesFile

Imports System
Imports System.IO



Module Module1
    Private m_AOLicenseInitializer As LicenseInitializer = New Statistika1061.LicenseInitializer()
    'Public regCreate As New janina.regService
    Public regCreate As New Kada.regService
    Public Operatoriai() As String
    Public Operatoriai2() As String
    Public iOperatoriuSk As Integer
    Public sRajonasPavad As String

    Public sFilialas As String
    Public sRajonas As String
    Public bSklypai As Boolean
    Public sDataNuo, sDataIki As String
    Public sOperatorius As String
    Public bOperator As Boolean
    Public sBatKelias As String

    Public pMouseCursor As IMouseCursor
    Public sMenuo As String

    Public Structure AA
        Dim sKodas As String
        Dim sRajonas As String
    End Structure
    Public Structure SS
        Dim sOperator As String

        Dim iGeod As Integer 'parcels
        Dim iGeodGeoM As Integer 'geomatininkas
        Dim iGeodKadaGIS As Integer 'kadagis
        Dim iPrel As Integer 'parcels

        Dim iPatikra_2 As Integer 'temp KadaGIS
        Dim iIsvada_1 As Integer 'temp KadaGIS
        Dim iLaikini_4 As Integer 'temp KadaGIS
        Dim iIsvadaPatikra_3 As Integer 'temp KadaGIS
        Dim iPatikraNZT_8 As Integer 'temp KadaGIS
        Dim iIsvadaNZT_9 As Integer 'temp KadaGIS
        Dim iPatikraKons_12 As Integer 'temp KadaGIS
        Dim iIsvadaKons_13 As Integer 'temp KadaGIS

        Dim iLaikiniGeom_5 As Integer 'temp Geomatininkas
        Dim iPatikraGeoM_6 As Integer 'temp Geomatininkas
        Dim iIsvadaGeoM_7 As Integer 'temp Geomatininkas
        Dim iPatikraNZTGeoM_10 As Integer 'temp Geomatininkas
        Dim iIsvadaNZTGeoM_11 As Integer 'temp Geomatininkas
        Dim iPatikraKonsGeoM_14 As Integer 'temp Geomatininkas
        Dim iIsvadaKonsGeoM_15 As Integer 'temp Geomatininkas

        Dim iSuma As Long
        Dim iTaskai As Long
    End Structure


    Public Structure DD
        Dim sOperator As String
        Dim iInz1 As Integer
        Dim iInz1Laikini As Integer
        Dim iInz2 As Integer
        Dim iInz2Laikini As Integer
        Dim iSuma As Integer
        Dim iTaskai As Integer
    End Structure

    Public cc() As AA  'rajonu kodai cia yra
    Public SklypuStatistika() As SS
    Public InzStatistika() As DD

    Public Structure Dubl_reiksmes
        Dim sDubl As String
        Dim sTekstas As String
        Dim iSavGyv As Short '1 - patikra; 2 - isvada
        Dim iStatistika As Integer
        Dim iStatistikaViso As Integer
    End Structure

    Public DublKadaGIS() As Dubl_reiksmes
    Public DublGeomat() As Dubl_reiksmes


    ''Sub main()
    ''    'ESRI License Initializer generated code.
    ''    m_AOLicenseInitializer.InitializeApplication(New esriLicenseProductCode() {esriLicenseProductCode.esriLicenseProductCodeArcEditor}, _
    ''    New esriLicenseExtensionCode() {})
    ''    ''If Not Dubl_KadaGIS() Then
    ''    ''    MsgBox("Nepavyko nusksityti KadaGIS.dbf") : Exit Sub
    ''    ''End If
    ''    ''If Not Dubl_Geomatininkas() Then
    ''    ''    MsgBox("Nepavyko nusksityti Geomatininkas.dbf") : Exit Sub
    ''    ''End If
    ''    Dim fr As New Form1
    ''    fr.ShowDialog()
    ''    'ESRI License Initializer generated code.
    ''    'Do not make any call to ArcObjects after ShutDownApplication()
    ''    m_AOLicenseInitializer.ShutdownApplication()
    ''End Sub

    Public Function Kaire(ByVal ss As String, ByVal i As Integer) As String
        Return (Left(ss, i))
    End Function
    Public Function Vidurys(ByVal ss As String, ByVal i As Integer, ByVal isk As Integer) As String
        Dim sDd As String
        Try
            sDd = Mid(ss, i, isk)
            Return sDd
        Catch ex As Exception
            MsgBox(ex.Message)
            Return ""
        Finally
        End Try
    End Function

    Public Function Rajonai(ByVal ss As String) As AA()
        Dim cc() As AA
        ReDim cc(0)
        Select Case ss
            Case "  " 'Visi
                ReDim cc(1)
                cc(0).sKodas = "  " : cc(0).sRajonas = ""
            Case "60" 'Alytaus
                ReDim cc(5)
                cc(0).sKodas = "59" : cc(0).sRajonas = "Lazdijø"
                cc(1).sKodas = "15" : cc(1).sRajonas = "Druskininkø"
                cc(2).sKodas = "38" : cc(2).sRajonas = "Varënos"
                cc(3).sKodas = "33" : cc(3).sRajonas = "Alytaus"
                cc(4).sKodas = "11" : cc(4).sRajonas = "Alytaus_m"
            Case "20" 'Kauno
                ReDim cc(7)
                cc(0).sKodas = "12" : cc(0).sRajonas = "Birðtono"
                cc(1).sKodas = "46" : cc(1).sRajonas = "Jonavos"
                cc(2).sKodas = "49" : cc(2).sRajonas = "Kaiðiadoriø"
                cc(3).sKodas = "52" : cc(3).sRajonas = "Kauno"
                cc(4).sKodas = "19" : cc(4).sRajonas = "Kauno_m"
                cc(5).sKodas = "53" : cc(5).sRajonas = "Këdainiø"
                cc(6).sKodas = "69" : cc(6).sRajonas = "Prienø"
            Case "50" 'Klaipëdos
                ReDim cc(6)
                cc(0).sKodas = "55" : cc(0).sRajonas = "Klaipëdos"
                cc(1).sKodas = "21" : cc(1).sRajonas = "Klaipëdos_m"
                cc(2).sKodas = "56" : cc(2).sRajonas = "Kretingos"
                cc(3).sKodas = "25" : cc(3).sRajonas = "Palangos_m"
                cc(4).sKodas = "88" : cc(4).sRajonas = "Ðilutës"
                cc(5).sKodas = "23" : cc(5).sRajonas = "Neringos m."
            Case "70" 'Marijampolës
                ReDim cc(5)
                cc(0).sKodas = "48" : cc(0).sRajonas = "Kalvarijø"
                cc(1).sKodas = "58" : cc(1).sRajonas = "Kazlø Rûdos"
                cc(2).sKodas = "18" : cc(2).sRajonas = "Marijampolës"
                cc(3).sKodas = "84" : cc(3).sRajonas = "Ðakiø"
                cc(4).sKodas = "34" : cc(4).sRajonas = "Vilkaviðkis"
            Case "97" 'Maþeikiø
                ReDim cc(3)
                cc(0).sKodas = "32" : cc(0).sRajonas = "Akmenës"
                cc(1).sKodas = "61" : cc(1).sRajonas = "Maþeikiø"
                cc(2).sKodas = "75" : cc(2).sRajonas = "Skuodo"
                ''cc(3).sKodas = "68" : cc(3).sRajonas = "Plungës"
                ''cc(4).sKodas = "74" : cc(4).sRajonas = "Rietavo"
                ''cc(5).sKodas = "78" : cc(5).sRajonas = "Telðiø"
            Case "35" 'Panevëþio
                ReDim cc(6)
                cc(0).sKodas = "32" : cc(0).sRajonas = "Birþø"
                cc(1).sKodas = "57" : cc(1).sRajonas = "Kupiðkio"
                cc(2).sKodas = "66" : cc(2).sRajonas = "Paneveþio"
                cc(3).sKodas = "27" : cc(3).sRajonas = "Panevëþio_m"
                cc(4).sKodas = "67" : cc(4).sRajonas = "Pasvalio"
                cc(5).sKodas = "73" : cc(5).sRajonas = "Rokiðkio"
            Case "40" 'Ðiauliø
                ReDim cc(6)
                cc(0).sKodas = "47" : cc(0).sRajonas = "Joniðkio"
                cc(1).sKodas = "54" : cc(1).sRajonas = "Kelmës"
                cc(2).sKodas = "65" : cc(2).sRajonas = "Pakruojo"
                cc(3).sKodas = "71" : cc(3).sRajonas = "Radviliðkio"
                cc(4).sKodas = "91" : cc(4).sRajonas = "Ðiauliø"
                cc(5).sKodas = "29" : cc(5).sRajonas = "Ðiauliø_m"
            Case "95" 'Tauragës
                ReDim cc(5)
                cc(0).sKodas = "94" : cc(0).sRajonas = "Jurbarko"
                cc(1).sKodas = "63" : cc(1).sRajonas = "Pagëgiø"
                cc(2).sKodas = "72" : cc(2).sRajonas = "Raseiniø"
                cc(3).sKodas = "87" : cc(3).sRajonas = "Ðilalës"
                cc(4).sKodas = "77" : cc(4).sRajonas = "Tauragës"
            Case "80" 'Telðiø
                ReDim cc(6)
                cc(0).sKodas = "68" : cc(0).sRajonas = "Plungës"
                cc(1).sKodas = "74" : cc(1).sRajonas = "Rietavo"
                cc(2).sKodas = "78" : cc(2).sRajonas = "Telðiø"
                cc(3).sKodas = "32" : cc(3).sRajonas = "Akmenës"
                cc(4).sKodas = "61" : cc(4).sRajonas = "Maþeikiø"
                cc(5).sKodas = "75" : cc(5).sRajonas = "Skuodo"
            Case "90" 'Utenos
                ReDim cc(6)
                cc(0).sKodas = "34" : cc(0).sRajonas = "Anykðèiø"
                cc(1).sKodas = "45" : cc(1).sRajonas = "Ignalinos"
                cc(2).sKodas = "62" : cc(2).sRajonas = "Molëtø"
                cc(3).sKodas = "82" : cc(3).sRajonas = "Utenos"
                cc(4).sKodas = "43" : cc(4).sRajonas = "Zarasø"
                cc(5).sKodas = "30" : cc(5).sRajonas = "Visaginas"
            Case "10" 'Vilniaus
                ReDim cc(8)
                cc(0).sKodas = "42" : cc(0).sRajonas = "Elektrënø"
                cc(1).sKodas = "85" : cc(1).sRajonas = "Ðalèininkø"
                cc(2).sKodas = "89" : cc(2).sRajonas = "Ðirvintø"
                cc(3).sKodas = "86" : cc(3).sRajonas = "Ðvenèioniø"
                cc(4).sKodas = "79" : cc(4).sRajonas = "Trakø"
                cc(5).sKodas = "81" : cc(5).sRajonas = "Ukmergës"
                cc(6).sKodas = "41" : cc(6).sRajonas = "Vilniaus"
                cc(7).sKodas = "13" : cc(7).sRajonas = "Vilniaus_m"
        End Select
        Return cc
    End Function

    Public Function Paieska_sklypai_oper_T_P(ByVal sFilialas As String, ByVal sRajonas As String, ByVal sMenuoTP As String, ByVal sDataNuo As String, ByVal sDataIki As String, ByVal sOperatorius As String, ByVal bVardasPavarde As Boolean) As Boolean

        Dim pFact As IWorkspaceFactory2 = Nothing
        Dim pMasterPropSet As IPropertySet = Nothing
        Dim pFCursorG As IFeatureCursor
        '  Dim pFCursorL As IFeatureCursor
        Dim pMasterWorkSpace As IWorkspace = Nothing
        Dim pFWorkSpace As IFeatureWorkspace = Nothing
        Dim pFClassR As IFeatureClass = Nothing
        Dim pFClassA As IFeatureClass = Nothing
        Dim pQFilt As IQueryFilter = Nothing
        Dim pSelSet As ISelectionSet = Nothing
        Dim sData As String
        Dim sDataAnul As String
        Dim sDataLaik As String
        'Anuliuoti 
        Dim iGeodAnul As Integer = 0 : Dim iPrelAnul As Integer = 0
        'KadaGIS laikini
        Dim iPatikraGeod_2 As Integer = 0  ' Dim iPatikra_2 As Integer = 0 :
        Dim iIsvadaGeod_1 As Integer = 0 '  Dim iIsvada_1 As Integer = 0 :
        : Dim iIsvadaGeoPatikra_3 As Integer = 0  ' Dim iIsvadaPatikra_3 As Integer = 0 
        Dim iLaikiniGeod_4 As Integer = 0  'Dim iLaikini_4 As Integer = 0 : 
        Dim iPatikraNZT_8 As Integer = 0 : Dim iIsvadaNZT_9 As Integer = 0
        Dim iPatikraKons_12 As Integer = 0 : Dim iIsvadaKons_13 As Integer = 0
        'Geomatininkas laikini
        Dim iLaikiniGeom_5 As Integer = 0 : Dim iPatikraGeoM_6 As Integer = 0
        Dim iIsvadaGeoM_7 As Integer = 0 : Dim iPatikraNZTGeoM_10 As Integer = 0
        Dim iIsvadaNZTGeoM_11 As Integer = 0 : Dim iPatikraKonsGeoM_14 As Integer = 0
        Dim iIsvadaKonsGeoM_15 As Integer = 0

        Dim i As Integer
        Dim pFeature As IFeature = Nothing
        Dim sRajPaieska As String = ""
        Dim pGeom As IGeometry = Nothing
        Dim pPntcol As IPointCollection = Nothing
        Dim pGeomColl As IGeometryCollection = Nothing
        Dim iTaskai As Integer = 0
        Dim op(2) As String
        Dim sSQLOperatorius As String = ""
        Dim sSQLOperatoriusAnul As String = ""

        Dim bGerai As Boolean = False

        ' Try
        On Error GoTo ERN
        ''pMasterPropSet = New PropertySet
        ''With pMasterPropSet
        ''    .SetProperty("Server", "SDEAPP")
        ''    .SetProperty("Instance", "5151")
        ''    .SetProperty("DATABASE", "")
        ''    .SetProperty("user", "adrviewer")
        ''    .SetProperty("password", "adr2008")
        ''    .SetProperty("version", "SDE.DEFAULT")
        ''End With
        ''pFact = New SdeWorkspaceFactory
        ''pMasterWorkSpace = pFact.Open(pMasterPropSet, 0)

        pMasterWorkSpace = ConnectToTransactionalVersionD("sdeapp", "sde:oracle11g:arcsdecdb.world", "adrviewer", "adr2008", "DATABASE", "SDE.DEFAULT", bGerai)  'Test'as


        pFWorkSpace = CType(pMasterWorkSpace, IFeatureWorkspace)
        sData = "DATA_R >= date '" & sDataNuo & "' and DATA_R <= date '" & sDataIki & "'"   'isrenkame laikotarpi
        sData = "DATA_R >= date '" & sDataNuo & "' and DATA_R < date '" & sDataIki & "'"   'isrenkame laikotarpi
        sDataAnul = "DATA_IKI >= date '" & sDataNuo & "' and DATA_IKI < date '" & sDataIki & "'"
        sDataLaik = " and PASTABOS LIKE '%" & sMenuoTP & "%'"

        'sklypai

        pFClassR = pFWorkSpace.OpenFeatureClass("KADA.PARCELS")
        pFClassA = pFWorkSpace.OpenFeatureClass("KADA.ANULIUOTI")
        If pFClassR Is Nothing Or pFClassR Is Nothing Then Exit Function
        ReDim SklypuStatistika(0)
        If sFilialas = "  " Then 'visa lietuva
            sRajPaieska = ""
        Else  'vienas filialas
            If sRajonas = "  " Then 'visi rajonai rajono kodai yra cc()
                sRajPaieska = "SAV_ID = '" & cc(0).sKodas & "'"
                For i = 1 To UBound(cc) - 1
                    sRajPaieska = sRajPaieska & " or SAV_ID = '" & cc(i).sKodas & "'"
                Next i
            Else
                sRajPaieska = "SAV_ID = '" & sRajonas & "'"  'vienas rajonas
            End If
        End If

        Dim sOper As String
        'Suksiu cikla per operatorius
        If Not (Trim(sOperatorius) = "" Or Trim(sOperatorius) = "Visi operatoriai") Then
            'visa lietuva
            If bVardasPavarde Then 'pilnas vardas ir pavarde
                sOper = sOperatorius
                SklypuStatistika(UBound(SklypuStatistika)).sOperator = sOper
                sSQLOperatorius = " and PASTABOS LIKE '%" & sOper & "%'" '= '" & sOper & "'"
                sSQLOperatoriusAnul = " and IVEDE = '" & sOper & "'"
            Else 'tik pavarde
                op = Split(sOperatorius, " ")
                sSQLOperatorius = " and ( PASTABOS LIKE '%" & UCase(op(1)) & "%' or PASTABOS LIKE '%" & op(1) & "%')"
                sSQLOperatoriusAnul = " and ( IVEDE LIKE '%" & UCase(op(1)) & "%' or IVEDE LIKE '%" & op(1) & "%')"
                SklypuStatistika(UBound(SklypuStatistika)).sOperator = op(1)
            End If



            ''sOper = sOperatorius
            ''SklypuStatistika(UBound(SklypuStatistika)).sOperator = sOper

            'Anuliuoti Sklypai - ANULIUOTI
            pQFilt = New QueryFilter 'Geodeziniai
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and " & sSQLOperatoriusAnul & " and " & sDataAnul & " and ( PASTABOS IS NULL OR NOT(PASTABOS LIKE '%TEMP%' ) )"
            Else
                pQFilt.WhereClause = sData & " and (" & sRajPaieska & " )" & sSQLOperatoriusAnul & " and " & sDataAnul & " and ( PASTABOS IS NULL OR NOT(PASTABOS LIKE '%TEMP%' ) )"
            End If
            pSelSet = pFClassA.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, Nothing)
            iGeodAnul = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iGeod = iGeodAnul

            pQFilt = New QueryFilter 'Preliminarus
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and " & sSQLOperatoriusAnul & " and " & sDataAnul & " and PASTABOS LIKE '%TEMP%'"
            Else
                pQFilt.WhereClause = sData & " and (" & sRajPaieska & " )" & sSQLOperatoriusAnul & " and " & sDataAnul & " and " & sDataAnul & " and PASTABOS LIKE '%TEMP%'"
            End If
            pSelSet = pFClassA.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, Nothing)
            iPrelAnul = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iPrel = iPrelAnul



            'Laikini sklypai -TEMP
            'KadaGIS
            ''pQFilt = New QueryFilter 'patikra preliminari
            ''If sRajPaieska = "" Then
            ''    pQFilt.WhereClause = sData & " and DUBL = 2" & sSQLOperatorius
            ''Else
            ''    pQFilt.WhereClause = sData & " and DUBL = 2 and (" & sRajPaieska & " ) " & sSQLOperatorius
            ''End If
            ''pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            ''iPatikra_2 = pSelSet.Count

            pQFilt = New QueryFilter 'patikra geodezine
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 2" & sSQLOperatorius & sDataLaik
            Else
                pQFilt.WhereClause = sData & " and DUBL = 2 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
            End If

            pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iPatikraGeod_2 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iPatikra_2 = iPatikraGeod_2



            ''pQFilt = New QueryFilter 'isvada preliminari registruota
            ''If sRajPaieska = "" Then
            ''    pQFilt.WhereClause = sData & " and DUBL = 1 and MATA_TIP = 2" & sSQLOperatorius
            ''Else
            ''    pQFilt.WhereClause = sData & " and DUBL = 1 and (" & sRajPaieska & " ) and MATA_TIP = 2" & sSQLOperatorius
            ''End If
            ''pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            ''iIsvada_1 = pSelSet.Count

            pQFilt = New QueryFilter 'isvada Geodezine registruota
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 1" & sSQLOperatorius & sDataLaik
            Else
                pQFilt.WhereClause = sData & " and DUBL = 1 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
            End If

            pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iIsvadaGeod_1 = pSelSet.Count

            SklypuStatistika(UBound(SklypuStatistika)).iIsvada_1 = iIsvadaGeod_1



            ''pQFilt = New QueryFilter 'isvada preliminari patikra
            ''If sRajPaieska = "" Then
            ''    pQFilt.WhereClause = sData & " and DUBL = 3 and MATA_TIP = 2" & sSQLOperatorius
            ''Else
            ''    pQFilt.WhereClause = sData & " and DUBL = 3 and (" & sRajPaieska & " ) and MATA_TIP = 2" & sSQLOperatorius
            ''End If
            ''pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            ''iIsvadaPatikra_3 = pSelSet.Count

            pQFilt = New QueryFilter 'isvada Geodezine patikra
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 3" & sSQLOperatorius & sDataLaik
            Else
                pQFilt.WhereClause = sData & " and DUBL = 3 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
            End If

            pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iIsvadaGeoPatikra_3 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iIsvadaPatikra_3 = iIsvadaGeoPatikra_3



            ''pQFilt = New QueryFilter 'laikinas prel
            ''If sRajPaieska = "" Then
            ''    pQFilt.WhereClause = sData & " and DUBL = 4 and MATA_TIP = 2" & sSQLOperatorius
            ''Else
            ''    pQFilt.WhereClause = sData & " and DUBL = 4 and (" & sRajPaieska & " ) and MATA_TIP = 2" & sSQLOperatorius
            ''End If
            ''pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            ''iLaikini_4 = pSelSet.Count

            pQFilt = New QueryFilter 'laikinas geod
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 4" & sSQLOperatorius & sDataLaik
            Else
                pQFilt.WhereClause = sData & " and DUBL = 4 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
            End If

            pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iLaikiniGeod_4 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iLaikini_4 = iLaikiniGeod_4


            pQFilt = New QueryFilter 'patikra be NZT
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 8" & sSQLOperatorius & sDataLaik
            Else
                pQFilt.WhereClause = sData & " and DUBL = 8 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
            End If

            pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iPatikraNZT_8 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iPatikraNZT_8 = iPatikraNZT_8


            pQFilt = New QueryFilter 'isvada be NZT
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 9" & sSQLOperatorius & sDataLaik
            Else
                pQFilt.WhereClause = sData & " and DUBL = 9 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
            End If

            pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iIsvadaNZT_9 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iIsvadaNZT_9 = iIsvadaNZT_9


            pQFilt = New QueryFilter 'patikra konsolidacija
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 12" & sSQLOperatorius & sDataLaik
            Else
                pQFilt.WhereClause = sData & " and DUBL = 12 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
            End If
            pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iPatikraKons_12 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iPatikraKons_12 = iPatikraKons_12


            pQFilt = New QueryFilter 'isvada konsolidacija
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 13" & sSQLOperatorius & sDataLaik
            Else
                pQFilt.WhereClause = sData & " and DUBL = 13 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
            End If

            pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iIsvadaKons_13 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iIsvadaKons_13 = iIsvadaKons_13


            'Geomatininkas
            pQFilt = New QueryFilter 'laikinas GeoMatininkas
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 5" & sSQLOperatorius & sDataLaik
            Else
                pQFilt.WhereClause = sData & " and DUBL = 5 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
            End If
            pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iLaikiniGeom_5 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iLaikiniGeom_5 = iLaikiniGeom_5


            pQFilt = New QueryFilter 'patikra GeoMatininkas
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 6" & sSQLOperatorius & sDataLaik
            Else
                pQFilt.WhereClause = sData & " and DUBL = 6 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
            End If
            pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iPatikraGeoM_6 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iPatikraGeoM_6 = iPatikraGeoM_6


            pQFilt = New QueryFilter 'isvada GeoMatininkas
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 7" & sSQLOperatorius & sDataLaik
            Else
                pQFilt.WhereClause = sData & " and DUBL = 7 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
            End If
            pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iIsvadaGeoM_7 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iIsvadaGeoM_7 = iIsvadaGeoM_7


            pQFilt = New QueryFilter 'patikra GeoMatininkas be NZT
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 10" & sSQLOperatorius & sDataLaik
            Else
                pQFilt.WhereClause = sData & " and DUBL = 10 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
            End If
            pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iPatikraNZTGeoM_10 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iPatikraNZTGeoM_10 = iPatikraNZTGeoM_10


            pQFilt = New QueryFilter 'isvada GeoMatininkas be NZT
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 11" & sSQLOperatorius & sDataLaik
            Else
                pQFilt.WhereClause = sData & " and DUBL = 11 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
            End If
            pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iIsvadaNZTGeoM_11 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iIsvadaNZTGeoM_11 = iIsvadaNZTGeoM_11

            pQFilt = New QueryFilter 'patikra GeoMatininkas konsolidacija
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 14" & sSQLOperatorius & sDataLaik
            Else
                pQFilt.WhereClause = sData & " and DUBL = 14 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
            End If
            pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iPatikraKonsGeoM_14 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iPatikraKonsGeoM_14 = iPatikraKonsGeoM_14


            pQFilt = New QueryFilter 'isvada GeoMatininkas konsolidacija
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 15" & sSQLOperatorius & sDataLaik
            Else
                pQFilt.WhereClause = sData & " and DUBL = 15 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
            End If
            pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iIsvadaKonsGeoM_15 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iIsvadaKonsGeoM_15 = iIsvadaKonsGeoM_15

            SklypuStatistika(UBound(SklypuStatistika)).iSuma = iGeodAnul + iPrelAnul + iPatikraGeod_2 + iIsvadaGeod_1 + iIsvadaGeoPatikra_3 + _
             iLaikiniGeod_4 + iPatikraNZT_8 + iIsvadaNZT_9 + iPatikraKons_12 + iIsvadaKons_13 + iLaikiniGeom_5 + iPatikraGeoM_6 + iIsvadaGeoM_7 + iPatikraNZTGeoM_10 + _
            iIsvadaNZTGeoM_11 + iPatikraKonsGeoM_14 + iIsvadaKonsGeoM_15
            '' SklypuStatistika(UBound(SklypuStatistika)).iSuma = iGeod + iPrel + iPatikra + iIsvada + iLaikini + iPatikraGeo + iIsvadaGeo + iLaikiniGeo + iIsvadaPatikra + iIsvadaGeoPatikra + iPatikraGeoM + iIsvadaGeoM + iPatikraNZT + iIsvadaNZT
            ' SklypuStatistika(UBound(SklypuStatistika)).iTaskai = iTaskai
            ReDim Preserve SklypuStatistika(UBound(SklypuStatistika) + 1)
        Else

            For i = 0 To iOperatoriuSk - 1
                'visa lietuva
                If bVardasPavarde Then 'pilnas vardas ir pavarde
                    sOper = Operatoriai(i)
                    SklypuStatistika(UBound(SklypuStatistika)).sOperator = sOper
                    ' sSQLOperatorius = " and IVEDE = '" & sOper & "'"
                    sSQLOperatorius = " and PASTABOS LIKE '%" & sOper & "%'" '= '" & sOper & "'"
                Else 'tik pavarde
                    op = Split(Operatoriai(i), " ")
                    '   sSQLOperatorius = " and ( IVEDE LIKE '%" & UCase(op(1)) & "%' or IVEDE LIKE '%" & op(1) & "%')" '" and IVEDE LIKE '%" & UCase(op(1)) & "%'"  
                    sSQLOperatorius = " and ( PASTABOS LIKE '%" & UCase(op(1)) & "%' or PASTABOS LIKE '%" & op(1) & "%')"
                    SklypuStatistika(UBound(SklypuStatistika)).sOperator = op(1)
                End If

                iGeodAnul = 0 : iPrelAnul = 0
                iPatikraGeod_2 = 0 : iIsvadaGeod_1 = 0
                iIsvadaGeoPatikra_3 = 0 : iLaikiniGeod_4 = 0
                iPatikraNZT_8 = 0 : iIsvadaNZT_9 = 0
                iPatikraKons_12 = 0 : iIsvadaKons_13 = 0
                iLaikiniGeom_5 = 0 : iPatikraGeoM_6 = 0 : iIsvadaGeoM_7 = 0
                iPatikraNZTGeoM_10 = 0 : iIsvadaNZTGeoM_11 = 0
                iPatikraKonsGeoM_14 = 0 : iIsvadaKonsGeoM_15 = 0
                iTaskai = 0

                ' ''anuliuoti ANULIUOTI
                pQFilt = New QueryFilter 'Geodeziniai
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and " & sSQLOperatorius & " and " & sDataAnul & " and ( PASTABOS IS NULL OR NOT(PASTABOS LIKE '%TEMP%' ) )"
                Else
                    pQFilt.WhereClause = sData & " and (" & sRajPaieska & " )" & sSQLOperatorius & " and " & sDataAnul & " and ( PASTABOS IS NULL OR NOT(PASTABOS LIKE '%TEMP%' ) )"
                End If
                pSelSet = pFClassA.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iGeodAnul = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iGeod = iGeodAnul

                pQFilt = New QueryFilter 'preliminarus
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and " & sSQLOperatorius & " and " & sDataAnul & " and PASTABOS LIKE '%TEMP%'"
                Else
                    pQFilt.WhereClause = sData & " and (" & sRajPaieska & " )" & sSQLOperatorius & " and " & sDataAnul & " and PASTABOS LIKE '%TEMP%'"
                End If
                pSelSet = pFClassA.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iPrelAnul = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iPrel = iPrelAnul


                'KadaGIS
                ''pQFilt = New QueryFilter 'patikra preliminari 
                ''If sRajPaieska = "" Then
                ''    pQFilt.WhereClause = sData & " and DUBL = 2 and MATA_TIP = 2" & sSQLOperatorius
                ''Else
                ''    pQFilt.WhereClause = sData & " and DUBL = 2 and (" & sRajPaieska & " ) and MATA_TIP = 2" & sSQLOperatorius
                ''End If
                ''pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                ''iPatikra_2 = pSelSet.Count

                pQFilt = New QueryFilter 'patikra geodezine
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 2" & sSQLOperatorius & sDataLaik
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 2 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
                End If
                ''  MsgBox(pQFilt.WhereClause)

                pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iPatikraGeod_2 = pSelSet.Count

                SklypuStatistika(UBound(SklypuStatistika)).iPatikra_2 = iPatikraGeod_2


                'pQFilt = New QueryFilter 'isvada preliminari registruota
                'If sRajPaieska = "" Then
                '    pQFilt.WhereClause = sData & " and DUBL = 1 and MATA_TIP = 2" & sSQLOperatorius
                'Else
                '    pQFilt.WhereClause = sData & " and DUBL = 1  and MATA_TIP = 2" & sSQLOperatorius
                'End If
                'pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                'iIsvada_1 = pSelSet.Count

                pQFilt = New QueryFilter 'isvada Geodezine registruota
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 1" & sSQLOperatorius & sDataLaik
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 1 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
                End If
                ''   MsgBox(pQFilt.WhereClause)

                pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iIsvadaGeod_1 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iIsvada_1 = iIsvadaGeod_1


                ''pQFilt = New QueryFilter 'isvada preliminari patikra
                ''If sRajPaieska = "" Then
                ''    pQFilt.WhereClause = sData & " and DUBL = 3 and MATA_TIP = 2" & sSQLOperatorius
                ''Else
                ''    pQFilt.WhereClause = sData & " and DUBL = 3 and MATA_TIP = 2" & sSQLOperatorius
                ''End If
                ''pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                ''iIsvadaPatikra_3 = pSelSet.Count

                pQFilt = New QueryFilter 'isvada Geodezine patikra
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 3" & sSQLOperatorius & sDataLaik
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 3 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
                End If

                ''   MsgBox(pQFilt.WhereClause)

                pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iIsvadaGeoPatikra_3 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iIsvadaPatikra_3 = iIsvadaGeoPatikra_3


                'pQFilt = New QueryFilter 'laikinas prel
                'If sRajPaieska = "" Then
                '    pQFilt.WhereClause = sData & " and DUBL = 4 and MATA_TIP = 2" & sSQLOperatorius
                'Else
                '    pQFilt.WhereClause = sData & " and DUBL = 4 and MATA_TIP = 2" & sSQLOperatorius
                'End If
                'pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                'iLaikini_4 = pSelSet.Count

                pQFilt = New QueryFilter 'laikinas geod
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 4" & sSQLOperatorius & sDataLaik
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 4 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
                End If

                ''     MsgBox(pQFilt.WhereClause)

                pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iLaikiniGeod_4 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iLaikini_4 = iLaikiniGeod_4

                pQFilt = New QueryFilter 'patikra be NZT
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 8" & sSQLOperatorius & sDataLaik
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 8 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
                End If

                '' MsgBox(pQFilt.WhereClause)

                pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iPatikraNZT_8 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iPatikraNZT_8 = iPatikraNZT_8


                pQFilt = New QueryFilter 'isvada be NZT
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 9" & sSQLOperatorius & sDataLaik
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 9 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
                End If
                ''   MsgBox(pQFilt.WhereClause)
                pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iIsvadaNZT_9 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iIsvadaNZT_9 = iIsvadaNZT_9


                pQFilt = New QueryFilter 'patikra konsolidacija
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 12" & sSQLOperatorius & sDataLaik
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 12 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
                End If
                ''   MsgBox(pQFilt.WhereClause)
                pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iPatikraKons_12 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iPatikraKons_12 = iPatikraKons_12


                pQFilt = New QueryFilter 'isvada konsolidacija
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 13" & sSQLOperatorius & sDataLaik
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 13 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
                End If
                ''     MsgBox(pQFilt.WhereClause)
                pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iIsvadaKons_13 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iIsvadaKons_13 = iIsvadaKons_13

                'Geomatininkas
                pQFilt = New QueryFilter 'laikinas GeoMatininkas
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 5" & sSQLOperatorius & sDataLaik
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 5 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
                End If
                ''   MsgBox(pQFilt.WhereClause)
                pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iLaikiniGeom_5 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iLaikiniGeom_5 = iLaikiniGeom_5


                pQFilt = New QueryFilter 'patikra GeoMatininkas
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 6" & sSQLOperatorius & sDataLaik
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 6 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
                End If
                ''   MsgBox(pQFilt.WhereClause)
                pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iPatikraGeoM_6 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iPatikraGeoM_6 = iPatikraGeoM_6


                pQFilt = New QueryFilter 'isvada GeoMatininkas
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 7" & sSQLOperatorius & sDataLaik
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 7 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
                End If
                ''   MsgBox(pQFilt.WhereClause)
                pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iIsvadaGeoM_7 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iIsvadaGeoM_7 = iIsvadaGeoM_7


                pQFilt = New QueryFilter 'patikra GeoMatininkas be NZT
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 10" & sSQLOperatorius & sDataLaik
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 10 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
                End If
                ''   MsgBox(pQFilt.WhereClause)
                pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iPatikraNZTGeoM_10 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iPatikraNZTGeoM_10 = iPatikraNZTGeoM_10


                pQFilt = New QueryFilter 'isvada GeoMatininkas be NZT
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 11" & sSQLOperatorius & sDataLaik
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 11 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
                End If
                '' MsgBox(pQFilt.WhereClause)
                pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iIsvadaNZTGeoM_11 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iIsvadaNZTGeoM_11 = iIsvadaNZTGeoM_11

                pQFilt = New QueryFilter 'patikra GeoMatininkas konsolidacija
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 14" & sSQLOperatorius & sDataLaik
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 14 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
                End If
                ''  MsgBox(pQFilt.WhereClause)
                pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iPatikraKonsGeoM_14 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iPatikraKonsGeoM_14 = iPatikraKonsGeoM_14


                pQFilt = New QueryFilter 'isvada GeoMatininkas konsolidacija
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 15" & sSQLOperatorius & sDataLaik
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 15 and (" & sRajPaieska & " )" & sSQLOperatorius & sDataLaik
                End If
                ''  MsgBox(pQFilt.WhereClause)
                pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iIsvadaKonsGeoM_15 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iIsvadaKonsGeoM_15 = iIsvadaKonsGeoM_15

                SklypuStatistika(UBound(SklypuStatistika)).iSuma = iGeodAnul + iPrelAnul + iPatikraGeod_2 + iIsvadaGeod_1 + iIsvadaGeoPatikra_3 + _
             iLaikiniGeod_4 + iPatikraNZT_8 + iIsvadaNZT_9 + iPatikraKons_12 + iIsvadaKons_13 + iLaikiniGeom_5 + iPatikraGeoM_6 + iIsvadaGeoM_7 + iPatikraNZTGeoM_10 + _
            iIsvadaNZTGeoM_11 + iPatikraKonsGeoM_14 + iIsvadaKonsGeoM_15
                'SklypuStatistika(UBound(SklypuStatistika)).iTaskai = iTaskai
                ReDim Preserve SklypuStatistika(UBound(SklypuStatistika) + 1)
            Next i
        End If
        Exit Function
ERN:    MsgBox(Err.Number & "  " & Err.Description)

    End Function


    Public Function Paieska_sklypai_oper(ByVal sFilialas As String, ByVal sRajonas As String, ByVal sDataNuo As String, ByVal sDataIki As String, ByVal sOperatorius As String, ByVal bVardasPavarde As Boolean) As Boolean

        Dim pFact As IWorkspaceFactory2 = Nothing
        Dim pMasterPropSet As IPropertySet = Nothing
        Dim pFCursorG As IFeatureCursor
        Dim pFCursorL As IFeatureCursor
        Dim pMasterWorkSpace As IWorkspace = Nothing
        Dim pFWorkSpace As IFeatureWorkspace = Nothing
        Dim pFClassR As IFeatureClass = Nothing
        Dim pFClassL As IFeatureClass = Nothing
        Dim pQFilt As IQueryFilter = Nothing
        Dim pSelSet As ISelectionSet = Nothing
        Dim sData As String
        'Registruoti 
        Dim iGeod As Integer = 0
        Dim iGeodGeoM As Integer = 0 : Dim iGeodKadaGIS As Integer = 0
        Dim iPrel As Integer = 0
        'KadaGIS laikini
        Dim iPatikra_2 As Integer = 0 : Dim iPatikraGeod_2 As Integer = 0
        Dim iIsvada_1 As Integer = 0 : Dim iIsvadaGeod_1 As Integer = 0
        Dim iIsvadaPatikra_3 As Integer = 0 : Dim iIsvadaGeoPatikra_3 As Integer = 0
        Dim iLaikini_4 As Integer = 0 : Dim iLaikiniGeod_4 As Integer = 0
        Dim iPatikraNZT_8 As Integer = 0
        Dim iIsvadaNZT_9 As Integer = 0
        Dim iPatikraKons_12 As Integer = 0
        Dim iIsvadaKons_13 As Integer = 0
        'Geomatininkas laikini
        Dim iLaikiniGeom_5 As Integer = 0
        Dim iPatikraGeoM_6 As Integer = 0
        Dim iIsvadaGeoM_7 As Integer = 0
        Dim iPatikraNZTGeoM_10 As Integer = 0
        Dim iIsvadaNZTGeoM_11 As Integer = 0
        Dim iPatikraKonsGeoM_14 As Integer = 0
        Dim iIsvadaKonsGeoM_15 As Integer = 0
        Dim bGerai As Boolean
        Dim i As Integer
        Dim pFeature As IFeature = Nothing
        Dim sRajPaieska As String = ""
        Dim pGeom As IGeometry = Nothing
        Dim pPntcol As IPointCollection = Nothing
        Dim pGeomColl As IGeometryCollection = Nothing
        Dim iTaskai As Integer = 0
        Dim sSQLOperatorius As String
        Dim sQlGeomat As String = "Dubl = 5 or Dubl = 6 or Dubl = 7 or dubl = 10 or dubl = 11 or dubl = 14 or dubl = 15"
        Dim sQlkadaGIS As String = "Dubl = 0 or Dubl = 4 or Dubl = 1 or dubl = 2 or dubl = 3 or dubl = 8 or dubl = 9 or dubl = 12 or dubl = 13"
        ' Try
        On Error GoTo ERN
        ''pMasterPropSet = New PropertySet
        ''With pMasterPropSet
        ''    .SetProperty("Server", "SDEAPP")
        ''    .SetProperty("Instance", "5151")
        ''    .SetProperty("DATABASE", "")
        ''    .SetProperty("user", "KADA")
        ''    .SetProperty("password", "NEZINAU8")
        ''    .SetProperty("version", "SDE.DEFAULT")
        ''End With
        ''pFact = New SdeWorkspaceFactory
        ''pMasterWorkSpace = pFact.Open(pMasterPropSet, 0)

        pMasterWorkSpace = ConnectToTransactionalVersionD("sdeapp", "sde:oracle11g:arcsdecdb.world", "adrviewer", "adr2008", "DATABASE", "SDE.DEFAULT", bGerai)  'Test'as
        If Not bGerai Then : Exit Function : End If


        pFWorkSpace = CType(pMasterWorkSpace, IFeatureWorkspace)
        sData = "DATA_R >= date '" & sDataNuo & "' and DATA_R <= date '" & sDataIki & "'"   'isrenkame laikotarpi
        'sklypai
        pFClassR = pFWorkSpace.OpenFeatureClass("KADA.PARCELS")
        pFClassL = pFWorkSpace.OpenFeatureClass("KADA.TEMP")
        If pFClassR Is Nothing Or pFClassR Is Nothing Then Exit Function
        ReDim SklypuStatistika(0)
        If sFilialas = "  " Then 'visa lietuva
            sRajPaieska = ""
        Else  'vienas filialas
            If sRajonas = "  " Then 'visi rajonai rajono kodai yra cc()
                sRajPaieska = "SAV_ID = '" & cc(0).sKodas & "'"
                For i = 1 To UBound(cc) - 1
                    sRajPaieska = sRajPaieska & " or SAV_ID = '" & cc(i).sKodas & "'"
                Next i
            Else
                sRajPaieska = "SAV_ID = '" & sRajonas & "'"  'vienas rajonas
            End If
        End If
        Dim sOper As String
        Dim op() As String

        If bVardasPavarde Then 'pilnas vardas ir pavarde
            sOper = sOperatorius
            SklypuStatistika(UBound(SklypuStatistika)).sOperator = sOper
            sSQLOperatorius = " IVEDE = '" & sOper & "'"

        Else 'tik pavarde
            op = Split(sOperatorius, " ")
            sOper = op(1)
            sSQLOperatorius = " (IVEDE LIKE '%" & UCase(sOper) & "%' or IVEDE LIKE '%" & sOper & "%')"
            '" ( IVEDE = '" & sOper & "'  or IVEDE = '" & UCase(sOper) & "' )"  ') '( PASTABOS LIKE '%" & UCase(op(1)) & "%' or PASTABOS LIKE '%" & op(1) & "%')"

            SklypuStatistika(UBound(SklypuStatistika)).sOperator = sOper
        End If




        'Suksiu cikla per operatorius
        If Not (Trim(sOperatorius) = "" Or Trim(sOperatorius) = "Visi operatoriai") Then
            sOper = sOperatorius
            SklypuStatistika(UBound(SklypuStatistika)).sOperator = sOper

            'Registruoti Sklypai - PARCELS
            pQFilt = New QueryFilter 'Geodeziniai visi
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and MATA_TIP = 1 and " & sSQLOperatorius
            Else
                pQFilt.WhereClause = sData & " and MATA_TIP = 1 and (" & sRajPaieska & " ) and " & sSQLOperatorius
            End If
            pFCursorG = pFClassR.Search(pQFilt, False)
            If Not pFCursorG Is Nothing Then
                pFeature = pFCursorG.NextFeature
                While Not pFeature Is Nothing
                    iGeod = iGeod + 1
                    pGeom = pFeature.ShapeCopy
                    pPntcol = CType(pGeom, IPointCollection)
                    pGeomColl = CType(pGeom, IGeometryCollection)
                    iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                    If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                    pFeature = pFCursorG.NextFeature
                End While
            End If
            SklypuStatistika(UBound(SklypuStatistika)).iGeod = iGeod


            pQFilt = New QueryFilter 'Geodeziniai i6 Geomatininko
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and MATA_TIP = 1 and " & sSQLOperatorius & " and ( " & sQlGeomat & " )"
            Else
                pQFilt.WhereClause = sData & " and MATA_TIP = 1 and (" & sRajPaieska & " ) and " & sSQLOperatorius & " and ( " & sQlGeomat & " )"
            End If
            pFCursorG = pFClassR.Search(pQFilt, False)
            If Not pFCursorG Is Nothing Then
                pFeature = pFCursorG.NextFeature
                While Not pFeature Is Nothing
                    iGeodGeoM = iGeodGeoM + 1
                    pGeom = pFeature.ShapeCopy
                    pPntcol = CType(pGeom, IPointCollection)
                    pGeomColl = CType(pGeom, IGeometryCollection)
                    iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                    If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                    pFeature = pFCursorG.NextFeature
                End While
            End If
            SklypuStatistika(UBound(SklypuStatistika)).iGeodGeoM = iGeodGeoM


            pQFilt = New QueryFilter 'Geodeziniai i6 KadaGIS
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and MATA_TIP = 1 and " & sSQLOperatorius & " and ( " & sQlkadaGIS & " )"
            Else
                pQFilt.WhereClause = sData & " and MATA_TIP = 1 and (" & sRajPaieska & " ) and " & sSQLOperatorius & " and ( " & sQlkadaGIS & " )"
            End If
            pFCursorG = pFClassR.Search(pQFilt, False)
            If Not pFCursorG Is Nothing Then
                pFeature = pFCursorG.NextFeature
                While Not pFeature Is Nothing
                    iGeodKadaGIS = iGeodKadaGIS + 1
                    pGeom = pFeature.ShapeCopy
                    pPntcol = CType(pGeom, IPointCollection)
                    pGeomColl = CType(pGeom, IGeometryCollection)
                    iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                    If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                    pFeature = pFCursorG.NextFeature
                End While
            End If
            SklypuStatistika(UBound(SklypuStatistika)).iGeodKadaGIS = iGeodKadaGIS






            pQFilt = New QueryFilter 'Preliminarus
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and MATA_TIP = 2 and " & sSQLOperatorius
            Else
                pQFilt.WhereClause = sData & " and MATA_TIP = 2 and (" & sRajPaieska & " ) and " & sSQLOperatorius
            End If
            pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iPrel = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iPrel = iPrel



            'Laikini sklypai -TEMP
            'KadaGIS
            ''pQFilt = New QueryFilter 'patikra preliminari
            ''If sRajPaieska = "" Then
            ''    pQFilt.WhereClause = sData & " and DUBL = 2 and IVEDE = '" & sOper & "' and MATA_TIP = 2"
            ''Else
            ''    pQFilt.WhereClause = sData & " and DUBL = 2 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 2"
            ''End If
            ''pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            ''iPatikra_2 = pSelSet.Count

            ''pQFilt = New QueryFilter 'patikra geodezine
            ''If sRajPaieska = "" Then
            ''    pQFilt.WhereClause = sData & " and DUBL = 2 and IVEDE = '" & sOper & "' and MATA_TIP = 1"
            ''Else
            ''    pQFilt.WhereClause = sData & " and DUBL = 2 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 1"
            ''End If
            ''pFCursorL = pFClassL.Search(pQFilt, False)
            ''If Not pFCursorL Is Nothing Then
            ''    pFeature = pFCursorL.NextFeature
            ''    While Not pFeature Is Nothing
            ''        iPatikraGeod_2 = iPatikraGeod_2 + 1
            ''        pGeom = pFeature.ShapeCopy
            ''        pPntcol = CType(pGeom, IPointCollection)
            ''        pGeomColl = CType(pGeom, IGeometryCollection)
            ''        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
            ''        If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
            ''        pFeature = pFCursorL.NextFeature
            ''    End While
            ''End If


            pQFilt = New QueryFilter 'patikra 
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 2 and " & sSQLOperatorius ' and MATA_TIP = 2"
            Else
                pQFilt.WhereClause = sData & " and DUBL = 2 and (" & sRajPaieska & " ) and " & sSQLOperatorius  ' and MATA_TIP = 2"
            End If
            pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iPatikra_2 = pSelSet.Count

            SklypuStatistika(UBound(SklypuStatistika)).iPatikra_2 = iPatikra_2 '+ iPatikraGeod_2



            ''pQFilt = New QueryFilter 'isvada preliminari registruota
            ''If sRajPaieska = "" Then
            ''    pQFilt.WhereClause = sData & " and DUBL = 1 and IVEDE = '" & sOper & "' and MATA_TIP = 2"
            ''Else
            ''    pQFilt.WhereClause = sData & " and DUBL = 1 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 2"
            ''End If
            ''pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            ''iIsvada_1 = pSelSet.Count

            ''pQFilt = New QueryFilter 'isvada Geodezine registruota
            ''If sRajPaieska = "" Then
            ''    pQFilt.WhereClause = sData & " and DUBL = 1 and IVEDE = '" & sOper & "' and MATA_TIP = 1"
            ''Else
            ''    pQFilt.WhereClause = sData & " and DUBL = 1 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 1"
            ''End If
            ''pFCursorL = pFClassL.Search(pQFilt, False)
            ''If Not pFCursorL Is Nothing Then
            ''    pFeature = pFCursorL.NextFeature
            ''    While Not pFeature Is Nothing
            ''        iIsvadaGeod_1 = iIsvadaGeod_1 + 1
            ''        pGeom = pFeature.ShapeCopy
            ''        pPntcol = CType(pGeom, IPointCollection)
            ''        pGeomColl = CType(pGeom, IGeometryCollection)
            ''        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
            ''        If Not pFeature Is Nothing Then
            ''            Marshal.ReleaseComObject(pFeature) : End If : pFeature = pFCursorL.NextFeature
            ''    End While
            ''End If
            ''SklypuStatistika(UBound(SklypuStatistika)).iIsvada_1 = iIsvada_1 + iIsvadaGeod_1


            pQFilt = New QueryFilter 'isvada  registruota
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 1 and " & sSQLOperatorius ' and MATA_TIP = 2"
            Else
                pQFilt.WhereClause = sData & " and DUBL = 1 and (" & sRajPaieska & " ) and " & sSQLOperatorius ' and MATA_TIP = 2"
            End If
            pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iIsvada_1 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iIsvada_1 = iIsvada_1


            ''pQFilt = New QueryFilter 'isvada preliminari patikra
            ''If sRajPaieska = "" Then
            ''    pQFilt.WhereClause = sData & " and DUBL = 3 and IVEDE = '" & sOper & "' and MATA_TIP = 2"
            ''Else
            ''    pQFilt.WhereClause = sData & " and DUBL = 3 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 2"
            ''End If
            ''pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            ''iIsvadaPatikra_3 = pSelSet.Count

            ''pQFilt = New QueryFilter 'isvada Geodezine patikra
            ''If sRajPaieska = "" Then
            ''    pQFilt.WhereClause = sData & " and DUBL = 3 and IVEDE = '" & sOper & "' and MATA_TIP = 1"
            ''Else
            ''    pQFilt.WhereClause = sData & " and DUBL = 3 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 1"
            ''End If
            ''pFCursorL = pFClassL.Search(pQFilt, False)
            ''If Not pFCursorL Is Nothing Then
            ''    pFeature = pFCursorL.NextFeature
            ''    While Not pFeature Is Nothing
            ''        iIsvadaGeoPatikra_3 = iIsvadaGeoPatikra_3 + 1
            ''        pGeom = pFeature.ShapeCopy
            ''        pPntcol = CType(pGeom, IPointCollection)
            ''        pGeomColl = CType(pGeom, IGeometryCollection)
            ''        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
            ''        If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
            ''        pFeature = pFCursorL.NextFeature
            ''    End While
            ''End If

            pQFilt = New QueryFilter 'isvada  patikra
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 3 and " & sSQLOperatorius ' and MATA_TIP = 2"
            Else
                pQFilt.WhereClause = sData & " and DUBL = 3 and (" & sRajPaieska & " ) and " & sSQLOperatorius ' and MATA_TIP = 2"
            End If
            pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iIsvadaPatikra_3 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iIsvadaPatikra_3 = iIsvadaPatikra_3



            ''pQFilt = New QueryFilter 'laikinas prel
            ''If sRajPaieska = "" Then
            ''    pQFilt.WhereClause = sData & " and DUBL = 4 and IVEDE = '" & sOper & "' and MATA_TIP = 2"
            ''Else
            ''    pQFilt.WhereClause = sData & " and DUBL = 4 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 2"
            ''End If
            ''pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            ''iLaikini_4 = pSelSet.Count

            ''pQFilt = New QueryFilter 'laikinas geod
            ''If sRajPaieska = "" Then
            ''    pQFilt.WhereClause = sData & " and DUBL = 4 and IVEDE = '" & sOper & "' and MATA_TIP = 1"
            ''Else
            ''    pQFilt.WhereClause = sData & " and DUBL = 4 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 1"
            ''End If
            ''pFCursorL = pFClassL.Search(pQFilt, False)
            ''If Not pFCursorL Is Nothing Then
            ''    pFeature = pFCursorL.NextFeature
            ''    While Not pFeature Is Nothing
            ''        iLaikiniGeod_4 = iLaikiniGeod_4 + 1
            ''        pGeom = pFeature.ShapeCopy
            ''        pPntcol = CType(pGeom, IPointCollection)
            ''        pGeomColl = CType(pGeom, IGeometryCollection)
            ''        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
            ''        If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
            ''        pFeature = pFCursorL.NextFeature
            ''    End While
            ''End If


            pQFilt = New QueryFilter 'laikinas 
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 4 and " & sSQLOperatorius ' and MATA_TIP = 2"
            Else
                pQFilt.WhereClause = sData & " and DUBL = 4 and (" & sRajPaieska & " ) and " & sSQLOperatorius ' and MATA_TIP = 2"
            End If
            pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iLaikini_4 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iLaikini_4 = iLaikini_4 '+ iLaikiniGeod_4




            pQFilt = New QueryFilter 'patikra be NZT
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 8 and " & sSQLOperatorius ' and MATA_TIP = 1"
            Else
                pQFilt.WhereClause = sData & " and DUBL = 8 and (" & sRajPaieska & " ) and " & sSQLOperatorius ' and MATA_TIP = 1"
            End If
            pFCursorL = pFClassL.Search(pQFilt, False)
            If Not pFCursorL Is Nothing Then
                pFeature = pFCursorL.NextFeature
                While Not pFeature Is Nothing
                    iPatikraNZT_8 = iPatikraNZT_8 + 1
                    pGeom = pFeature.ShapeCopy
                    pPntcol = CType(pGeom, IPointCollection)
                    pGeomColl = CType(pGeom, IGeometryCollection)
                    iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                    If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                    pFeature = pFCursorL.NextFeature
                End While
            End If
            SklypuStatistika(UBound(SklypuStatistika)).iPatikraNZT_8 = iPatikraNZT_8


            pQFilt = New QueryFilter 'isvada be NZT
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 9 and " & sSQLOperatorius ' and MATA_TIP = 1"
            Else
                pQFilt.WhereClause = sData & " and DUBL = 9 and (" & sRajPaieska & " ) and " & sSQLOperatorius ' and MATA_TIP = 1"
            End If
            pFCursorL = pFClassL.Search(pQFilt, False)
            If Not pFCursorL Is Nothing Then
                pFeature = pFCursorL.NextFeature
                While Not pFeature Is Nothing
                    iIsvadaNZT_9 = iIsvadaNZT_9 + 1
                    pGeom = pFeature.ShapeCopy
                    pPntcol = CType(pGeom, IPointCollection)
                    pGeomColl = CType(pGeom, IGeometryCollection)
                    iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                    If Not pFeature Is Nothing Then
                        Marshal.ReleaseComObject(pFeature)
                    End If
                    pFeature = pFCursorL.NextFeature
                End While
            End If
            SklypuStatistika(UBound(SklypuStatistika)).iIsvadaNZT_9 = iIsvadaNZT_9


            pQFilt = New QueryFilter 'patikra konsolidacija
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 12 and " & sSQLOperatorius ' and MATA_TIP = 1"
            Else
                pQFilt.WhereClause = sData & " and DUBL = 12 and (" & sRajPaieska & " ) " & sSQLOperatorius ' and MATA_TIP = 1"
            End If
            pFCursorL = pFClassL.Search(pQFilt, False)
            If Not pFCursorL Is Nothing Then
                pFeature = pFCursorL.NextFeature
                While Not pFeature Is Nothing
                    iPatikraKons_12 = iPatikraKons_12 + 1
                    pGeom = pFeature.ShapeCopy
                    pPntcol = CType(pGeom, IPointCollection)
                    pGeomColl = CType(pGeom, IGeometryCollection)
                    iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                    If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                    pFeature = pFCursorL.NextFeature
                End While
            End If
            SklypuStatistika(UBound(SklypuStatistika)).iPatikraKons_12 = iPatikraKons_12


            pQFilt = New QueryFilter 'isvada konsolidacija
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 13 and " & sSQLOperatorius ' and MATA_TIP = 1"
            Else
                pQFilt.WhereClause = sData & " and DUBL = 13 and (" & sRajPaieska & " ) and " & sSQLOperatorius ' and MATA_TIP = 1"
            End If
            pFCursorL = pFClassL.Search(pQFilt, False)
            If Not pFCursorL Is Nothing Then
                pFeature = pFCursorL.NextFeature
                While Not pFeature Is Nothing
                    iIsvadaKons_13 = iIsvadaKons_13 + 1
                    pGeom = pFeature.ShapeCopy
                    pPntcol = CType(pGeom, IPointCollection)
                    pGeomColl = CType(pGeom, IGeometryCollection)
                    iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                    If Not pFeature Is Nothing Then
                        Marshal.ReleaseComObject(pFeature)
                    End If
                    pFeature = pFCursorL.NextFeature
                End While
            End If
            SklypuStatistika(UBound(SklypuStatistika)).iIsvadaKons_13 = iIsvadaKons_13


            'Geomatininkas
            pQFilt = New QueryFilter 'laikinas GeoMatininkas
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 5 and " & sSQLOperatorius ' and MATA_TIP = 1"
            Else
                pQFilt.WhereClause = sData & " and DUBL = 5 and (" & sRajPaieska & " ) and " & sSQLOperatorius ' and MATA_TIP = 1"
            End If
            pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iLaikiniGeom_5 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iLaikiniGeom_5 = iLaikiniGeom_5


            pQFilt = New QueryFilter 'patikra GeoMatininkas
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 6 and " & sSQLOperatorius ' and MATA_TIP = 1"
            Else
                pQFilt.WhereClause = sData & " and DUBL = 6 and (" & sRajPaieska & " ) and " & sSQLOperatorius ' and MATA_TIP = 1"
            End If
            pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iPatikraGeoM_6 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iPatikraGeoM_6 = iPatikraGeoM_6


            pQFilt = New QueryFilter 'isvada GeoMatininkas
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 7 and " & sSQLOperatorius  ' and MATA_TIP = 1"
            Else
                pQFilt.WhereClause = sData & " and DUBL = 7 and (" & sRajPaieska & " ) and " & sSQLOperatorius  ' and MATA_TIP = 1"
            End If
            pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iIsvadaGeoM_7 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iIsvadaGeoM_7 = iIsvadaGeoM_7


            pQFilt = New QueryFilter 'patikra GeoMatininkas be NZT
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 10 and " & sSQLOperatorius ' and MATA_TIP = 1"
            Else
                pQFilt.WhereClause = sData & " and DUBL = 10 and (" & sRajPaieska & " ) and " & sSQLOperatorius ' and MATA_TIP = 1"
            End If
            pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iPatikraNZTGeoM_10 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iPatikraNZTGeoM_10 = iPatikraNZTGeoM_10


            pQFilt = New QueryFilter 'isvada GeoMatininkas be NZT
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 11 and " & sSQLOperatorius  ' and MATA_TIP = 1"
            Else
                pQFilt.WhereClause = sData & " and DUBL = 11 and (" & sRajPaieska & " ) and " & sSQLOperatorius  ' and MATA_TIP = 1"
            End If
            pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iIsvadaNZTGeoM_11 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iIsvadaNZTGeoM_11 = iIsvadaNZTGeoM_11

            pQFilt = New QueryFilter 'patikra GeoMatininkas konsolidacija
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 14 and " & sSQLOperatorius ' and MATA_TIP = 1"
            Else
                pQFilt.WhereClause = sData & " and DUBL = 14 and (" & sRajPaieska & " ) and " & sSQLOperatorius ' and MATA_TIP = 1"
            End If
            pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iPatikraKonsGeoM_14 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iPatikraKonsGeoM_14 = iPatikraKonsGeoM_14


            pQFilt = New QueryFilter 'isvada GeoMatininkas konsolidacija
            If sRajPaieska = "" Then
                pQFilt.WhereClause = sData & " and DUBL = 15 and " & sSQLOperatorius ' and MATA_TIP = 1"
            Else
                pQFilt.WhereClause = sData & " and DUBL = 15 and (" & sRajPaieska & " ) and " & sSQLOperatorius  ' and MATA_TIP = 1"
            End If
            pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
            iIsvadaKonsGeoM_15 = pSelSet.Count
            SklypuStatistika(UBound(SklypuStatistika)).iIsvadaKonsGeoM_15 = iIsvadaKonsGeoM_15

            SklypuStatistika(UBound(SklypuStatistika)).iSuma = iGeod + iPrel + iPatikra_2 + iPatikraGeod_2 + iIsvada_1 + iIsvadaGeod_1 + iIsvadaPatikra_3 + iIsvadaGeoPatikra_3 + _
            iLaikini_4 + iLaikiniGeod_4 + iPatikraNZT_8 + iIsvadaNZT_9 + iPatikraKons_12 + iIsvadaKons_13 + iLaikiniGeom_5 + iPatikraGeoM_6 + iIsvadaGeoM_7 + iPatikraNZTGeoM_10 + _
            iIsvadaNZTGeoM_11 + iPatikraKonsGeoM_14 + iIsvadaKonsGeoM_15
            '  SklypuStatistika(UBound(SklypuStatistika)).iSuma = iGeod + iPrel + iPatikra + iIsvada + iLaikini + iPatikraGeo + iIsvadaGeo + iLaikiniGeo + iIsvadaPatikra + iIsvadaGeoPatikra + iPatikraGeoM + iIsvadaGeoM + iPatikraNZT + iIsvadaNZT
            SklypuStatistika(UBound(SklypuStatistika)).iTaskai = iTaskai
            ReDim Preserve SklypuStatistika(UBound(SklypuStatistika) + 1)
        Else
            For i = 0 To iOperatoriuSk - 1
                'visa lietuva
                sOper = Operatoriai(i)
                SklypuStatistika(UBound(SklypuStatistika)).sOperator = sOper
                If bVardasPavarde Then 'pilnas vardas ir pavarde
                    sSQLOperatorius = " IVEDE = '" & sOper & "'"
                Else 'tik pavarde
                    op = Split(sOper, " ")
                    sOper = op(1)
                    sSQLOperatorius = " (IVEDE LIKE '%" & UCase(sOper) & "%' or IVEDE LIKE '%" & sOper & "%')"
                    ' sSQLOperatorius = " ( IVEDE = '" & sOper & "'  or IVEDE = '" & UCase(sOper) & "' )"  ') '( PASTABOS LIKE '%" & UCase(op(1)) & "%' or PASTABOS LIKE '%" & op(1) & "%')"
                End If

                iGeod = 0 : iPrel = 0
                iGeodGeoM = 0 : iGeodKadaGIS = 0
                iPatikra_2 = 0 : iPatikraGeod_2 = 0
                iIsvada_1 = 0 : iIsvadaGeod_1 = 0
                iIsvadaPatikra_3 = 0 : iIsvadaGeoPatikra_3 = 0
                iLaikini_4 = 0 : iLaikiniGeod_4 = 0
                iPatikraNZT_8 = 0 : iIsvadaNZT_9 = 0
                iPatikraKons_12 = 0 : iIsvadaKons_13 = 0
                iLaikiniGeom_5 = 0 : iPatikraGeoM_6 = 0 : iIsvadaGeoM_7 = 0
                iPatikraNZTGeoM_10 = 0 : iIsvadaNZTGeoM_11 = 0
                iPatikraKonsGeoM_14 = 0 : iIsvadaKonsGeoM_15 = 0
                iTaskai = 0

                'registruoti PARCELS
                pQFilt = New QueryFilter 'Geodeziniai visi
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and MATA_TIP = 1 and " & sSQLOperatorius
                Else
                    pQFilt.WhereClause = sData & " and MATA_TIP = 1 and (" & sRajPaieska & " ) and " & sSQLOperatorius
                End If
                pFCursorG = pFClassR.Search(pQFilt, False)
                If Not pFCursorG Is Nothing Then
                    pFeature = pFCursorG.NextFeature
                    While Not pFeature Is Nothing
                        iGeod = iGeod + 1
                        pGeom = pFeature.ShapeCopy
                        pPntcol = CType(pGeom, IPointCollection)
                        pGeomColl = CType(pGeom, IGeometryCollection)
                        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                        If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                        pFeature = pFCursorG.NextFeature
                    End While
                End If
                SklypuStatistika(UBound(SklypuStatistika)).iGeod = iGeod


                pQFilt = New QueryFilter 'Geodeziniai geomatininkas
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and MATA_TIP = 1 and " & sSQLOperatorius & " and ( " & sQlGeomat & " )"
                Else
                    pQFilt.WhereClause = sData & " and MATA_TIP = 1 and (" & sRajPaieska & " ) and " & sSQLOperatorius & " and ( " & sQlGeomat & " )"
                End If
                pFCursorG = pFClassR.Search(pQFilt, False)
                If Not pFCursorG Is Nothing Then
                    pFeature = pFCursorG.NextFeature
                    While Not pFeature Is Nothing
                        iGeodGeoM = iGeodGeoM + 1
                        pGeom = pFeature.ShapeCopy
                        pPntcol = CType(pGeom, IPointCollection)
                        pGeomColl = CType(pGeom, IGeometryCollection)
                        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                        If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                        pFeature = pFCursorG.NextFeature
                    End While
                End If
                SklypuStatistika(UBound(SklypuStatistika)).iGeodGeoM = iGeodGeoM


                pQFilt = New QueryFilter 'Geodeziniai kadaGIS
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and MATA_TIP = 1 and " & sSQLOperatorius & "' and ( " & sQlkadaGIS & " )"
                Else
                    pQFilt.WhereClause = sData & " and MATA_TIP = 1 and (" & sRajPaieska & " ) and " & sSQLOperatorius & " and ( " & sQlkadaGIS & " )"
                End If
                pFCursorG = pFClassR.Search(pQFilt, False)
                If Not pFCursorG Is Nothing Then
                    pFeature = pFCursorG.NextFeature
                    While Not pFeature Is Nothing
                        iGeodKadaGIS = iGeodKadaGIS + 1
                        pGeom = pFeature.ShapeCopy
                        pPntcol = CType(pGeom, IPointCollection)
                        pGeomColl = CType(pGeom, IGeometryCollection)
                        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                        If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                        pFeature = pFCursorG.NextFeature
                    End While
                End If
                SklypuStatistika(UBound(SklypuStatistika)).iGeodKadaGIS = iGeodKadaGIS



                pQFilt = New QueryFilter 'preliminarus
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and MATA_TIP = 2 " & sSQLOperatorius
                Else
                    pQFilt.WhereClause = sData & " and MATA_TIP = 2 and (" & sRajPaieska & " ) and " & sSQLOperatorius
                End If
                pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iPrel = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iPrel = iPrel


                'KadaGIS
                ''pQFilt = New QueryFilter 'patikra preliminari 
                ''If sRajPaieska = "" Then
                ''    pQFilt.WhereClause = sData & " and DUBL = 2 and IVEDE = '" & sOper & "' and MATA_TIP = 2"
                ''Else
                ''    pQFilt.WhereClause = sData & " and DUBL = 2 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 2"
                ''End If
                ''pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                ''iPatikra_2 = pSelSet.Count

                ''pQFilt = New QueryFilter 'patikra geodezine
                ''If sRajPaieska = "" Then
                ''    pQFilt.WhereClause = sData & " and DUBL = 2 and IVEDE = '" & sOper & "' and MATA_TIP = 1"
                ''Else
                ''    pQFilt.WhereClause = sData & " and DUBL = 2 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 1"
                ''End If
                ''pFCursorL = pFClassL.Search(pQFilt, False)
                ''If Not pFCursorL Is Nothing Then
                ''    pFeature = pFCursorL.NextFeature
                ''    While Not pFeature Is Nothing
                ''        iPatikraGeod_2 = iPatikraGeod_2 + 1
                ''        pGeom = pFeature.ShapeCopy
                ''        pPntcol = CType(pGeom, IPointCollection)
                ''        pGeomColl = CType(pGeom, IGeometryCollection)
                ''        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                ''        If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                ''        pFeature = pFCursorL.NextFeature
                ''    End While
                ''End If

                pQFilt = New QueryFilter 'patikra  
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 2 and " & sSQLOperatorius ' and MATA_TIP = 2"
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 2 and (" & sRajPaieska & " ) and " & sSQLOperatorius ' and MATA_TIP = 2"
                End If
                pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iPatikra_2 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iPatikra_2 = iPatikra_2 '+ iPatikraGeod_2


                ''pQFilt = New QueryFilter 'isvada preliminari registruota
                ''If sRajPaieska = "" Then
                ''    pQFilt.WhereClause = sData & " and DUBL = 1 and IVEDE = '" & sOper & "' and MATA_TIP = 2"
                ''Else
                ''    pQFilt.WhereClause = sData & " and DUBL = 1 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 2"
                ''End If
                ''pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                ''iIsvada_1 = pSelSet.Count

                ''pQFilt = New QueryFilter 'isvada Geodezine registruota
                ''If sRajPaieska = "" Then
                ''    pQFilt.WhereClause = sData & " and DUBL = 1 and IVEDE = '" & sOper & "' and MATA_TIP = 1"
                ''Else
                ''    pQFilt.WhereClause = sData & " and DUBL = 1 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 1"
                ''End If
                ''pFCursorL = pFClassL.Search(pQFilt, False)
                ''If Not pFCursorL Is Nothing Then
                ''    pFeature = pFCursorL.NextFeature
                ''    While Not pFeature Is Nothing
                ''        iIsvadaGeod_1 = iIsvadaGeod_1 + 1
                ''        pGeom = pFeature.ShapeCopy
                ''        pPntcol = CType(pGeom, IPointCollection)
                ''        pGeomColl = CType(pGeom, IGeometryCollection)
                ''        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                ''        If Not pFeature Is Nothing Then
                ''            Marshal.ReleaseComObject(pFeature) : End If : pFeature = pFCursorL.NextFeature
                ''    End While
                ''End If

                pQFilt = New QueryFilter 'isvada  registruota
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 1 and " & sSQLOperatorius ' and MATA_TIP = 2"
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 1 and (" & sRajPaieska & " ) and " & sSQLOperatorius ' and MATA_TIP = 2"
                End If
                pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iIsvada_1 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iIsvada_1 = iIsvada_1 '+ iIsvadaGeod_1


                ''pQFilt = New QueryFilter 'isvada preliminari patikra
                ''If sRajPaieska = "" Then
                ''    pQFilt.WhereClause = sData & " and DUBL = 3 and IVEDE = '" & sOper & "' and MATA_TIP = 2"
                ''Else
                ''    pQFilt.WhereClause = sData & " and DUBL = 3 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 2"
                ''End If
                ''pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                ''iIsvadaPatikra_3 = pSelSet.Count

                ''pQFilt = New QueryFilter 'isvada Geodezine patikra
                ''If sRajPaieska = "" Then
                ''    pQFilt.WhereClause = sData & " and DUBL = 3 and IVEDE = '" & sOper & "' and MATA_TIP = 1"
                ''Else
                ''    pQFilt.WhereClause = sData & " and DUBL = 3 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 1"
                ''End If
                ''pFCursorL = pFClassL.Search(pQFilt, False)
                ''If Not pFCursorL Is Nothing Then
                ''    pFeature = pFCursorL.NextFeature
                ''    While Not pFeature Is Nothing
                ''        iIsvadaGeoPatikra_3 = iIsvadaGeoPatikra_3 + 1
                ''        pGeom = pFeature.ShapeCopy
                ''        pPntcol = CType(pGeom, IPointCollection)
                ''        pGeomColl = CType(pGeom, IGeometryCollection)
                ''        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                ''        If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                ''        pFeature = pFCursorL.NextFeature
                ''    End While
                ''End If

                pQFilt = New QueryFilter 'isvada  patikra
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 3 and " & sSQLOperatorius ' and MATA_TIP = 2"
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 3 and (" & sRajPaieska & " ) and " & sSQLOperatorius ' and MATA_TIP = 2"
                End If
                pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iIsvadaPatikra_3 = pSelSet.Count

                SklypuStatistika(UBound(SklypuStatistika)).iIsvadaPatikra_3 = iIsvadaPatikra_3 '+ iIsvadaGeoPatikra_3


                ''pQFilt = New QueryFilter 'laikinas prel
                ''If sRajPaieska = "" Then
                ''    pQFilt.WhereClause = sData & " and DUBL = 4 and IVEDE = '" & sOper & "' and MATA_TIP = 2"
                ''Else
                ''    pQFilt.WhereClause = sData & " and DUBL = 4 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 2"
                ''End If
                ''pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                ''iLaikini_4 = pSelSet.Count

                ''pQFilt = New QueryFilter 'laikinas geod
                ''If sRajPaieska = "" Then
                ''    pQFilt.WhereClause = sData & " and DUBL = 4 and IVEDE = '" & sOper & "' and MATA_TIP = 1"
                ''Else
                ''    pQFilt.WhereClause = sData & " and DUBL = 4 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 1"
                ''End If
                ''pFCursorL = pFClassL.Search(pQFilt, False)
                ''If Not pFCursorL Is Nothing Then
                ''    pFeature = pFCursorL.NextFeature
                ''    While Not pFeature Is Nothing
                ''        iLaikiniGeod_4 = iLaikiniGeod_4 + 1
                ''        pGeom = pFeature.ShapeCopy
                ''        pPntcol = CType(pGeom, IPointCollection)
                ''        pGeomColl = CType(pGeom, IGeometryCollection)
                ''        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                ''        If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                ''        pFeature = pFCursorL.NextFeature
                ''    End While
                ''End If


                pQFilt = New QueryFilter 'laikinas 
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 4 and " & sSQLOperatorius ' and MATA_TIP = 2"
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 4 and (" & sRajPaieska & " ) and " & sSQLOperatorius ' and MATA_TIP = 2"
                End If
                pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iLaikini_4 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iLaikini_4 = iLaikini_4 '+ iLaikiniGeod_4

                pQFilt = New QueryFilter 'patikra be NZT
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 8 and " & sSQLOperatorius ' and MATA_TIP = 1"
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 8 and (" & sRajPaieska & " ) and " & sSQLOperatorius ' and MATA_TIP = 1"
                End If
                pFCursorL = pFClassL.Search(pQFilt, False)
                If Not pFCursorL Is Nothing Then
                    pFeature = pFCursorL.NextFeature
                    While Not pFeature Is Nothing
                        iPatikraNZT_8 = iPatikraNZT_8 + 1
                        pGeom = pFeature.ShapeCopy
                        pPntcol = CType(pGeom, IPointCollection)
                        pGeomColl = CType(pGeom, IGeometryCollection)
                        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                        If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                        pFeature = pFCursorL.NextFeature
                    End While
                End If
                SklypuStatistika(UBound(SklypuStatistika)).iPatikraNZT_8 = iPatikraNZT_8


                pQFilt = New QueryFilter 'isvada be NZT
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 9 and " & sSQLOperatorius ' and MATA_TIP = 1"
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 9 and (" & sRajPaieska & " ) and " & sSQLOperatorius ' and MATA_TIP = 1"
                End If
                pFCursorL = pFClassL.Search(pQFilt, False)
                If Not pFCursorL Is Nothing Then
                    pFeature = pFCursorL.NextFeature
                    While Not pFeature Is Nothing
                        iIsvadaNZT_9 = iIsvadaNZT_9 + 1
                        pGeom = pFeature.ShapeCopy
                        pPntcol = CType(pGeom, IPointCollection)
                        pGeomColl = CType(pGeom, IGeometryCollection)
                        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                        If Not pFeature Is Nothing Then
                            Marshal.ReleaseComObject(pFeature)
                        End If
                        pFeature = pFCursorL.NextFeature
                    End While
                End If
                SklypuStatistika(UBound(SklypuStatistika)).iIsvadaNZT_9 = iIsvadaNZT_9


                pQFilt = New QueryFilter 'patikra konsolidacija
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 12 and " & sSQLOperatorius ' and MATA_TIP = 1"
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 12 and (" & sRajPaieska & " ) and " & sSQLOperatorius ' and MATA_TIP = 1"
                End If
                pFCursorL = pFClassL.Search(pQFilt, False)
                If Not pFCursorL Is Nothing Then
                    pFeature = pFCursorL.NextFeature
                    While Not pFeature Is Nothing
                        iPatikraKons_12 = iPatikraKons_12 + 1
                        pGeom = pFeature.ShapeCopy
                        pPntcol = CType(pGeom, IPointCollection)
                        pGeomColl = CType(pGeom, IGeometryCollection)
                        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                        If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                        pFeature = pFCursorL.NextFeature
                    End While
                End If
                SklypuStatistika(UBound(SklypuStatistika)).iPatikraKons_12 = iPatikraKons_12


                pQFilt = New QueryFilter 'isvada konsolidacija
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 13 and " & sSQLOperatorius ' and MATA_TIP = 1"
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 13 and (" & sRajPaieska & " ) and " & sSQLOperatorius ' and MATA_TIP = 1"
                End If
                pFCursorL = pFClassL.Search(pQFilt, False)
                If Not pFCursorL Is Nothing Then
                    pFeature = pFCursorL.NextFeature
                    While Not pFeature Is Nothing
                        iIsvadaKons_13 = iIsvadaKons_13 + 1
                        pGeom = pFeature.ShapeCopy
                        pPntcol = CType(pGeom, IPointCollection)
                        pGeomColl = CType(pGeom, IGeometryCollection)
                        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                        If Not pFeature Is Nothing Then
                            Marshal.ReleaseComObject(pFeature)
                        End If
                        pFeature = pFCursorL.NextFeature
                    End While
                End If
                SklypuStatistika(UBound(SklypuStatistika)).iIsvadaKons_13 = iIsvadaKons_13

                'Geomatininkas
                pQFilt = New QueryFilter 'laikinas GeoMatininkas
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 5 and " & sSQLOperatorius ' and MATA_TIP = 1"
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 5 and (" & sRajPaieska & " ) and " & sSQLOperatorius ' and MATA_TIP = 1"
                End If
                pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iLaikiniGeom_5 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iLaikiniGeom_5 = iLaikiniGeom_5


                pQFilt = New QueryFilter 'patikra GeoMatininkas
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 6 and " & sSQLOperatorius ' and MATA_TIP = 1"
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 6 and (" & sRajPaieska & " ) and " & sSQLOperatorius ' and MATA_TIP = 1"
                End If
                pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iPatikraGeoM_6 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iPatikraGeoM_6 = iPatikraGeoM_6

                pQFilt = New QueryFilter 'isvada GeoMatininkas
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 7 and " & sSQLOperatorius  ' and MATA_TIP = 1"
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 7 and (" & sRajPaieska & " ) and " & sSQLOperatorius  ' and MATA_TIP = 1"
                End If
                pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iIsvadaGeoM_7 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iIsvadaGeoM_7 = iIsvadaGeoM_7


                pQFilt = New QueryFilter 'patikra GeoMatininkas be NZT
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 10 and " & sSQLOperatorius ' and MATA_TIP = 1"
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 10 and (" & sRajPaieska & " ) and " & sSQLOperatorius ' and MATA_TIP = 1"
                End If
                pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iPatikraNZTGeoM_10 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iPatikraNZTGeoM_10 = iPatikraNZTGeoM_10


                pQFilt = New QueryFilter 'isvada GeoMatininkas be NZT
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 11 and " & sSQLOperatorius ' and MATA_TIP = 1"
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 11 and (" & sRajPaieska & " ) and " & sSQLOperatorius  ' and MATA_TIP = 1"
                End If
                pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iIsvadaNZTGeoM_11 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iIsvadaNZTGeoM_11 = iIsvadaNZTGeoM_11

                pQFilt = New QueryFilter 'patikra GeoMatininkas konsolidacija
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 14 and " & sSQLOperatorius ' and MATA_TIP = 1"
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 14 and (" & sRajPaieska & " ) and " & sSQLOperatorius ' and MATA_TIP = 1"
                End If
                pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iPatikraKonsGeoM_14 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iPatikraKonsGeoM_14 = iPatikraKonsGeoM_14


                pQFilt = New QueryFilter 'isvada GeoMatininkas konsolidacija
                If sRajPaieska = "" Then
                    pQFilt.WhereClause = sData & " and DUBL = 15 and " & sSQLOperatorius ' and MATA_TIP = 1"
                Else
                    pQFilt.WhereClause = sData & " and DUBL = 15 and (" & sRajPaieska & " ) and " & sSQLOperatorius ' and MATA_TIP = 1"
                End If
                pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iIsvadaKonsGeoM_15 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iIsvadaKonsGeoM_15 = iIsvadaKonsGeoM_15

                SklypuStatistika(UBound(SklypuStatistika)).iSuma = iGeod + iPrel + iPatikra_2 + iPatikraGeod_2 + iIsvada_1 + iIsvadaGeod_1 + iIsvadaPatikra_3 + iIsvadaGeoPatikra_3 + _
            iLaikini_4 + iLaikiniGeod_4 + iPatikraNZT_8 + iIsvadaNZT_9 + iPatikraKons_12 + iIsvadaKons_13 + iLaikiniGeom_5 + iPatikraGeoM_6 + iIsvadaGeoM_7 + iPatikraNZTGeoM_10 + _
            iIsvadaNZTGeoM_11 + iPatikraKonsGeoM_14 + iIsvadaKonsGeoM_15
                SklypuStatistika(UBound(SklypuStatistika)).iTaskai = iTaskai
                ReDim Preserve SklypuStatistika(UBound(SklypuStatistika) + 1)
            Next i
        End If
        Exit Function
ERN:    MsgBox(Err.Number & "  " & Err.Description)

    End Function



    ' ''    Public Function Paieska_sklypai_oper(ByVal sFilialas As String, ByVal sRajonas As String, ByVal sDataNuo As String, ByVal sDataIki As String, ByVal sOperatorius As String) As Boolean

    ' ''        Dim pFact As IWorkspaceFactory2 = Nothing
    ' ''        Dim pMasterPropSet As IPropertySet = Nothing
    ' ''        Dim pFCursorG As IFeatureCursor = Nothing
    ' ''        Dim pFCursorL1 As IFeatureCursor = Nothing
    ' ''        Dim pFCursorL As IFeatureCursor = Nothing
    ' ''        Dim pFCursorL2 As IFeatureCursor = Nothing
    ' ''        Dim pFCursorL3 As IFeatureCursor = Nothing
    ' ''        Dim pMasterWorkSpace As IWorkspace = Nothing
    ' ''        Dim pFWorkSpace As IFeatureWorkspace = Nothing
    ' ''        Dim pFClassR As IFeatureClass = Nothing
    ' ''        Dim pFClassL As IFeatureClass = Nothing
    ' ''        Dim pQFilt As IQueryFilter = Nothing
    ' ''        '  Dim pQFilt2 As IQueryFilter = Nothing
    ' ''        Dim pQFiltL As IQueryFilter = Nothing
    ' ''        ''Dim pQFiltL2 As IQueryFilter = Nothing
    ' ''        ''Dim pQFiltL3 As IQueryFilter = Nothing
    ' ''        Dim pSelSet As ISelectionSet = Nothing
    ' ''        Dim sData As String
    ' ''        Dim iGeod As Integer = 0
    ' ''        Dim iPrel As Integer = 0

    ' ''        ''Dim iPatikra As Integer = 0
    ' ''        ''Dim iIsvada As Integer = 0
    ' ''        ''Dim iIsvadaPatikra As Integer = 0
    ' ''        ''Dim iLaikini As Integer = 0
    ' ''        ''Dim iPatikraGeo As Integer = 0
    ' ''        ''Dim iIsvadaGeo As Integer = 0
    ' ''        ''Dim iIsvadaGeoPatikra As Integer = 0
    ' ''        ''Dim iLaikiniGeo As Integer = 0
    ' ''        ''Dim iPatikraGeoM As Integer = 0
    ' ''        ''Dim iIsvadaGeoM As Integer = 0
    ' ''        ''Dim iPatikraNZT As Integer = 0
    ' ''        ''Dim iIsvadaNZT As Integer = 0
    ' ''        Dim i As Integer
    ' ''        Dim pFeature As IFeature
    ' ''        Dim sRajPaieska As String = ""
    ' ''        Dim pGeom As IGeometry = Nothing
    ' ''        Dim pPntcol As IPointCollection = Nothing
    ' ''        Dim pGeomColl As IGeometryCollection = Nothing
    ' ''        Dim iTaskai As Integer = 0
    ' ''        ' Try
    ' ''        On Error GoTo ERN
    ' ''        pMasterPropSet = New PropertySet
    ' ''        With pMasterPropSet
    ' ''            .SetProperty("Server", "SDEAPP")
    ' ''            .SetProperty("Instance", "5151")
    ' ''            .SetProperty("DATABASE", "")
    ' ''            .SetProperty("user", "SDE")
    ' ''            .SetProperty("password", "SDENAUJAS")
    ' ''            .SetProperty("version", "SDE.DEFAULT")
    ' ''        End With
    ' ''        pFact = New SdeWorkspaceFactory
    ' ''        pMasterWorkSpace = pFact.Open(pMasterPropSet, 0)
    ' ''        pFWorkSpace = CType(pMasterWorkSpace, IFeatureWorkspace)
    ' ''        'isrenkame laikotarpi
    ' ''        sData = "DATA_R >= date '" & sDataNuo & "' and DATA_R <= date '" & sDataIki & "'"
    ' ''        'sklypai
    ' ''        pFClassR = pFWorkSpace.OpenFeatureClass("KADA.PARCELS")
    ' ''        pFClassL = pFWorkSpace.OpenFeatureClass("KADA.TEMP")
    ' ''        If pFClassR Is Nothing Or pFClassR Is Nothing Then Exit Function
    ' ''        ReDim SklypuStatistika(0)
    ' ''        If sFilialas = "  " Then  'visa lietuva
    ' ''            sRajPaieska = ""
    ' ''        Else   'vienas filialas
    ' ''            If sRajonas = "  " Then   'visi rajonai rajono kodai yra cc()
    ' ''                sRajPaieska = "SAV_ID = '" & cc(0).sKodas & "'"
    ' ''                For i = 1 To UBound(cc) - 1
    ' ''                    sRajPaieska = sRajPaieska & " or SAV_ID = '" & cc(i).sKodas & "'"
    ' ''                Next i
    ' ''            Else  'vienas rajonas
    ' ''                sRajPaieska = "SAV_ID = '" & sRajonas & "'"
    ' ''            End If
    ' ''        End If

    ' ''        Dim sOper As String
    ' ''        'Suksiu cikla per operatorius
    ' ''        If Not (Trim(sOperatorius) = "" Or Trim(sOperatorius) = "Visi operatoriai") Then
    ' ''            sOper = sOperatorius
    ' ''            SklypuStatistika(UBound(SklypuStatistika)).sOperator = sOper

    ' ''            pQFilt = New QueryFilter 'Geodeziniai 
    ' ''            If sRajPaieska = "" Then
    ' ''                pQFilt.WhereClause = sData & " and MATA_TIP = 1 and IVEDE = '" & sOper & "'"
    ' ''            Else
    ' ''                pQFilt.WhereClause = sData & " and MATA_TIP = 1 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "'"
    ' ''            End If
    ' ''            pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
    ' ''            SklypuStatistika(UBound(SklypuStatistika)).iGeod = pSelSet.Count

    ' ''            pQFilt = New QueryFilter 'Preliminarus
    ' ''            If sRajPaieska = "" Then
    ' ''                pQFilt.WhereClause = sData & " and MATA_TIP = 2 and IVEDE = '" & sOper & "'"
    ' ''            Else
    ' ''                pQFilt.WhereClause = sData & " and MATA_TIP = 2 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "'"
    ' ''            End If
    ' ''            pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
    ' ''            SklypuStatistika(UBound(SklypuStatistika)).iPrel = pSelSet.Count

    ' ''            pQFiltL = New QueryFilter 'patikra preliminari
    ' ''            If sRajPaieska = "" Then
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 2 and IVEDE = '" & sOper & "' and MATA_TIP = 2"
    ' ''            Else
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 2 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 2"
    ' ''            End If
    ' ''            'pQFiltL.WhereClause = sData & " and DUBL = '2'" 'patikra
    ' ''            pSelSet = pFClassL.Select(pQFiltL, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
    ' ''            ipatikra = pSelSet.Count

    ' ''            pQFiltL = New QueryFilter 'patikra geodezine
    ' ''            If sRajPaieska = "" Then
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 2 and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''            Else
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 2 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''            End If
    ' ''            pSelSet = pFClassL.Select(pQFiltL, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
    ' ''            iPatikraGeo = pSelSet.Count
    ' ''            SklypuStatistika(UBound(SklypuStatistika)).iPatikra = iPatikra + iPatikraGeo


    ' ''            pQFiltL = New QueryFilter 'isvada preliminari registruota
    ' ''            If sRajPaieska = "" Then
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 1 and IVEDE = '" & sOper & "' and MATA_TIP = 2"
    ' ''            Else
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 1 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 2"
    ' ''            End If
    ' ''            pSelSet = pFClassL.Select(pQFiltL, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
    ' ''            iIsvada = pSelSet.Count

    ' ''            pQFiltL = New QueryFilter 'isvada Geodezine registruota
    ' ''            If sRajPaieska = "" Then
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 1 and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''            Else
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 1 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''            End If
    ' ''            pSelSet = pFClassL.Select(pQFiltL, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
    ' ''            iIsvadaGeo = pSelSet.Count
    ' ''            SklypuStatistika(UBound(SklypuStatistika)).iIsvada = iIsvada + iIsvadaGeo


    ' ''            pQFiltL = New QueryFilter 'isvada preliminari patikra
    ' ''            If sRajPaieska = "" Then
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 3 and IVEDE = '" & sOper & "' and MATA_TIP = 2"
    ' ''            Else
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 3 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 2"
    ' ''            End If
    ' ''            pSelSet = pFClassL.Select(pQFiltL, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
    ' ''            iIsvadaPatikra = pSelSet.Count

    ' ''            pQFiltL = New QueryFilter 'isvada Geodezine patikra
    ' ''            If sRajPaieska = "" Then
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 3 and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''            Else
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 3 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''            End If
    ' ''            pSelSet = pFClassL.Select(pQFiltL, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
    ' ''            iIsvadaPatikra = pSelSet.Count
    ' ''            SklypuStatistika(UBound(SklypuStatistika)).iIsvadaPatikra = iIsvadaPatikra + iIsvadaGeoPatikra


    ' ''            pQFiltL = New QueryFilter 'patikra GeoMatininkas
    ' ''            If sRajPaieska = "" Then
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 6 and IVEDE = '" & sOper & "'" ' and MATA_TIP = 1"
    ' ''            Else
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 6 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "'" ' and MATA_TIP = 1"
    ' ''            End If
    ' ''            pSelSet = pFClassL.Select(pQFiltL, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
    ' ''            iPatikraGeoM = pSelSet.Count
    ' ''            SklypuStatistika(UBound(SklypuStatistika)).iPatikraGeoM = iPatikraGeoM

    ' ''            pQFiltL = New QueryFilter 'isvada GeoMatininkas
    ' ''            If sRajPaieska = "" Then
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 7 and IVEDE = '" & sOper & "'"  ' and MATA_TIP = 1"
    ' ''            Else
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 7 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "'"  ' and MATA_TIP = 1"
    ' ''            End If
    ' ''            pSelSet = pFClassL.Select(pQFiltL, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
    ' ''            iIsvadaGeoM = pSelSet.Count
    ' ''            SklypuStatistika(UBound(SklypuStatistika)).iIsvadaGeoM = iIsvadaGeoM

    ' ''            pQFiltL = New QueryFilter 'patikra be NZT
    ' ''            If sRajPaieska = "" Then
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 8 and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''            Else
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 8 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''            End If
    ' ''            pSelSet = pFClassL.Select(pQFiltL, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
    ' ''            iPatikraNZT = pSelSet.Count
    ' ''            SklypuStatistika(UBound(SklypuStatistika)).iPatikraNZT = iPatikraNZT

    ' ''            pQFiltL = New QueryFilter 'isvada be NZT
    ' ''            If sRajPaieska = "" Then
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 9 and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''            Else
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 9 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''            End If
    ' ''            pSelSet = pFClassL.Select(pQFiltL, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
    ' ''            iIsvadaNZT = pSelSet.Count
    ' ''            SklypuStatistika(UBound(SklypuStatistika)).iIsvadaNZT = iIsvadaNZT


    ' ''            pQFiltL = New QueryFilter 'laikinas prel
    ' ''            If sRajPaieska = "" Then
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 4 and IVEDE = '" & sOper & "' and MATA_TIP = 2"
    ' ''            Else
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 4 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 2"
    ' ''            End If
    ' ''            pSelSet = pFClassL.Select(pQFiltL, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
    ' ''            iLaikini = pSelSet.Count
    ' ''            pQFiltL = New QueryFilter 'laikinas geod
    ' ''            If sRajPaieska = "" Then
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 4 and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''            Else
    ' ''                pQFiltL.WhereClause = sData & " and DUBL = 4 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''            End If
    ' ''            pSelSet = pFClassL.Select(pQFiltL, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
    ' ''            iLaikiniGeo = pSelSet.Count
    ' ''            SklypuStatistika(UBound(SklypuStatistika)).iLaikini = iLaikini + iLaikiniGeo
    ' ''            SklypuStatistika(UBound(SklypuStatistika)).iSuma = iGeod + iPrel + iPatikra + iIsvada + iLaikini + iPatikraGeo + iIsvadaGeo + iLaikiniGeo + iIsvadaPatikra + iIsvadaGeoPatikra + iPatikraGeoM + iIsvadaGeoM + iPatikraNZT + iIsvadaNZT
    ' ''            SklypuStatistika(UBound(SklypuStatistika)).iTaskai = iTaskai
    ' ''            ReDim Preserve SklypuStatistika(UBound(SklypuStatistika) + 1)
    ' ''        Else
    ' ''            For i = 0 To iOperatoriuSk - 1
    ' ''                'visa lietuva
    ' ''                sOper = Operatoriai(i)
    ' ''                SklypuStatistika(UBound(SklypuStatistika)).sOperator = sOper
    ' ''                iGeod = 0
    ' ''                iPrel = 0
    ' ''                iPatikra = 0
    ' ''                iIsvada = 0
    ' ''                iIsvadaPatikra = 0
    ' ''                iLaikini = 0
    ' ''                iPatikraGeo = 0
    ' ''                iIsvadaGeo = 0
    ' ''                iIsvadaGeoPatikra = 0
    ' ''                iLaikiniGeo = 0
    ' ''                iPatikraGeoM = 0
    ' ''                iIsvadaGeoM = 0
    ' ''                iPatikraNZT = 0
    ' ''                iIsvadaNZT = 0
    ' ''                iTaskai = 0
    ' ''                pQFilt = New QueryFilter
    ' ''                If sRajPaieska = "" Then 'Geodezinis
    ' ''                    pQFilt.WhereClause = sData & " and MATA_TIP = 1 and IVEDE = '" & sOper & "'"
    ' ''                Else
    ' ''                    pQFilt.WhereClause = sData & " and MATA_TIP = 1 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "'"
    ' ''                End If
    ' ''                pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
    ' ''                iGeod = pSelSet.Count
    ' ''                SklypuStatistika(UBound(SklypuStatistika)).iGeod = iGeod

    ' ''                pQFilt = New QueryFilter 'Preliminarus
    ' ''                If sRajPaieska = "" Then
    ' ''                    pQFilt.WhereClause = sData & " and MATA_TIP = 2 and IVEDE = '" & sOper & "'"
    ' ''                Else
    ' ''                    pQFilt.WhereClause = sData & " and MATA_TIP = 2 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "'"
    ' ''                End If
    ' ''                pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
    ' ''                iPrel = pSelSet.Count
    ' ''                SklypuStatistika(UBound(SklypuStatistika)).iPrel = iPrel

    ' ''                pQFiltL = New QueryFilter 'patikra preliminari 
    ' ''                If sRajPaieska = "" Then
    ' ''                    pQFiltL.WhereClause = sData & " and DUBL = 2 and IVEDE = '" & sOper & "' and MATA_TIP = 2"
    ' ''                Else
    ' ''                    pQFiltL.WhereClause = sData & " and DUBL = 2 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 2"
    ' ''                End If
    ' ''                pSelSet = pFClassR.Select(pQFiltL, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
    ' ''                iPatikra = pSelSet.Count

    ' ''                pQFiltL = New QueryFilter 'patikra geodezine 
    ' ''                If sRajPaieska = "" Then
    ' ''                    pQFiltL.WhereClause = sData & " and DUBL = 2 and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''                Else
    ' ''                    pQFiltL.WhereClause = sData & " and DUBL = 2 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''                End If
    ' ''                pFCursorL1 = pFClassL.Search(pQFiltL, False)
    ' ''                If Not pFCursorL1 Is Nothing Then
    ' ''                    pFeature = pFCursorL1.NextFeature
    ' ''                    While Not pFeature Is Nothing
    ' ''                        iPatikraGeo = iPatikraGeo + 1
    ' ''                        pGeom = pFeature.ShapeCopy
    ' ''                        pPntcol = CType(pGeom, IPointCollection)
    ' ''                        pGeomColl = CType(pGeom, IGeometryCollection)
    ' ''                        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
    ' ''                        If Not pFeature Is Nothing Then
    ' ''                            Marshal.ReleaseComObject(pFeature)
    ' ''                        End If
    ' ''                        pFeature = pFCursorL1.NextFeature
    ' ''                    End While
    ' ''                End If
    ' ''                SklypuStatistika(UBound(SklypuStatistika)).iPatikra = iPatikra + iPatikraGeo


    ' ''                pQFiltL2 = New QueryFilter 'isvada preliminari registruota
    ' ''                If sRajPaieska = "" Then
    ' ''                    pQFiltL2.WhereClause = sData & " and DUBL = 1 and IVEDE = '" & sOper & "' and MATA_TIP = 2"
    ' ''                Else
    ' ''                    pQFiltL2.WhereClause = sData & " and DUBL = 1 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 2"
    ' ''                End If
    ' ''                pFCursorL2 = pFClassL.Search(pQFiltL2, False)
    ' ''                If Not pFCursorL2 Is Nothing Then
    ' ''                    pFeature = pFCursorL2.NextFeature
    ' ''                    While Not pFeature Is Nothing
    ' ''                        iIsvada = iIsvada + 1
    ' ''                        If Not pFeature Is Nothing Then
    ' ''                            Marshal.ReleaseComObject(pFeature)
    ' ''                        End If
    ' ''                        pFeature = pFCursorL2.NextFeature
    ' ''                    End While
    ' ''                End If
    ' ''                pQFiltL2 = New QueryFilter 'isvada Geodezine registruota
    ' ''                If sRajPaieska = "" Then
    ' ''                    pQFiltL2.WhereClause = sData & " and DUBL = 1 and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''                Else
    ' ''                    pQFiltL2.WhereClause = sData & " and DUBL = 1 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''                End If
    ' ''                pFCursorL2 = pFClassL.Search(pQFiltL2, False)
    ' ''                If Not pFCursorL2 Is Nothing Then
    ' ''                    pFeature = pFCursorL2.NextFeature
    ' ''                    While Not pFeature Is Nothing
    ' ''                        iIsvadaGeo = iIsvadaGeo + 1
    ' ''                        pGeom = pFeature.ShapeCopy
    ' ''                        pPntcol = CType(pGeom, IPointCollection)
    ' ''                        pGeomColl = CType(pGeom, IGeometryCollection)
    ' ''                        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
    ' ''                        If Not pFeature Is Nothing Then
    ' ''                            Marshal.ReleaseComObject(pFeature)
    ' ''                        End If
    ' ''                        pFeature = pFCursorL2.NextFeature
    ' ''                    End While
    ' ''                End If
    ' ''                SklypuStatistika(UBound(SklypuStatistika)).iIsvada = iIsvada + iIsvadaGeo


    ' ''                pQFiltL2 = New QueryFilter 'isvada preliminari patikra
    ' ''                If sRajPaieska = "" Then
    ' ''                    pQFiltL2.WhereClause = sData & " and DUBL = 3 and IVEDE = '" & sOper & "' and MATA_TIP = 2"
    ' ''                Else
    ' ''                    pQFiltL2.WhereClause = sData & " and DUBL = 3 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 2"
    ' ''                End If
    ' ''                pFCursorL2 = pFClassL.Search(pQFiltL2, False)
    ' ''                If Not pFCursorL2 Is Nothing Then
    ' ''                    pFeature = pFCursorL2.NextFeature
    ' ''                    While Not pFeature Is Nothing
    ' ''                        iIsvadaPatikra = iIsvadaPatikra + 1
    ' ''                        If Not pFeature Is Nothing Then
    ' ''                            Marshal.ReleaseComObject(pFeature)
    ' ''                        End If
    ' ''                        pFeature = pFCursorL2.NextFeature
    ' ''                    End While
    ' ''                End If
    ' ''                pQFiltL2 = New QueryFilter 'isvada Geodezine patikra
    ' ''                If sRajPaieska = "" Then
    ' ''                    pQFiltL2.WhereClause = sData & " and DUBL = 3 and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''                Else
    ' ''                    pQFiltL2.WhereClause = sData & " and DUBL = 3 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''                End If
    ' ''                pFCursorL2 = pFClassL.Search(pQFiltL2, False)
    ' ''                If Not pFCursorL2 Is Nothing Then
    ' ''                    pFeature = pFCursorL2.NextFeature
    ' ''                    While Not pFeature Is Nothing
    ' ''                        iIsvadaGeoPatikra = iIsvadaGeoPatikra + 1
    ' ''                        pGeom = pFeature.ShapeCopy
    ' ''                        pPntcol = CType(pGeom, IPointCollection)
    ' ''                        pGeomColl = CType(pGeom, IGeometryCollection)
    ' ''                        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
    ' ''                        If Not pFeature Is Nothing Then
    ' ''                            Marshal.ReleaseComObject(pFeature)
    ' ''                        End If
    ' ''                        pFeature = pFCursorL2.NextFeature
    ' ''                    End While
    ' ''                End If
    ' ''                SklypuStatistika(UBound(SklypuStatistika)).iIsvadaPatikra = iIsvadaPatikra + iIsvadaGeoPatikra


    ' ''                pQFiltL3 = New QueryFilter 'patikra GeoMatininkas
    ' ''                If sRajPaieska = "" Then
    ' ''                    pQFiltL3.WhereClause = sData & " and DUBL = 6 and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''                Else
    ' ''                    pQFiltL3.WhereClause = sData & " and DUBL = 6 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''                End If
    ' ''                pFCursorL3 = pFClassL.Search(pQFiltL3, False)
    ' ''                If Not pFCursorL3 Is Nothing Then
    ' ''                    pFeature = pFCursorL3.NextFeature

    ' ''                    While Not pFeature Is Nothing
    ' ''                        iPatikraGeoM = iPatikraGeoM + 1
    ' ''                        pGeom = pFeature.ShapeCopy
    ' ''                        pPntcol = CType(pGeom, IPointCollection)
    ' ''                        pGeomColl = CType(pGeom, IGeometryCollection)
    ' ''                        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
    ' ''                        If Not pFeature Is Nothing Then
    ' ''                            Marshal.ReleaseComObject(pFeature)
    ' ''                        End If
    ' ''                        pFeature = pFCursorL3.NextFeature
    ' ''                    End While
    ' ''                End If
    ' ''                SklypuStatistika(UBound(SklypuStatistika)).iPatikraGeoM = iPatikraGeoM

    ' ''                pQFiltL3 = New QueryFilter 'isvada GeoMatininkas
    ' ''                If sRajPaieska = "" Then
    ' ''                    pQFiltL3.WhereClause = sData & " and DUBL = 7 and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''                Else
    ' ''                    pQFiltL3.WhereClause = sData & " and DUBL = 7 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''                End If
    ' ''                pFCursorL3 = pFClassL.Search(pQFiltL3, False)
    ' ''                If Not pFCursorL3 Is Nothing Then
    ' ''                    pFeature = pFCursorL3.NextFeature

    ' ''                    While Not pFeature Is Nothing
    ' ''                        iIsvadaGeoM = iIsvadaGeoM + 1
    ' ''                        pGeom = pFeature.ShapeCopy
    ' ''                        pPntcol = CType(pGeom, IPointCollection)
    ' ''                        pGeomColl = CType(pGeom, IGeometryCollection)
    ' ''                        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
    ' ''                        If Not pFeature Is Nothing Then
    ' ''                            Marshal.ReleaseComObject(pFeature)
    ' ''                        End If
    ' ''                        pFeature = pFCursorL3.NextFeature
    ' ''                    End While
    ' ''                End If
    ' ''                SklypuStatistika(UBound(SklypuStatistika)).iIsvadaGeoM = iIsvadaGeoM

    ' ''                pQFiltL3 = New QueryFilter 'patikra be NZT
    ' ''                If sRajPaieska = "" Then
    ' ''                    pQFiltL3.WhereClause = sData & " and DUBL = 8 and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''                Else
    ' ''                    pQFiltL3.WhereClause = sData & " and DUBL = 8 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''                End If
    ' ''                pFCursorL3 = pFClassL.Search(pQFiltL3, False)
    ' ''                If Not pFCursorL3 Is Nothing Then
    ' ''                    pFeature = pFCursorL3.NextFeature

    ' ''                    While Not pFeature Is Nothing
    ' ''                        iPatikraNZT = iPatikraNZT + 1
    ' ''                        pGeom = pFeature.ShapeCopy
    ' ''                        pPntcol = CType(pGeom, IPointCollection)
    ' ''                        pGeomColl = CType(pGeom, IGeometryCollection)
    ' ''                        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
    ' ''                        If Not pFeature Is Nothing Then
    ' ''                            Marshal.ReleaseComObject(pFeature)
    ' ''                        End If
    ' ''                        pFeature = pFCursorL3.NextFeature
    ' ''                    End While
    ' ''                End If
    ' ''                SklypuStatistika(UBound(SklypuStatistika)).iPatikraNZT = iPatikraNZT

    ' ''                pQFiltL3 = New QueryFilter 'isvada be NZT
    ' ''                If sRajPaieska = "" Then
    ' ''                    pQFiltL3.WhereClause = sData & " and DUBL = 9 and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''                Else
    ' ''                    pQFiltL3.WhereClause = sData & " and DUBL = 9 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''                End If
    ' ''                pFCursorL3 = pFClassL.Search(pQFiltL3, False)
    ' ''                If Not pFCursorL3 Is Nothing Then
    ' ''                    pFeature = pFCursorL3.NextFeature

    ' ''                    While Not pFeature Is Nothing
    ' ''                        iIsvadaNZT = iIsvadaNZT + 1
    ' ''                        pGeom = pFeature.ShapeCopy
    ' ''                        pPntcol = CType(pGeom, IPointCollection)
    ' ''                        pGeomColl = CType(pGeom, IGeometryCollection)
    ' ''                        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
    ' ''                        If Not pFeature Is Nothing Then
    ' ''                            Marshal.ReleaseComObject(pFeature)
    ' ''                        End If
    ' ''                        pFeature = pFCursorL3.NextFeature
    ' ''                    End While
    ' ''                End If
    ' ''                SklypuStatistika(UBound(SklypuStatistika)).iIsvadaNZT = iIsvadaNZT



    ' ''                pQFiltL3 = New QueryFilter 'laikinas prel
    ' ''                If sRajPaieska = "" Then
    ' ''                    pQFiltL3.WhereClause = sData & " and DUBL = 4 and IVEDE = '" & sOper & "' and MATA_TIP = 2"
    ' ''                Else
    ' ''                    pQFiltL3.WhereClause = sData & " and DUBL = 4 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 2"
    ' ''                End If
    ' ''                'pQFiltL3.WhereClause = sData & " and DUBL = '0'" 
    ' ''                pFCursorL3 = pFClassL.Search(pQFiltL3, False)
    ' ''                If Not pFCursorL3 Is Nothing Then
    ' ''                    pFeature = pFCursorL3.NextFeature

    ' ''                    While Not pFeature Is Nothing
    ' ''                        iLaikini = iLaikini + 1
    ' ''                        If Not pFeature Is Nothing Then
    ' ''                            Marshal.ReleaseComObject(pFeature)
    ' ''                        End If
    ' ''                        pFeature = pFCursorL3.NextFeature
    ' ''                    End While
    ' ''                End If
    ' ''                pQFiltL3 = New QueryFilter 'laikinas geod
    ' ''                If sRajPaieska = "" Then
    ' ''                    pQFiltL3.WhereClause = sData & " and DUBL = 4 and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''                Else
    ' ''                    pQFiltL3.WhereClause = sData & " and DUBL = 4 and (" & sRajPaieska & " ) and IVEDE = '" & sOper & "' and MATA_TIP = 1"
    ' ''                End If
    ' ''                pFCursorL3 = pFClassL.Search(pQFiltL3, False)
    ' ''                If Not pFCursorL3 Is Nothing Then
    ' ''                    pFeature = pFCursorL3.NextFeature

    ' ''                    While Not pFeature Is Nothing
    ' ''                        iLaikiniGeo = iLaikiniGeo + 1
    ' ''                        pGeom = pFeature.ShapeCopy
    ' ''                        pPntcol = CType(pGeom, IPointCollection)
    ' ''                        pGeomColl = CType(pGeom, IGeometryCollection)
    ' ''                        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
    ' ''                        If Not pFeature Is Nothing Then
    ' ''                            Marshal.ReleaseComObject(pFeature)
    ' ''                        End If
    ' ''                        pFeature = pFCursorL3.NextFeature
    ' ''                    End While
    ' ''                End If
    ' ''                SklypuStatistika(UBound(SklypuStatistika)).iLaikini = iLaikini + iLaikiniGeo
    ' ''                SklypuStatistika(UBound(SklypuStatistika)).iSuma = iGeod + iPrel + iPatikra + iIsvada + iLaikini + iPatikraGeo + iIsvadaGeo + iLaikiniGeo + iIsvadaPatikra + iIsvadaGeoPatikra + iPatikraGeoM + iIsvadaGeoM + iPatikraNZT + iIsvadaNZT
    ' ''                SklypuStatistika(UBound(SklypuStatistika)).iTaskai = iTaskai
    ' ''                ReDim Preserve SklypuStatistika(UBound(SklypuStatistika) + 1)
    ' ''            Next i
    ' ''        End If
    ' ''        'Catch ex As Exception
    ' ''        '    MsgBox(ex.Message)
    ' ''        'Finally
    ' ''        'End Try
    ' ''        Exit Function
    ' ''ERN:    MsgBox(Err.Number & "  " & Err.Description)

    ' ''    End Function
    Public Function Paieska_sklypai_raj(ByVal sFilialas As String, ByVal sRajonas As String, ByVal sDataNuo As String, ByVal sDataIki As String, ByVal sOperatorius As String) As Boolean

        Dim pFact As IWorkspaceFactory2 = Nothing
        Dim pMasterPropSet As IPropertySet = Nothing
        Dim pFCursorG As IFeatureCursor
        Dim pFCursorL As IFeatureCursor
        Dim pMasterWorkSpace As IWorkspace = Nothing
        Dim pFWorkSpace As IFeatureWorkspace = Nothing
        Dim pFClassR As IFeatureClass = Nothing
        Dim pFClassL As IFeatureClass = Nothing
        Dim pQFilt As IQueryFilter = Nothing
        Dim pSelSet As ISelectionSet = Nothing
        Dim sData As String
        'Registruoti 
        Dim iGeod As Integer = 0
        Dim iGeodGeoM As Integer = 0 : Dim iGeodkadaGIS As Integer = 0
        Dim iPrel As Integer = 0
        'KadaGIS laikini
        Dim iPatikra_2 As Integer = 0 : Dim iPatikraGeod_2 As Integer = 0
        Dim iIsvada_1 As Integer = 0 : Dim iIsvadaGeod_1 As Integer = 0
        Dim iIsvadaPatikra_3 As Integer = 0 : Dim iIsvadaGeoPatikra_3 As Integer = 0
        Dim iLaikini_4 As Integer = 0 : Dim iLaikiniGeod_4 As Integer = 0
        Dim iPatikraNZT_8 As Integer = 0
        Dim iIsvadaNZT_9 As Integer = 0
        Dim iPatikraKons_12 As Integer = 0
        Dim iIsvadaKons_13 As Integer = 0
        'Geomatininkas laikini
        Dim iLaikiniGeom_5 As Integer = 0
        Dim iPatikraGeoM_6 As Integer = 0
        Dim iIsvadaGeoM_7 As Integer = 0
        Dim iPatikraNZTGeoM_10 As Integer = 0
        Dim iIsvadaNZTGeoM_11 As Integer = 0
        Dim iPatikraKonsGeoM_14 As Integer = 0
        Dim iIsvadaKonsGeoM_15 As Integer = 0
        Dim bGerai As Boolean
        Dim i As Integer
        Dim pFeature As IFeature = Nothing
        Dim sRajPaieska As String = ""
        Dim pGeom As IGeometry = Nothing
        Dim pPntcol As IPointCollection = Nothing
        Dim pGeomColl As IGeometryCollection = Nothing
        Dim sQlGeomat As String = "Dubl = 5 or Dubl = 6 or Dubl = 7 or dubl = 10 or dubl = 11 or dubl = 14 or dubl = 15"
        Dim sQlkadaGIS As String = "Dubl = 0 or Dubl = 4 or Dubl = 1 or dubl = 2 or dubl = 3 or dubl = 8 or dubl = 9 or dubl = 12 or dubl = 13"
        Dim iTaskai As Integer = 0

        Try
            ''pMasterPropSet = New PropertySet
            ''With pMasterPropSet
            ''    .SetProperty("Server", "SDEAPP")
            ''    .SetProperty("Instance", "5151")
            ''    .SetProperty("DATABASE", "")
            ''    .SetProperty("user", "SDE")
            ''    .SetProperty("password", "SDENAUJAS")
            ''    .SetProperty("version", "SDE.DEFAULT")
            ''End With
            ''pFact = New SdeWorkspaceFactory
            ''pMasterWorkSpace = pFact.Open(pMasterPropSet, 0)

            pMasterWorkSpace = ConnectToTransactionalVersionD("sdeapp", "sde:oracle11g:arcsdecdb.world", "adrviewer", "adr2008", "DATABASE", "SDE.DEFAULT", bGerai)  'Test'as

            If Not bGerai Then : Exit Function : End If
            pFWorkSpace = CType(pMasterWorkSpace, IFeatureWorkspace)
            sData = "DATA_R >= date '" & sDataNuo & "' and DATA_R <= date '" & sDataIki & "'"  'isrenkame laikotarpi
            'sklypai
            pFClassR = pFWorkSpace.OpenFeatureClass("KADA.PARCELS")
            pFClassL = pFWorkSpace.OpenFeatureClass("KADA.TEMP")
            If pFClassR Is Nothing Or pFClassR Is Nothing Then Exit Function
            ReDim SklypuStatistika(0)
            If sFilialas = "  " Then  'visa lietuva
                sRajPaieska = ""
            Else  'vienas filialas
                If sRajonas = "  " Then
                    'visi rajonai rajono kodai yra cc()
                    sRajPaieska = "SAV_ID = '" & cc(0).sKodas & "'"
                    For i = 1 To UBound(cc) - 1
                        sRajPaieska = sRajPaieska & " or SAV_ID = '" & cc(i).sKodas & "'"
                    Next i
                Else 'vienas rajonas
                    sRajPaieska = "SAV_ID = '" & sRajonas & "'"
                End If
            End If

            'Suksiu cikla per operatorius
            If sRajonas <> "  " Then
                sRajPaieska = "SAV_ID = '" & sRajonas & "'"
                SklypuStatistika(UBound(SklypuStatistika)).sOperator = sRajonasPavad

                'Regustruoti sklypai PARCELS
                pQFilt = New QueryFilter 'Geodeziniai
                pQFilt.WhereClause = sData & " and MATA_TIP = 1 and (" & sRajPaieska & " )"
                pFCursorG = pFClassR.Search(pQFilt, False)
                If Not pFCursorG Is Nothing Then
                    pFeature = pFCursorG.NextFeature
                    While Not pFeature Is Nothing
                        iGeod = iGeod + 1
                        pGeom = pFeature.ShapeCopy
                        pPntcol = CType(pGeom, IPointCollection)
                        pGeomColl = CType(pGeom, IGeometryCollection)
                        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                        If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                        pFeature = pFCursorG.NextFeature
                    End While
                End If
                SklypuStatistika(UBound(SklypuStatistika)).iGeod = iGeod

                pQFilt = New QueryFilter 'Geodeziniai geomatininkas
                pQFilt.WhereClause = sData & " and MATA_TIP = 1 and (" & sRajPaieska & " ) and ( " & sQlGeomat & " )"
                pFCursorG = pFClassR.Search(pQFilt, False)
                If Not pFCursorG Is Nothing Then
                    pFeature = pFCursorG.NextFeature
                    While Not pFeature Is Nothing
                        iGeodGeoM = iGeodGeoM + 1
                        pGeom = pFeature.ShapeCopy
                        pPntcol = CType(pGeom, IPointCollection)
                        pGeomColl = CType(pGeom, IGeometryCollection)
                        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                        If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                        pFeature = pFCursorG.NextFeature
                    End While
                End If
                SklypuStatistika(UBound(SklypuStatistika)).iGeodGeoM = iGeodGeoM

                pQFilt = New QueryFilter 'Geodeziniai kadagis
                pQFilt.WhereClause = sData & " and MATA_TIP = 1 and (" & sRajPaieska & " ) and ( " & sQlkadaGIS & " )"
                pFCursorG = pFClassR.Search(pQFilt, False)
                If Not pFCursorG Is Nothing Then
                    pFeature = pFCursorG.NextFeature
                    While Not pFeature Is Nothing
                        iGeodkadaGIS = iGeodkadaGIS + 1
                        pGeom = pFeature.ShapeCopy
                        pPntcol = CType(pGeom, IPointCollection)
                        pGeomColl = CType(pGeom, IGeometryCollection)
                        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                        If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                        pFeature = pFCursorG.NextFeature
                    End While
                End If
                SklypuStatistika(UBound(SklypuStatistika)).iGeodKadaGIS = iGeodkadaGIS



                pQFilt = New QueryFilter
                pQFilt.WhereClause = sData & " and MATA_TIP = 2 and (" & sRajPaieska & " )"
                pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iPrel = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iPrel = iPrel

                'Laikini sklypai -TEMP
                'KadaGIS
                ''pQFilt = New QueryFilter 'patikra preliminari
                ''pQFilt.WhereClause = sData & " and DUBL = 2 and (" & sRajPaieska & " ) and MATA_TIP = 2"
                ''pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                ''iPatikra_2 = pSelSet.Count

                ''pQFilt = New QueryFilter 'patikra geodezine
                ''pQFilt.WhereClause = sData & " and DUBL = 2 and (" & sRajPaieska & " ) and MATA_TIP = 1"
                ''pFCursorL = pFClassL.Search(pQFilt, False)
                ''If Not pFCursorL Is Nothing Then
                ''    pFeature = pFCursorL.NextFeature
                ''    While Not pFeature Is Nothing
                ''        iPatikraGeod_2 = iPatikraGeod_2 + 1
                ''        pGeom = pFeature.ShapeCopy
                ''        pPntcol = CType(pGeom, IPointCollection)
                ''        pGeomColl = CType(pGeom, IGeometryCollection)
                ''        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                ''        If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                ''        pFeature = pFCursorL.NextFeature
                ''    End While
                ''End If

                pQFilt = New QueryFilter 'patikra 
                pQFilt.WhereClause = sData & " and DUBL = 2 and (" & sRajPaieska & " )" ' and MATA_TIP = 2"
                pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iPatikra_2 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iPatikra_2 = iPatikra_2 '+ iPatikraGeod_2


                ''pQFilt = New QueryFilter 'isvada preliminari registruota
                ''pQFilt.WhereClause = sData & " and DUBL = 1 and (" & sRajPaieska & " ) and MATA_TIP = 2"
                ''pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                ''iIsvada_1 = pSelSet.Count

                ''pQFilt = New QueryFilter 'isvada Geodezine registruota
                ''pQFilt.WhereClause = sData & " and DUBL = 1 and (" & sRajPaieska & " ) and MATA_TIP = 1"
                ''pFCursorL = pFClassL.Search(pQFilt, False)
                ''If Not pFCursorL Is Nothing Then
                ''    pFeature = pFCursorL.NextFeature
                ''    While Not pFeature Is Nothing
                ''        iIsvadaGeod_1 = iIsvadaGeod_1 + 1
                ''        pGeom = pFeature.ShapeCopy
                ''        pPntcol = CType(pGeom, IPointCollection)
                ''        pGeomColl = CType(pGeom, IGeometryCollection)
                ''        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                ''        If Not pFeature Is Nothing Then
                ''            Marshal.ReleaseComObject(pFeature) : End If : pFeature = pFCursorL.NextFeature
                ''    End While
                ''End If


                pQFilt = New QueryFilter 'isvada preliminari registruota
                pQFilt.WhereClause = sData & " and DUBL = 1 and (" & sRajPaieska & " )" ' and MATA_TIP = 2"
                pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iIsvada_1 = pSelSet.Count

                SklypuStatistika(UBound(SklypuStatistika)).iIsvada_1 = iIsvada_1 '+ iIsvadaGeod_1


                ''pQFilt = New QueryFilter 'isvada preliminari patikra
                ''pQFilt.WhereClause = sData & " and DUBL = 3 and (" & sRajPaieska & " ) and MATA_TIP = 2"
                ''pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                ''iIsvadaPatikra_3 = pSelSet.Count

                ''pQFilt = New QueryFilter 'isvada Geodezine patikra
                ''pQFilt.WhereClause = sData & " and DUBL = 3 and (" & sRajPaieska & " ) and MATA_TIP = 1"
                ''pFCursorL = pFClassL.Search(pQFilt, False)
                ''If Not pFCursorL Is Nothing Then
                ''    pFeature = pFCursorL.NextFeature
                ''    While Not pFeature Is Nothing
                ''        iIsvadaGeoPatikra_3 = iIsvadaGeoPatikra_3 + 1
                ''        pGeom = pFeature.ShapeCopy
                ''        pPntcol = CType(pGeom, IPointCollection)
                ''        pGeomColl = CType(pGeom, IGeometryCollection)
                ''        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                ''        If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                ''        pFeature = pFCursorL.NextFeature
                ''    End While
                ''End If

                pQFilt = New QueryFilter 'isvada preliminari patikra
                pQFilt.WhereClause = sData & " and DUBL = 3 and (" & sRajPaieska & " )" ' and MATA_TIP = 2"
                pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iIsvadaPatikra_3 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iIsvadaPatikra_3 = iIsvadaPatikra_3 '+ iIsvadaGeoPatikra_3


                ''pQFilt = New QueryFilter 'laikinas prel
                ''pQFilt.WhereClause = sData & " and DUBL = 4 and (" & sRajPaieska & " ) and MATA_TIP = 2"
                ''pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                ''iLaikini_4 = pSelSet.Count

                ''pQFilt = New QueryFilter 'laikinas geod
                ''pQFilt.WhereClause = sData & " and DUBL = 4 and (" & sRajPaieska & " ) and MATA_TIP = 1"
                ''pFCursorL = pFClassL.Search(pQFilt, False)
                ''If Not pFCursorL Is Nothing Then
                ''    pFeature = pFCursorL.NextFeature
                ''    While Not pFeature Is Nothing
                ''        iLaikiniGeod_4 = iLaikiniGeod_4 + 1
                ''        pGeom = pFeature.ShapeCopy
                ''        pPntcol = CType(pGeom, IPointCollection)
                ''        pGeomColl = CType(pGeom, IGeometryCollection)
                ''        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                ''        If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                ''        pFeature = pFCursorL.NextFeature
                ''    End While
                ''End If

                pQFilt = New QueryFilter 'laikinas prel
                pQFilt.WhereClause = sData & " and DUBL = 4 and (" & sRajPaieska & " )" ' and MATA_TIP = 2"
                pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iLaikini_4 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iLaikini_4 = iLaikini_4 '+ iLaikiniGeod_4


                pQFilt = New QueryFilter 'patikra be NZT
                pQFilt.WhereClause = sData & " and DUBL = 8 and (" & sRajPaieska & " )" ' and MATA_TIP = 1"
                pFCursorL = pFClassL.Search(pQFilt, False)
                If Not pFCursorL Is Nothing Then
                    pFeature = pFCursorL.NextFeature
                    While Not pFeature Is Nothing
                        iPatikraNZT_8 = iPatikraNZT_8 + 1
                        pGeom = pFeature.ShapeCopy
                        pPntcol = CType(pGeom, IPointCollection)
                        pGeomColl = CType(pGeom, IGeometryCollection)
                        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                        If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                        pFeature = pFCursorL.NextFeature
                    End While
                End If
                SklypuStatistika(UBound(SklypuStatistika)).iPatikraNZT_8 = iPatikraNZT_8


                pQFilt = New QueryFilter 'isvada be NZT
                pQFilt.WhereClause = sData & " and DUBL = 9 and (" & sRajPaieska & " )" ' and MATA_TIP = 1"
                pFCursorL = pFClassL.Search(pQFilt, False)
                If Not pFCursorL Is Nothing Then
                    pFeature = pFCursorL.NextFeature
                    While Not pFeature Is Nothing
                        iIsvadaNZT_9 = iIsvadaNZT_9 + 1
                        pGeom = pFeature.ShapeCopy
                        pPntcol = CType(pGeom, IPointCollection)
                        pGeomColl = CType(pGeom, IGeometryCollection)
                        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                        If Not pFeature Is Nothing Then
                            Marshal.ReleaseComObject(pFeature)
                        End If
                        pFeature = pFCursorL.NextFeature
                    End While
                End If
                SklypuStatistika(UBound(SklypuStatistika)).iIsvadaNZT_9 = iIsvadaNZT_9


                pQFilt = New QueryFilter 'patikra konsolidacija
                pQFilt.WhereClause = sData & " and DUBL = 12 and (" & sRajPaieska & " )" ' and MATA_TIP = 1"
                pFCursorL = pFClassL.Search(pQFilt, False)
                If Not pFCursorL Is Nothing Then
                    pFeature = pFCursorL.NextFeature
                    While Not pFeature Is Nothing
                        iPatikraKons_12 = iPatikraKons_12 + 1
                        pGeom = pFeature.ShapeCopy
                        pPntcol = CType(pGeom, IPointCollection)
                        pGeomColl = CType(pGeom, IGeometryCollection)
                        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                        If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                        pFeature = pFCursorL.NextFeature
                    End While
                End If
                SklypuStatistika(UBound(SklypuStatistika)).iPatikraKons_12 = iPatikraKons_12


                pQFilt = New QueryFilter 'isvada konsolidacija
                pQFilt.WhereClause = sData & " and DUBL = 13 and (" & sRajPaieska & " )" ' and MATA_TIP = 1"
                pFCursorL = pFClassL.Search(pQFilt, False)
                If Not pFCursorL Is Nothing Then
                    pFeature = pFCursorL.NextFeature
                    While Not pFeature Is Nothing
                        iIsvadaKons_13 = iIsvadaKons_13 + 1
                        pGeom = pFeature.ShapeCopy
                        pPntcol = CType(pGeom, IPointCollection)
                        pGeomColl = CType(pGeom, IGeometryCollection)
                        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                        If Not pFeature Is Nothing Then
                            Marshal.ReleaseComObject(pFeature)
                        End If
                        pFeature = pFCursorL.NextFeature
                    End While
                End If
                SklypuStatistika(UBound(SklypuStatistika)).iIsvadaKons_13 = iIsvadaKons_13


                'Geomatininkas
                pQFilt = New QueryFilter 'laikinas GeoMatininkas
                pQFilt.WhereClause = sData & " and DUBL = 5 and (" & sRajPaieska & " )" ' and MATA_TIP = 1"
                pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iLaikiniGeom_5 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iLaikiniGeom_5 = iLaikiniGeom_5


                pQFilt = New QueryFilter 'patikra GeoMatininkas
                pQFilt.WhereClause = sData & " and DUBL = 6 and (" & sRajPaieska & " )" ' and MATA_TIP = 1"
                pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iPatikraGeoM_6 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iPatikraGeoM_6 = iPatikraGeoM_6


                pQFilt = New QueryFilter 'isvada GeoMatininkas
                pQFilt.WhereClause = sData & " and DUBL = 7 and (" & sRajPaieska & " )" ' and MATA_TIP = 1"
                pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iIsvadaGeoM_7 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iIsvadaGeoM_7 = iIsvadaGeoM_7


                pQFilt = New QueryFilter 'patikra GeoMatininkas be NZT
                pQFilt.WhereClause = sData & " and DUBL = 10 and (" & sRajPaieska & " )" ' and MATA_TIP = 1"
                pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iPatikraNZTGeoM_10 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iPatikraNZTGeoM_10 = iPatikraNZTGeoM_10


                pQFilt = New QueryFilter 'isvada GeoMatininkas be NZT
                pQFilt.WhereClause = sData & " and DUBL = 11 and (" & sRajPaieska & " )" ' and MATA_TIP = 1"
                pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iIsvadaNZTGeoM_11 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iIsvadaNZTGeoM_11 = iIsvadaNZTGeoM_11

                pQFilt = New QueryFilter 'patikra GeoMatininkas konsolidacija
                pQFilt.WhereClause = sData & " and DUBL = 14 and (" & sRajPaieska & " )" ' and MATA_TIP = 1"
                pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iPatikraKonsGeoM_14 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iPatikraKonsGeoM_14 = iPatikraKonsGeoM_14


                pQFilt = New QueryFilter 'isvada GeoMatininkas konsolidacija
                pQFilt.WhereClause = sData & " and DUBL = 15 and (" & sRajPaieska & " )" ' and MATA_TIP = 1"
                pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                iIsvadaKonsGeoM_15 = pSelSet.Count
                SklypuStatistika(UBound(SklypuStatistika)).iIsvadaKonsGeoM_15 = iIsvadaKonsGeoM_15

                SklypuStatistika(UBound(SklypuStatistika)).iSuma = iGeod + iPrel + iPatikra_2 + iPatikraGeod_2 + iIsvada_1 + iIsvadaGeod_1 + iIsvadaPatikra_3 + iIsvadaGeoPatikra_3 + _
                iLaikini_4 + iLaikiniGeod_4 + iPatikraNZT_8 + iIsvadaNZT_9 + iPatikraKons_12 + iIsvadaKons_13 + iLaikiniGeom_5 + iPatikraGeoM_6 + iIsvadaGeoM_7 + iPatikraNZTGeoM_10 + _
                iIsvadaNZTGeoM_11 + iPatikraKonsGeoM_14 + iIsvadaKonsGeoM_15
                  SklypuStatistika(UBound(SklypuStatistika)).iTaskai = iTaskai
                ReDim Preserve SklypuStatistika(UBound(SklypuStatistika) + 1)
            Else
                For i = 0 To UBound(cc) - 1
                    iGeod = 0 : iPrel = 0
                    iGeodGeoM = 0 : iGeodkadaGIS = 0
                    iPatikra_2 = 0 : iPatikraGeod_2 = 0
                    iIsvada_1 = 0 : iIsvadaGeod_1 = 0
                    iIsvadaPatikra_3 = 0 : iIsvadaGeoPatikra_3 = 0
                    iLaikini_4 = 0 : iLaikiniGeod_4 = 0
                    iPatikraNZT_8 = 0 : iIsvadaNZT_9 = 0
                    iPatikraKons_12 = 0 : iIsvadaKons_13 = 0
                    iLaikiniGeom_5 = 0 : iPatikraGeoM_6 = 0 : iIsvadaGeoM_7 = 0
                    iPatikraNZTGeoM_10 = 0 : iIsvadaNZTGeoM_11 = 0
                    iPatikraKonsGeoM_14 = 0 : iIsvadaKonsGeoM_15 = 0
                    iTaskai = 0
                    'visa lietuva
                    SklypuStatistika(UBound(SklypuStatistika)).sOperator = cc(i).sRajonas

                    pQFilt = New QueryFilter 'geodezinis
                    pQFilt.WhereClause = sData & " and MATA_TIP = 1 and SAV_ID = '" & cc(i).sKodas & "'"
                    pFCursorG = pFClassR.Search(pQFilt, False) 'Geodeziniai
                    If Not pFCursorG Is Nothing Then
                        pFeature = pFCursorG.NextFeature
                        While Not pFeature Is Nothing
                            iGeod = iGeod + 1
                            pGeom = pFeature.ShapeCopy
                            pPntcol = CType(pGeom, IPointCollection)
                            pGeomColl = CType(pGeom, IGeometryCollection)
                            iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                            If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                            pFeature = pFCursorG.NextFeature
                        End While
                    End If
                    SklypuStatistika(UBound(SklypuStatistika)).iGeod = iGeod

                    pQFilt = New QueryFilter 'geodezinis geomatininkas
                    pQFilt.WhereClause = sData & " and MATA_TIP = 1 and SAV_ID = '" & cc(i).sKodas & "' and ( " & sQlGeomat & " )"
                    pFCursorG = pFClassR.Search(pQFilt, False) 'Geodeziniai
                    If Not pFCursorG Is Nothing Then
                        pFeature = pFCursorG.NextFeature
                        While Not pFeature Is Nothing
                            iGeodGeoM = iGeodGeoM + 1
                            pGeom = pFeature.ShapeCopy
                            pPntcol = CType(pGeom, IPointCollection)
                            pGeomColl = CType(pGeom, IGeometryCollection)
                            iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                            If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                            pFeature = pFCursorG.NextFeature
                        End While
                    End If
                    SklypuStatistika(UBound(SklypuStatistika)).iGeodGeoM = iGeodGeoM


                    pQFilt = New QueryFilter 'geodezinis kadagis
                    pQFilt.WhereClause = sData & " and MATA_TIP = 1 and SAV_ID = '" & cc(i).sKodas & "' and ( " & sQlkadaGIS & " )"
                    pFCursorG = pFClassR.Search(pQFilt, False) 'Geodeziniai
                    If Not pFCursorG Is Nothing Then
                        pFeature = pFCursorG.NextFeature
                        While Not pFeature Is Nothing
                            iGeodkadaGIS = iGeodkadaGIS + 1
                            pGeom = pFeature.ShapeCopy
                            pPntcol = CType(pGeom, IPointCollection)
                            pGeomColl = CType(pGeom, IGeometryCollection)
                            iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                            If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                            pFeature = pFCursorG.NextFeature
                        End While
                    End If
                    SklypuStatistika(UBound(SklypuStatistika)).iGeodKadaGIS = iGeodkadaGIS


                    pQFilt = New QueryFilter  'Preliminarus
                    pQFilt.WhereClause = sData & " and MATA_TIP = 2 and SAV_ID = '" & cc(i).sKodas & "'"
                    pSelSet = pFClassR.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                    iPrel = pSelSet.Count
                    SklypuStatistika(UBound(SklypuStatistika)).iPrel = iPrel
                   
                    'Laikini sklypai -TEMP
                    'KadaGIS
                    ''pQFilt = New QueryFilter 'patikra preliminari
                    ''pQFilt.WhereClause = sData & " and DUBL = 2 and SAV_ID = '" & cc(i).sKodas & "' and MATA_TIP = 2"
                    ''pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                    ''iPatikra_2 = pSelSet.Count

                    ''pQFilt = New QueryFilter 'patikra geodezine
                    ''pQFilt.WhereClause = sData & " and DUBL = 2 and SAV_ID = '" & cc(i).sKodas & "' and MATA_TIP = 1"
                    ''pFCursorL = pFClassL.Search(pQFilt, False)
                    ''If Not pFCursorL Is Nothing Then
                    ''    pFeature = pFCursorL.NextFeature
                    ''    While Not pFeature Is Nothing
                    ''        iPatikraGeod_2 = iPatikraGeod_2 + 1
                    ''        pGeom = pFeature.ShapeCopy
                    ''        pPntcol = CType(pGeom, IPointCollection)
                    ''        pGeomColl = CType(pGeom, IGeometryCollection)
                    ''        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                    ''        If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                    ''        pFeature = pFCursorL.NextFeature
                    ''    End While
                    ''End If

                    pQFilt = New QueryFilter 'patikra 
                    pQFilt.WhereClause = sData & " and DUBL = 2 and SAV_ID = '" & cc(i).sKodas & "'" ' and MATA_TIP = 2"
                    pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                    iPatikra_2 = pSelSet.Count
                    SklypuStatistika(UBound(SklypuStatistika)).iPatikra_2 = iPatikra_2 '+ iPatikraGeod_2


                    ''pQFilt = New QueryFilter 'isvada preliminari registruota
                    ''pQFilt.WhereClause = sData & " and DUBL = 1 and SAV_ID = '" & cc(i).sKodas & "' and MATA_TIP = 2"
                    ''pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                    ''iIsvada_1 = pSelSet.Count

                    ''pQFilt = New QueryFilter 'isvada Geodezine registruota
                    ''pQFilt.WhereClause = sData & " and DUBL = 1 and SAV_ID = '" & cc(i).sKodas & "' and MATA_TIP = 1"
                    ''pFCursorL = pFClassL.Search(pQFilt, False)
                    ''If Not pFCursorL Is Nothing Then
                    ''    pFeature = pFCursorL.NextFeature
                    ''    While Not pFeature Is Nothing
                    ''        iIsvadaGeod_1 = iIsvadaGeod_1 + 1
                    ''        pGeom = pFeature.ShapeCopy
                    ''        pPntcol = CType(pGeom, IPointCollection)
                    ''        pGeomColl = CType(pGeom, IGeometryCollection)
                    ''        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                    ''        If Not pFeature Is Nothing Then
                    ''            Marshal.ReleaseComObject(pFeature) : End If : pFeature = pFCursorL.NextFeature
                    ''    End While
                    ''End If

                    pQFilt = New QueryFilter 'isvada  registruota
                    pQFilt.WhereClause = sData & " and DUBL = 1 and SAV_ID = '" & cc(i).sKodas & "'" ' and MATA_TIP = 2"
                    pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                    iIsvada_1 = pSelSet.Count
                    SklypuStatistika(UBound(SklypuStatistika)).iIsvada_1 = iIsvada_1 '+ iIsvadaGeod_1


                    ''pQFilt = New QueryFilter 'isvada preliminari patikra
                    ''pQFilt.WhereClause = sData & " and DUBL = 3 and SAV_ID = '" & cc(i).sKodas & "' and MATA_TIP = 2"
                    ''pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                    ''iIsvadaPatikra_3 = pSelSet.Count

                    ''pQFilt = New QueryFilter 'isvada Geodezine patikra
                    ''pQFilt.WhereClause = sData & " and DUBL = 3 and SAV_ID = '" & cc(i).sKodas & "' and MATA_TIP = 1"
                    ''pFCursorL = pFClassL.Search(pQFilt, False)
                    ''If Not pFCursorL Is Nothing Then
                    ''    pFeature = pFCursorL.NextFeature
                    ''    While Not pFeature Is Nothing
                    ''        iIsvadaGeoPatikra_3 = iIsvadaGeoPatikra_3 + 1
                    ''        pGeom = pFeature.ShapeCopy
                    ''        pPntcol = CType(pGeom, IPointCollection)
                    ''        pGeomColl = CType(pGeom, IGeometryCollection)
                    ''        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                    ''        If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                    ''        pFeature = pFCursorL.NextFeature
                    ''    End While
                    ''End If

                    pQFilt = New QueryFilter 'isvada  patikra
                    pQFilt.WhereClause = sData & " and DUBL = 3 and SAV_ID = '" & cc(i).sKodas & "'" ' and MATA_TIP = 2"
                    pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                    iIsvadaPatikra_3 = pSelSet.Count
                    SklypuStatistika(UBound(SklypuStatistika)).iIsvadaPatikra_3 = iIsvadaPatikra_3 '+ iIsvadaGeoPatikra_3


                    ''pQFilt = New QueryFilter 'laikinas prel
                    ''pQFilt.WhereClause = sData & " and DUBL = 4 and SAV_ID = '" & cc(i).sKodas & "' and MATA_TIP = 2"
                    ''pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                    ''iLaikini_4 = pSelSet.Count

                    ''pQFilt = New QueryFilter 'laikinas geod
                    ''pQFilt.WhereClause = sData & " and DUBL = 4 and SAV_ID = '" & cc(i).sKodas & "' and MATA_TIP = 1"
                    ''pFCursorL = pFClassL.Search(pQFilt, False)
                    ''If Not pFCursorL Is Nothing Then
                    ''    pFeature = pFCursorL.NextFeature
                    ''    While Not pFeature Is Nothing
                    ''        iLaikiniGeod_4 = iLaikiniGeod_4 + 1
                    ''        pGeom = pFeature.ShapeCopy
                    ''        pPntcol = CType(pGeom, IPointCollection)
                    ''        pGeomColl = CType(pGeom, IGeometryCollection)
                    ''        iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                    ''        If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                    ''        pFeature = pFCursorL.NextFeature
                    ''    End While
                    ''End If

                    pQFilt = New QueryFilter 'laikinas 
                    pQFilt.WhereClause = sData & " and DUBL = 4 and SAV_ID = '" & cc(i).sKodas & "'" ' and MATA_TIP = 2"
                    pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                    iLaikini_4 = pSelSet.Count
                    SklypuStatistika(UBound(SklypuStatistika)).iLaikini_4 = iLaikini_4 '+ iLaikiniGeod_4


                    pQFilt = New QueryFilter 'patikra be NZT
                    pQFilt.WhereClause = sData & " and DUBL = 8 and SAV_ID = '" & cc(i).sKodas & "'" ' and MATA_TIP = 1"
                    pFCursorL = pFClassL.Search(pQFilt, False)
                    If Not pFCursorL Is Nothing Then
                        pFeature = pFCursorL.NextFeature
                        While Not pFeature Is Nothing
                            iPatikraNZT_8 = iPatikraNZT_8 + 1
                            pGeom = pFeature.ShapeCopy
                            pPntcol = CType(pGeom, IPointCollection)
                            pGeomColl = CType(pGeom, IGeometryCollection)
                            iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                            If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                            pFeature = pFCursorL.NextFeature
                        End While
                    End If
                    SklypuStatistika(UBound(SklypuStatistika)).iPatikraNZT_8 = iPatikraNZT_8


                    pQFilt = New QueryFilter 'isvada be NZT
                    pQFilt.WhereClause = sData & " and DUBL = 9 and SAV_ID = '" & cc(i).sKodas & "'" ' and MATA_TIP = 1"
                    pFCursorL = pFClassL.Search(pQFilt, False)
                    If Not pFCursorL Is Nothing Then
                        pFeature = pFCursorL.NextFeature
                        While Not pFeature Is Nothing
                            iIsvadaNZT_9 = iIsvadaNZT_9 + 1
                            pGeom = pFeature.ShapeCopy
                            pPntcol = CType(pGeom, IPointCollection)
                            pGeomColl = CType(pGeom, IGeometryCollection)
                            iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                            If Not pFeature Is Nothing Then
                                Marshal.ReleaseComObject(pFeature)
                            End If
                            pFeature = pFCursorL.NextFeature
                        End While
                    End If
                    SklypuStatistika(UBound(SklypuStatistika)).iIsvadaNZT_9 = iIsvadaNZT_9


                    pQFilt = New QueryFilter 'patikra konsolidacija
                    pQFilt.WhereClause = sData & " and DUBL = 12 and SAV_ID = '" & cc(i).sKodas & "'" ' and MATA_TIP = 1"
                    pFCursorL = pFClassL.Search(pQFilt, False)
                    If Not pFCursorL Is Nothing Then
                        pFeature = pFCursorL.NextFeature
                        While Not pFeature Is Nothing
                            iPatikraKons_12 = iPatikraKons_12 + 1
                            pGeom = pFeature.ShapeCopy
                            pPntcol = CType(pGeom, IPointCollection)
                            pGeomColl = CType(pGeom, IGeometryCollection)
                            iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                            If Not pFeature Is Nothing Then : Marshal.ReleaseComObject(pFeature) : End If
                            pFeature = pFCursorL.NextFeature
                        End While
                    End If
                    SklypuStatistika(UBound(SklypuStatistika)).iPatikraKons_12 = iPatikraKons_12


                    pQFilt = New QueryFilter 'isvada konsolidacija
                    pQFilt.WhereClause = sData & " and DUBL = 13 and SAV_ID = '" & cc(i).sKodas & "'" ' and MATA_TIP = 1"
                    pFCursorL = pFClassL.Search(pQFilt, False)
                    If Not pFCursorL Is Nothing Then
                        pFeature = pFCursorL.NextFeature
                        While Not pFeature Is Nothing
                            iIsvadaKons_13 = iIsvadaKons_13 + 1
                            pGeom = pFeature.ShapeCopy
                            pPntcol = CType(pGeom, IPointCollection)
                            pGeomColl = CType(pGeom, IGeometryCollection)
                            iTaskai = iTaskai + pPntcol.PointCount - pGeomColl.GeometryCount
                            If Not pFeature Is Nothing Then
                                Marshal.ReleaseComObject(pFeature)
                            End If
                            pFeature = pFCursorL.NextFeature
                        End While
                    End If
                    SklypuStatistika(UBound(SklypuStatistika)).iIsvadaKons_13 = iIsvadaKons_13


                    'Geomatininkas
                    pQFilt = New QueryFilter 'laikinas GeoMatininkas
                    pQFilt.WhereClause = sData & " and DUBL = 5 and SAV_ID = '" & cc(i).sKodas & "'" ' and MATA_TIP = 1"
                    pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                    iLaikiniGeom_5 = pSelSet.Count
                    SklypuStatistika(UBound(SklypuStatistika)).iLaikiniGeom_5 = iLaikiniGeom_5


                    pQFilt = New QueryFilter 'patikra GeoMatininkas
                    pQFilt.WhereClause = sData & " and DUBL = 6 and SAV_ID = '" & cc(i).sKodas & "'" ' and MATA_TIP = 1"
                    pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                    iPatikraGeoM_6 = pSelSet.Count
                    SklypuStatistika(UBound(SklypuStatistika)).iPatikraGeoM_6 = iPatikraGeoM_6


                    pQFilt = New QueryFilter 'isvada GeoMatininkas
                    pQFilt.WhereClause = sData & " and DUBL = 7 and SAV_ID = '" & cc(i).sKodas & "'" ' and MATA_TIP = 1"
                    pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                    iIsvadaGeoM_7 = pSelSet.Count
                    SklypuStatistika(UBound(SklypuStatistika)).iIsvadaGeoM_7 = iIsvadaGeoM_7


                    pQFilt = New QueryFilter 'patikra GeoMatininkas be NZT
                    pQFilt.WhereClause = sData & " and DUBL = 10 and SAV_ID = '" & cc(i).sKodas & "'" ' and MATA_TIP = 1"
                    pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                    iPatikraNZTGeoM_10 = pSelSet.Count
                    SklypuStatistika(UBound(SklypuStatistika)).iPatikraNZTGeoM_10 = iPatikraNZTGeoM_10


                    pQFilt = New QueryFilter 'isvada GeoMatininkas be NZT
                    pQFilt.WhereClause = sData & " and DUBL = 11 and SAV_ID = '" & cc(i).sKodas & "'" ' and MATA_TIP = 1"
                    pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                    iIsvadaNZTGeoM_11 = pSelSet.Count
                    SklypuStatistika(UBound(SklypuStatistika)).iIsvadaNZTGeoM_11 = iIsvadaNZTGeoM_11

                    pQFilt = New QueryFilter 'patikra GeoMatininkas konsolidacija
                    pQFilt.WhereClause = sData & " and DUBL = 14 and SAV_ID = '" & cc(i).sKodas & "'" ' and MATA_TIP = 1"
                    pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                    iPatikraKonsGeoM_14 = pSelSet.Count
                    SklypuStatistika(UBound(SklypuStatistika)).iPatikraKonsGeoM_14 = iPatikraKonsGeoM_14


                    pQFilt = New QueryFilter 'isvada GeoMatininkas konsolidacija
                    pQFilt.WhereClause = sData & " and DUBL = 15 and SAV_ID = '" & cc(i).sKodas & "'" ' and MATA_TIP = 1"
                    pSelSet = pFClassL.Select(pQFilt, esriSelectionType.esriSelectionTypeIDSet, esriSelectionOption.esriSelectionOptionNormal, pFWorkSpace)
                    iIsvadaKonsGeoM_15 = pSelSet.Count
                    SklypuStatistika(UBound(SklypuStatistika)).iIsvadaKonsGeoM_15 = iIsvadaKonsGeoM_15

                    SklypuStatistika(UBound(SklypuStatistika)).iSuma = iGeod + iPrel + iPatikra_2 + iPatikraGeod_2 + iIsvada_1 + iIsvadaGeod_1 + iIsvadaPatikra_3 + iIsvadaGeoPatikra_3 + _
                    iLaikini_4 + iLaikiniGeod_4 + iPatikraNZT_8 + iIsvadaNZT_9 + iPatikraKons_12 + iIsvadaKons_13 + iLaikiniGeom_5 + iPatikraGeoM_6 + iIsvadaGeoM_7 + iPatikraNZTGeoM_10 + _
                    iIsvadaNZTGeoM_11 + iPatikraKonsGeoM_14 + iIsvadaKonsGeoM_15
                    SklypuStatistika(UBound(SklypuStatistika)).iTaskai = iTaskai
                    ReDim Preserve SklypuStatistika(UBound(SklypuStatistika) + 1)
                Next i
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
        End Try
    End Function
    Public Function Paieska_inzineriniai_oper(ByVal sFilialas As String, ByVal sRajonas As String, ByVal sDataNuo As String, ByVal sDataIki As String, ByVal sOperatorius As String) As Boolean

        Dim pFact As IWorkspaceFactory2
        Dim pMasterPropSet As IPropertySet2
        Dim pMasterWorkSpace As IWorkspace
        Dim pFWorkSpace As IFeatureWorkspace
        Dim pFCursor1 As IFeatureCursor
        Dim pFCursor2 As IFeatureCursor
        Dim pFCursorL1 As IFeatureCursor
        Dim pFCursorL2 As IFeatureCursor
        Dim pFClass1 As IFeatureClass
        Dim pFClass2 As IFeatureClass
        Dim pFClass1L As IFeatureClass
        Dim pFClass2L As IFeatureClass
        Dim pQFilt1 As IQueryFilter
        Dim pQFilt2 As IQueryFilter
        Dim pQFilt1L As IQueryFilter
        Dim pQFilt2L As IQueryFilter
        Dim sData As String
        Dim iInz As Integer = 0
        Dim iInz2 As Integer = 0
        Dim iInzL As Integer = 0
        Dim iInz2L As Integer = 0
        Dim iInzGeo As Integer = 0
        Dim iInz2Geo As Integer = 0
        Dim iInzLGeo As Integer = 0
        Dim iInz2LGeo As Integer = 0
        Dim i As Integer
        Dim pFeature As IFeature
        Dim sRajPaieska As String = ""
        Dim bGerai As Boolean
        Try
            ''pMasterPropSet = New PropertySet
            ''With pMasterPropSet
            ''    .SetProperty("Server", "SDEAPP")
            ''    .SetProperty("Instance", "5151")
            ''    .SetProperty("DATABASE", "")
            ''    .SetProperty("user", "SDE")
            ''    .SetProperty("password", "SDENAUJAS")
            ''    .SetProperty("version", "SDE.DEFAULT")
            ''End With
            ''pFact = New SdeWorkspaceFactory
            ''pMasterWorkSpace = pFact.Open(pMasterPropSet, 0)

            pMasterWorkSpace = ConnectToTransactionalVersionD("sdeapp", "sde:oracle11g:arcsdecdb.world", "adrviewer", "adr2008", "DATABASE", "SDE.DEFAULT", bGerai)  'Test'as

            If Not bGerai Then : Exit Function : End If

            pFWorkSpace = CType(pMasterWorkSpace, IFeatureWorkspace)
            'isrenkame laikotarpi
            sData = "DATA_R >= date '" & sDataNuo & "' and DATA_R <= date '" & sDataIki & "'"
            'inzineriniai
            pFClass1 = pFWorkSpace.OpenFeatureClass("KADA.INZINER")
            pFClass2 = pFWorkSpace.OpenFeatureClass("KADA.INZINER2")
            pFClass1L = pFWorkSpace.OpenFeatureClass("KADA.INZINER_LAIK")
            pFClass2L = pFWorkSpace.OpenFeatureClass("KADA.INZINER2_LAIK")
            If pFClass1 Is Nothing Or pFClass2 Is Nothing Or pFClass1L Is Nothing Or pFClass2L Is Nothing Then Exit Function

                sRajPaieska = "FILIALAS = '" & sFilialas & "'"
            
            ReDim InzStatistika(0)
            Dim sOper As String
            Dim pGeom As IGeometry
            Dim pPntcol As IPointCollection
            Dim pGeomColl As IGeometryCollection
            'Suksiu cikla per operatorius
            If Not (Trim(sOperatorius) = "" Or Trim(sOperatorius) = "Visi operatoriai") Then
                sOper = sOperatorius
                InzStatistika(UBound(InzStatistika)).sOperator = sOper
                'visa lietuva
                pQFilt1 = New QueryFilter
                pQFilt1.WhereClause = sData & " and IVEDE = '" & sOper & "' and MATA_TIP = 1"
                pFCursor1 = pFClass1.Search(pQFilt1, False) 'Geodeziniai
                If Not pFCursor1 Is Nothing Then
                    pFeature = pFCursor1.NextFeature
                    While Not pFeature Is Nothing
                        pGeom = pFeature.ShapeCopy
                        pPntcol = CType(pGeom, IPointCollection)
                        iInzGeo = iInzGeo + pPntcol.PointCount
                        pFeature = pFCursor1.NextFeature
                    End While
                End If
                pQFilt1 = New QueryFilter
                pQFilt1.WhereClause = sData & " and IVEDE = '" & sOper & "'"
                pFCursor1 = pFClass1.Search(pQFilt1, False) 'Geodeziniai
                If Not pFCursor1 Is Nothing Then
                    pFeature = pFCursor1.NextFeature
                    While Not pFeature Is Nothing
                        iInz = iInz + 1
                        If Not pFeature Is Nothing Then
                            Marshal.ReleaseComObject(pFeature)
                        End If
                        pFeature = pFCursor1.NextFeature
                    End While
                End If
                InzStatistika(UBound(InzStatistika)).iInz1 = iInz
                'MsgBox(iInz)


                pQFilt2 = New QueryFilter
                pQFilt2.WhereClause = sData & " and IVEDE = '" & sOper & "' and MATA_TIP = 1"
                pFCursor2 = pFClass2.Search(pQFilt2, False) 'Geodeziniai
                If Not pFCursor2 Is Nothing Then
                    pFeature = pFCursor2.NextFeature
                    While Not pFeature Is Nothing
                        pGeom = pFeature.ShapeCopy
                        pPntcol = CType(pGeom, IPointCollection)
                        pGeomColl = CType(pGeom, IGeometryCollection)
                        iInz2Geo = iInz2Geo + pPntcol.PointCount - pGeomColl.GeometryCount
                        pFeature = pFCursor2.NextFeature
                    End While
                End If
                pQFilt2 = New QueryFilter
                pQFilt2.WhereClause = sData & " and IVEDE = '" & sOper & "'"
                pFCursor2 = pFClass2.Search(pQFilt2, False)  'Preliminarus
                If Not pFCursor2 Is Nothing Then
                    pFeature = pFCursor2.NextFeature
                    While Not pFeature Is Nothing
                        iInz2 = iInz2 + 1
                        If Not pFeature Is Nothing Then
                            Marshal.ReleaseComObject(pFeature)
                        End If
                        pFeature = pFCursor2.NextFeature
                    End While
                End If
                InzStatistika(UBound(InzStatistika)).iInz2 = iInz2
                'MsgBox(iInz2)


                pQFilt1L = New QueryFilter
                pQFilt1L.WhereClause = sData & " and IVEDE = '" & sOper & "' and MATA_TIP = 1"
                pFCursorL1 = pFClass1L.Search(pQFilt1L, False) 'Geodeziniai
                If Not pFCursorL1 Is Nothing Then
                    pFeature = pFCursorL1.NextFeature
                    While Not pFeature Is Nothing
                        pGeom = pFeature.ShapeCopy
                        pPntcol = CType(pGeom, IPointCollection)
                        iInzLGeo = iInzLGeo + pPntcol.PointCount
                        pFeature = pFCursorL1.NextFeature
                    End While
                End If
                pQFilt1L = New QueryFilter
                pQFilt1L.WhereClause = sData & " and IVEDE = '" & sOper & "'"
                pFCursorL1 = pFClass1L.Search(pQFilt1L, False)
                If Not pFCursorL1 Is Nothing Then
                    pFeature = pFCursorL1.NextFeature
                    While Not pFeature Is Nothing
                        iInzL = iInzL + 1
                        If Not pFeature Is Nothing Then
                            Marshal.ReleaseComObject(pFeature)
                        End If
                        pFeature = pFCursorL1.NextFeature
                    End While
                End If
                InzStatistika(UBound(InzStatistika)).iInz1Laikini = iInzL
                'MsgBox(iInzL)


                pQFilt2L = New QueryFilter
                pQFilt2L.WhereClause = sData & " and IVEDE = '" & sOper & "' and MATA_TIP = 1"
                pFCursorL2 = pFClass2L.Search(pQFilt2L, False) 'Geodeziniai
                If Not pFCursorL2 Is Nothing Then
                    pFeature = pFCursorL2.NextFeature
                    While Not pFeature Is Nothing
                        pGeom = pFeature.ShapeCopy
                        pPntcol = CType(pGeom, IPointCollection)
                        pGeomColl = CType(pGeom, IGeometryCollection)
                        iInz2LGeo = iInz2LGeo + pPntcol.PointCount - pGeomColl.GeometryCount
                        pFeature = pFCursorL2.NextFeature
                    End While
                End If
                pQFilt2L = New QueryFilter
                pQFilt2L.WhereClause = sData & " and IVEDE = '" & sOper & "'"
                pFCursorL2 = pFClass2L.Search(pQFilt2L, False)
                If Not pFCursorL2 Is Nothing Then
                    pFeature = pFCursorL2.NextFeature
                    While Not pFeature Is Nothing
                        iInz2L = iInz2L + 1
                        If Not pFeature Is Nothing Then
                            Marshal.ReleaseComObject(pFeature)
                        End If
                        pFeature = pFCursorL2.NextFeature
                    End While
                End If
                InzStatistika(UBound(InzStatistika)).iInz2Laikini = iInz2L
                InzStatistika(UBound(InzStatistika)).iSuma = iInz + iInz2 + iInzL + iInz2L
                InzStatistika(UBound(InzStatistika)).iTaskai = iInzGeo + iInz2Geo + iInzLGeo + iInz2LGeo
                'MsgBox(iInz2L)iInzLGeo
                ReDim Preserve InzStatistika(UBound(InzStatistika) + 1)
            Else
                For i = 0 To iOperatoriuSk - 1
                    'visa lietuva
                    sOper = Operatoriai(i)
                    InzStatistika(UBound(InzStatistika)).sOperator = sOper
                    iInz = 0
                    iInz2 = 0
                    iInzL = 0
                    iInz2L = 0
                    iInzGeo = 0
                    iInz2Geo = 0
                    iInzLGeo = 0
                    iInz2LGeo = 0
                    'visa lietuva

                    pQFilt1 = New QueryFilter
                    pQFilt1.WhereClause = sData & " and IVEDE = '" & sOper & "' and MATA_TIP = 1"
                    pFCursor1 = pFClass1.Search(pQFilt1, False) 'Geodeziniai
                    If Not pFCursor1 Is Nothing Then
                        pFeature = pFCursor1.NextFeature
                        While Not pFeature Is Nothing
                            pGeom = pFeature.ShapeCopy
                            pPntcol = CType(pGeom, IPointCollection)
                            iInzGeo = iInzGeo + pPntcol.PointCount
                            pFeature = pFCursor1.NextFeature
                        End While
                    End If
                    pQFilt1 = New QueryFilter
                    pQFilt1.WhereClause = sData & " and IVEDE = '" & sOper & "'"
                    pFCursor1 = pFClass1.Search(pQFilt1, False) 'Geodeziniai
                    If Not pFCursor1 Is Nothing Then
                        pFeature = pFCursor1.NextFeature
                        While Not pFeature Is Nothing
                            iInz = iInz + 1
                            If Not pFeature Is Nothing Then
                                Marshal.ReleaseComObject(pFeature)
                            End If
                            pFeature = pFCursor1.NextFeature
                        End While
                    End If
                    InzStatistika(UBound(InzStatistika)).iInz1 = iInz
                    'MsgBox(iInz)


                    pQFilt2 = New QueryFilter
                    pQFilt2.WhereClause = sData & " and IVEDE = '" & sOper & "' and MATA_TIP = 1"
                    pFCursor2 = pFClass2.Search(pQFilt2, False) 'Geodeziniai
                    If Not pFCursor2 Is Nothing Then
                        pFeature = pFCursor2.NextFeature
                        While Not pFeature Is Nothing
                            pGeom = pFeature.ShapeCopy
                            pPntcol = CType(pGeom, IPointCollection)
                            pGeomColl = CType(pGeom, IGeometryCollection)
                            iInz2Geo = iInz2Geo + pPntcol.PointCount - pGeomColl.GeometryCount
                            pFeature = pFCursor2.NextFeature
                        End While
                    End If
                    pQFilt2 = New QueryFilter
                    pQFilt2.WhereClause = sData & " and IVEDE = '" & sOper & "'"
                    pFCursor2 = pFClass2.Search(pQFilt2, False)  'Preliminarus
                    If Not pFCursor2 Is Nothing Then
                        pFeature = pFCursor2.NextFeature
                        While Not pFeature Is Nothing
                            iInz2 = iInz2 + 1
                            If Not pFeature Is Nothing Then
                                Marshal.ReleaseComObject(pFeature)
                            End If
                            pFeature = pFCursor2.NextFeature
                        End While
                    End If
                    InzStatistika(UBound(InzStatistika)).iInz2 = iInz2
                    'MsgBox(iInz2)

                    pQFilt1L = New QueryFilter
                    pQFilt1L.WhereClause = sData & " and IVEDE = '" & sOper & "' and MATA_TIP = 1"
                    pFCursorL1 = pFClass1L.Search(pQFilt1L, False) 'Geodeziniai
                    If Not pFCursorL1 Is Nothing Then
                        pFeature = pFCursorL1.NextFeature
                        While Not pFeature Is Nothing
                            pGeom = pFeature.ShapeCopy
                            pPntcol = CType(pGeom, IPointCollection)
                            iInzLGeo = iInzLGeo + pPntcol.PointCount
                            pFeature = pFCursorL1.NextFeature
                        End While
                    End If
                    pQFilt1L = New QueryFilter
                    pQFilt1L.WhereClause = sData & " and IVEDE = '" & sOper & "'"
                    pFCursorL1 = pFClass1L.Search(pQFilt1L, False)
                    If Not pFCursorL1 Is Nothing Then
                        pFeature = pFCursorL1.NextFeature
                        While Not pFeature Is Nothing
                            iInzL = iInzL + 1
                            If Not pFeature Is Nothing Then
                                Marshal.ReleaseComObject(pFeature)
                            End If
                            pFeature = pFCursorL1.NextFeature
                        End While
                    End If
                    InzStatistika(UBound(InzStatistika)).iInz1Laikini = iInzL
                    'MsgBox(iInzL)

                    pQFilt2L = New QueryFilter
                    pQFilt2L.WhereClause = sData & " and IVEDE = '" & sOper & "' and MATA_TIP = 1"
                    pFCursorL2 = pFClass2L.Search(pQFilt2L, False) 'Geodeziniai
                    If Not pFCursorL2 Is Nothing Then
                        pFeature = pFCursorL2.NextFeature
                        While Not pFeature Is Nothing
                            pGeom = pFeature.ShapeCopy
                            pPntcol = CType(pGeom, IPointCollection)
                            pGeomColl = CType(pGeom, IGeometryCollection)
                            iInz2LGeo = iInz2LGeo + pPntcol.PointCount - pGeomColl.GeometryCount
                            pFeature = pFCursorL2.NextFeature
                        End While
                    End If
                    pQFilt2L = New QueryFilter
                    pQFilt2L.WhereClause = sData & " and IVEDE = '" & sOper & "'"
                    ' pQFiltL2.WhereClause = sData & " and DUBL = '1'" 'isvada
                    pFCursorL2 = pFClass2L.Search(pQFilt2L, False)
                    If Not pFCursorL2 Is Nothing Then
                        pFeature = pFCursorL2.NextFeature
                        While Not pFeature Is Nothing
                            iInz2L = iInz2L + 1
                            If Not pFeature Is Nothing Then
                                Marshal.ReleaseComObject(pFeature)
                            End If
                            pFeature = pFCursorL2.NextFeature
                        End While
                    End If
                    InzStatistika(UBound(InzStatistika)).iInz2Laikini = iInz2L
                    InzStatistika(UBound(InzStatistika)).iSuma = iInz + iInz2 + iInzL + iInz2L
                    InzStatistika(UBound(InzStatistika)).iTaskai = iInzGeo + iInz2Geo + iInzLGeo + iInz2LGeo
                    ReDim Preserve InzStatistika(UBound(InzStatistika) + 1)
                Next i
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
        End Try
    End Function
    Public Function Paieska_inzineriniai_raj(ByVal sFilialas As String, ByVal sRajonas As String, ByVal sDataNuo As String, ByVal sDataIki As String, ByVal sOperatorius As String) As Boolean

        Dim pFact As IWorkspaceFactory2
        Dim pMasterPropSet As IPropertySet2
        Dim pMasterWorkSpace As IWorkspace
        Dim pFWorkSpace As IFeatureWorkspace
        Dim pFCursor1 As IFeatureCursor
        Dim pFCursor2 As IFeatureCursor
        Dim pFCursorL1 As IFeatureCursor
        Dim pFCursorL2 As IFeatureCursor
        Dim pFClass1 As IFeatureClass
        Dim pFClass2 As IFeatureClass
        Dim pFClass1L As IFeatureClass
        Dim pFClass2L As IFeatureClass
        Dim pQFilt1 As IQueryFilter
        Dim pQFilt2 As IQueryFilter
        Dim pQFilt1L As IQueryFilter
        Dim pQFilt2L As IQueryFilter
        Dim sData As String
        Dim iInz As Integer = 0
        Dim iInz2 As Integer = 0
        Dim iInzL As Integer = 0
        Dim iInz2L As Integer = 0
        Dim iInzGeo As Integer = 0
        Dim iInz2Geo As Integer = 0
        Dim iInzLGeo As Integer = 0
        Dim iInz2LGeo As Integer = 0
        Dim pFeature As IFeature
        Dim sRajPaieska As String = ""
        Dim bGerai As Boolean
        Try
            ''pMasterPropSet = New PropertySet
            ''With pMasterPropSet
            ''    .SetProperty("Server", "SDEAPP")
            ''    .SetProperty("Instance", "5151")
            ''    .SetProperty("DATABASE", "")
            ''    .SetProperty("user", "SDE")
            ''    .SetProperty("password", "SDENAUJAS")
            ''    .SetProperty("version", "SDE.DEFAULT")
            ''End With
            ''pFact = New SdeWorkspaceFactory
            ''pMasterWorkSpace = pFact.Open(pMasterPropSet, 0)

            pMasterWorkSpace = ConnectToTransactionalVersionD("sdeapp", "sde:oracle11g:arcsdecdb.world", "adrviewer", "adr2008", "DATABASE", "SDE.DEFAULT", bGerai)  'Test'as
            If Not bGerai Then : Exit Function : End If

            pFWorkSpace = CType(pMasterWorkSpace, IFeatureWorkspace)
            'isrenkame laikotarpi
            sData = "DATA_R >= date '" & sDataNuo & "' and DATA_R <= date '" & sDataIki & "'"
            'inzineriniai
            pFClass1 = pFWorkSpace.OpenFeatureClass("KADA.INZINER")
            pFClass2 = pFWorkSpace.OpenFeatureClass("KADA.INZINER2")
            pFClass1L = pFWorkSpace.OpenFeatureClass("KADA.INZINER_LAIK")
            pFClass2L = pFWorkSpace.OpenFeatureClass("KADA.INZINER2_LAIK")
            If pFClass1 Is Nothing Or pFClass2 Is Nothing Or pFClass1L Is Nothing Or pFClass2L Is Nothing Then Exit Function
            If sFilialas = "  " Then
                'visa lietuva
                sRajPaieska = ""
            Else
                'vienas filialas
                'If sRajonas = "  " Then
                '    'visi rajonai rajono kodai yra cc()
                '    sRajPaieska = "SAV_ID = '" & cc(0).sKodas & "'"
                '    For i = 1 To UBound(cc) - 1
                '        sRajPaieska = sRajPaieska & " or SAV_ID = '" & cc(i).sKodas & "'"
                '    Next i
                'Else
                '    'vienas rajonas
                '    sRajPaieska = "SAV_ID = '" & sRajonas & "'"
                'End If
                sRajPaieska = "FILIALAS = '" & sFilialas & "'"
            End If

            ReDim InzStatistika(0)
            'Suksiu cikla per operatorius
            'If Not (Trim(sOperatorius) = "" Or Trim(sOperatorius) = "Visi operatoriai") Then

            InzStatistika(UBound(InzStatistika)).sOperator = sRajonasPavad
            'visa lietuva
            pQFilt1 = New QueryFilter
            If sRajPaieska = "" Then
                pQFilt1.WhereClause = sData
            Else
                pQFilt1.WhereClause = sData & " and (" & sRajPaieska & " )"
            End If
            pFCursor1 = pFClass1.Search(pQFilt1, False) 'Geodeziniai
            If Not pFCursor1 Is Nothing Then
                pFeature = pFCursor1.NextFeature
                While Not pFeature Is Nothing
                    iInz = iInz + 1
                    If Not pFeature Is Nothing Then
                        Marshal.ReleaseComObject(pFeature)
                    End If
                    pFeature = pFCursor1.NextFeature
                End While
            End If
            InzStatistika(UBound(InzStatistika)).iInz1 = iInz
            'MsgBox(iInz)

            pQFilt2 = New QueryFilter
            If sRajPaieska = "" Then
                pQFilt2.WhereClause = sData
            Else
                pQFilt2.WhereClause = sData & " and (" & sRajPaieska & " )"
            End If
            pFCursor2 = pFClass2.Search(pQFilt2, False)  'Preliminarus
            If Not pFCursor2 Is Nothing Then
                pFeature = pFCursor2.NextFeature
                While Not pFeature Is Nothing
                    iInz2 = iInz2 + 1
                    If Not pFeature Is Nothing Then
                        Marshal.ReleaseComObject(pFeature)
                    End If
                    pFeature = pFCursor2.NextFeature
                End While
            End If
            InzStatistika(UBound(InzStatistika)).iInz2 = iInz2
            'MsgBox(iInz2)

            pQFilt1L = New QueryFilter
            If sRajPaieska = "" Then
                pQFilt1L.WhereClause = sData
            Else
                pQFilt1L.WhereClause = sData & " and (" & sRajPaieska & " )"
            End If
            'pQFiltL.WhereClause = sData & " and DUBL = '2'" 'patikra
            pFCursorL1 = pFClass1L.Search(pQFilt1L, False)
            If Not pFCursorL1 Is Nothing Then
                pFeature = pFCursorL1.NextFeature
                While Not pFeature Is Nothing
                    iInzL = iInzL + 1
                    If Not pFeature Is Nothing Then
                        Marshal.ReleaseComObject(pFeature)
                    End If
                    pFeature = pFCursorL1.NextFeature
                End While
            End If
            InzStatistika(UBound(InzStatistika)).iInz1Laikini = iInzL
            'MsgBox(iInzL)

            pQFilt2L = New QueryFilter
            If sRajPaieska = "" Then
                pQFilt2L.WhereClause = sData
            Else
                pQFilt2L.WhereClause = sData & " and (" & sRajPaieska & " )"
            End If
            ' pQFiltL2.WhereClause = sData & " and DUBL = '1'" 'isvada
            pFCursorL2 = pFClass2L.Search(pQFilt2L, False)
            If Not pFCursorL2 Is Nothing Then
                pFeature = pFCursorL2.NextFeature
                While Not pFeature Is Nothing
                    iInz2L = iInz2L + 1
                    If Not pFeature Is Nothing Then
                        Marshal.ReleaseComObject(pFeature)
                    End If
                    pFeature = pFCursorL2.NextFeature
                End While
            End If
            InzStatistika(UBound(InzStatistika)).iInz2Laikini = iInzL
            'MsgBox(iInz2L)
            ReDim Preserve InzStatistika(UBound(InzStatistika) + 1)
            ' Else
            'For i = 0 To UBound(cc) - 1
            '    'visa lietuva
            '    InzStatistika(UBound(InzStatistika)).sOperator = cc(i).sRajonas
            '    'visa lietuva
            '    pQFilt1 = New QueryFilter
            '    If sRajPaieska = "" Then
            '        pQFilt1.WhereClause = sData
            '    Else
            '        pQFilt1.WhereClause = sData & " and (" & cc(i).sKodas & " )"
            '    End If
            '    pFCursor1 = pFClass1.Search(pQFilt1, False) 'Geodeziniai
            '    If Not pFCursor1 Is Nothing Then
            '        pFeature = pFCursor1.NextFeature
            '        While Not pFeature Is Nothing
            '            iInz = iInz + 1
            '            If Not pFeature Is Nothing Then
            '                Marshal.ReleaseComObject(pFeature)
            '            End If
            '            pFeature = pFCursor1.NextFeature
            '        End While
            '    End If
            '    InzStatistika(UBound(InzStatistika)).iInz1 = iInz
            '    'MsgBox(iInz)

            '    pQFilt2 = New QueryFilter
            '    If sRajPaieska = "" Then
            '        pQFilt2.WhereClause = sData
            '    Else
            '        pQFilt2.WhereClause = sData & " and (" & cc(i).sKodas & " )"
            '    End If
            '    pFCursor2 = pFClass2.Search(pQFilt2, False)  'Preliminarus
            '    If Not pFCursor2 Is Nothing Then
            '        pFeature = pFCursor2.NextFeature
            '        While Not pFeature Is Nothing
            '            iInz2 = iInz2 + 1
            '            If Not pFeature Is Nothing Then
            '                Marshal.ReleaseComObject(pFeature)
            '            End If
            '            pFeature = pFCursor2.NextFeature
            '        End While
            '    End If
            '    InzStatistika(UBound(InzStatistika)).iInz2 = iInz2
            '    'MsgBox(iInz2)

            '    pQFilt1L = New QueryFilter
            '    If sRajPaieska = "" Then
            '        pQFilt1L.WhereClause = sData
            '    Else
            '        pQFilt1L.WhereClause = sData & " and (" & cc(i).sKodas & " )"
            '    End If
            '    'pQFiltL.WhereClause = sData & " and DUBL = '2'" 'patikra
            '    pFCursorL1 = pFClass1L.Search(pQFilt1L, False)
            '    If Not pFCursorL1 Is Nothing Then
            '        pFeature = pFCursorL1.NextFeature
            '        While Not pFeature Is Nothing
            '            iInzL = iInzL + 1
            '            If Not pFeature Is Nothing Then
            '                Marshal.ReleaseComObject(pFeature)
            '            End If
            '            pFeature = pFCursorL1.NextFeature
            '        End While
            '    End If
            '    InzStatistika(UBound(InzStatistika)).iInz1Laikini = iInzL
            '    'MsgBox(iInzL)

            '    pQFilt2L = New QueryFilter
            '    If sRajPaieska = "" Then
            '        pQFilt2L.WhereClause = sData
            '    Else
            '        pQFilt2L.WhereClause = sData & " and (" & cc(i).sKodas & " )"
            '    End If
            '    ' pQFiltL2.WhereClause = sData & " and DUBL = '1'" 'isvada
            '    pFCursorL2 = pFClass2L.Search(pQFilt2L, False)
            '    If Not pFCursorL2 Is Nothing Then
            '        pFeature = pFCursorL2.NextFeature
            '        While Not pFeature Is Nothing
            '            iInz2L = iInz2L + 1
            '            If Not pFeature Is Nothing Then
            '                Marshal.ReleaseComObject(pFeature)
            '            End If
            '            pFeature = pFCursorL2.NextFeature
            '        End While
            '    End If
            '    InzStatistika(UBound(InzStatistika)).iInz2Laikini = iInzL
            '    'MsgBox(iInz2L)
            '    ReDim Preserve InzStatistika(UBound(InzStatistika) + 1)
            'Next i
            'End If

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
        End Try
    End Function

    Public Function createDBF(ByVal strName As String, ByVal strFolder As String, ByVal sINZ_SKL As String, ByVal SName As String, ByVal bTempToParcel As Boolean) As ITable      ' createDBF: simple function to create a DBASE file.
        ' Open the Workspace
        Dim pFWS As IFeatureWorkspace = Nothing
        Dim pWorkspaceFactory As IWorkspaceFactory = Nothing
        Dim pFieldsEdit As IFieldsEdit = Nothing
        Dim pFieldEdit As IFieldEdit2 = Nothing
        Dim pField As IField = Nothing
        Dim pFields As IFields = Nothing
        '  Try
        On Error GoTo ERN
        pWorkspaceFactory = New ShapefileWorkspaceFactory
        pFWS = CType(pWorkspaceFactory.OpenFromFile(strFolder, 0), IFeatureWorkspace)
        Select Case sINZ_SKL
            Case "S"

                pFields = New Fields
                pFieldsEdit = CType(pFields, IFieldsEdit)
                pFieldsEdit.FieldCount_2 = 22

                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Name_2 = "OID"  'OID
                    .Type_2 = esriFieldType.esriFieldTypeOID
                End With
                pFieldsEdit.Field_2(0) = pField

                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Length_2 = 30
                    .Name_2 = SName 'operatorius
                    .Type_2 = esriFieldType.esriFieldTypeString
                End With
                pFieldsEdit.Field_2(1) = pField

                If bTempToParcel Then
                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "A_Parcels"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(2) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "-GeoM"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(3) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "-KadaGIS"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(4) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "A_Temp"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(5) = pField

                Else
                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "R_Geo_Visi"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(2) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "-GeoM"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(3) = pField
                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "-KadaGIS"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(4) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "R_Prelim"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(5) = pField

                End If


                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Name_2 = "Dubl_2"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With
                pFieldsEdit.Field_2(6) = pField

                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Name_2 = "Dubl_1"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With
                pFieldsEdit.Field_2(7) = pField

                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Name_2 = "Dubl_3"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With
                pFieldsEdit.Field_2(8) = pField

                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Name_2 = "Dubl_4"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With
                pFieldsEdit.Field_2(9) = pField

                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Name_2 = "Dubl_8"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With
                pFieldsEdit.Field_2(10) = pField

                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Name_2 = "Dubl_9"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With
                pFieldsEdit.Field_2(11) = pField

                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Name_2 = "Dubl_12"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With
                pFieldsEdit.Field_2(12) = pField

                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Name_2 = "Dubl_13"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With
                pFieldsEdit.Field_2(13) = pField

                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Name_2 = "Dubl_5"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With
                pFieldsEdit.Field_2(14) = pField

                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Name_2 = "Dubl_6"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With
                pFieldsEdit.Field_2(15) = pField

                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Name_2 = "Dubl_7"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With
                pFieldsEdit.Field_2(16) = pField

                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Name_2 = "Dubl_10"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With
                pFieldsEdit.Field_2(17) = pField

                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Name_2 = "Dubl_11"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With
                pFieldsEdit.Field_2(18) = pField

                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Name_2 = "Dubl_14"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With
                pFieldsEdit.Field_2(19) = pField

                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Name_2 = "Dubl_15"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With
                pFieldsEdit.Field_2(20) = pField

                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Name_2 = "Visi_skl"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With
                pFieldsEdit.Field_2(21) = pField


            Case "I"
                pFields = New Fields
                pFieldsEdit = CType(pFields, IFieldsEdit)
                pFieldsEdit.FieldCount_2 = 7

                'OID
                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Name_2 = "OID"
                    .Type_2 = esriFieldType.esriFieldTypeOID
                End With
                pFieldsEdit.Field_2(0) = pField

                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Length_2 = 50
                    .Name_2 = SName  'operatorius
                    .Type_2 = esriFieldType.esriFieldTypeString
                End With
                pFieldsEdit.Field_2(1) = pField

                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Name_2 = "Inz1"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With
                pFieldsEdit.Field_2(2) = pField

                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Name_2 = "Inz2"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With
                pFieldsEdit.Field_2(3) = pField

                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Name_2 = "Inz1L"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With
                pFieldsEdit.Field_2(4) = pField

                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Name_2 = "Inz2L"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With
                pFieldsEdit.Field_2(5) = pField

                pField = New Field
                pFieldEdit = CType(pField, IFieldEdit2)
                With pFieldEdit
                    .Name_2 = "Visi_inz"
                    .Type_2 = esriFieldType.esriFieldTypeInteger
                End With
                pFieldsEdit.Field_2(6) = pField

        End Select
        If bTempToParcel Then
            createDBF = pFWS.CreateTable(strName & "T_P", pFields, Nothing, Nothing, "")
        Else
            createDBF = pFWS.CreateTable(strName, pFields, Nothing, Nothing, "")
        End If

        Return createDBF
        Exit Function
ERN:    MsgBox(Err.Number & " " & Err.Description)
        Return Nothing
        


    End Function


    Public Function createDBF_senas(ByVal strName As String, ByVal strFolder As String, ByVal sINZ_SKL As String, ByVal SName As String) As ITable      ' createDBF: simple function to create a DBASE file.
        ' Open the Workspace
        Dim pFWS As IFeatureWorkspace = Nothing
        Dim pWorkspaceFactory As IWorkspaceFactory = Nothing
        Dim pFieldsEdit As IFieldsEdit = Nothing
        Dim pFieldEdit As IFieldEdit2 = Nothing
        Dim pField As IField = Nothing
        Dim pFields As IFields = Nothing
        Try
            pWorkspaceFactory = New ShapefileWorkspaceFactory
            pFWS = CType(pWorkspaceFactory.OpenFromFile(strFolder, 0), IFeatureWorkspace)
            Select Case sINZ_SKL
                Case "S"
                    pFields = New Fields
                    pFieldsEdit = CType(pFields, IFieldsEdit)
                    pFieldsEdit.FieldCount_2 = 23

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "OID"  'OID
                        .Type_2 = esriFieldType.esriFieldTypeOID
                    End With
                    pFieldsEdit.Field_2(0) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Length_2 = 30
                        .Name_2 = SName 'operatorius
                        .Type_2 = esriFieldType.esriFieldTypeString
                    End With
                    pFieldsEdit.Field_2(1) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "R_GeodViso"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(2) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "R_GeoM"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(3) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "R_KadaGIS"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(4) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "R_Prelim"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(5) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "L_Patikra"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(6) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "L_Isvada"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(7) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "L_IsvadaPatikra"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(8) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "L_Laikinas"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(9) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "L_PatikraNZT"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(10) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "L_IsvadaNZT"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(11) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "L_PatikraKonsol"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(12) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "L_IsvadaKonsol"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(13) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "L_LaikinasGeoM"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(14) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "L_PatikraGeoM"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(15) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "L_IsvadaGeoM"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(16) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "L_PatikraNZTGeoM"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(17) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "L_IsvadaNZTGeoM"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(18) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "L_PatikraKonsGeoM"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(19) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "L_IsvadaKonsGeoM"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(20) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "Visi_skl"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(21) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "Taskai_geod"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(22) = pField

                Case "I"
                    pFields = New Fields
                    pFieldsEdit = CType(pFields, IFieldsEdit)
                    pFieldsEdit.FieldCount_2 = 8

                    'OID
                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "OID"
                        .Type_2 = esriFieldType.esriFieldTypeOID
                    End With
                    pFieldsEdit.Field_2(0) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Length_2 = 30
                        .Name_2 = SName  'operatorius
                        .Type_2 = esriFieldType.esriFieldTypeString
                    End With
                    pFieldsEdit.Field_2(1) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "Inz1"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(2) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "Inz2"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(3) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "Inz1L"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(4) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "Inz2L"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(5) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "Visi_inz"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(6) = pField

                    pField = New Field
                    pFieldEdit = CType(pField, IFieldEdit2)
                    With pFieldEdit
                        .Name_2 = "Taskai_geo"
                        .Type_2 = esriFieldType.esriFieldTypeInteger
                    End With
                    pFieldsEdit.Field_2(7) = pField
            End Select
            createDBF_senas = pFWS.CreateTable(strName, pFields, Nothing, Nothing, "")
            Return createDBF_senas
        Catch ex As Exception
            MsgBox(ex.Message())
            Return (Nothing)
        Finally
        End Try


    End Function
    Public Function Dubl_KadaGIS() As Boolean

        Dim pWorkspace As IWorkspace = Nothing
        Dim pFact As IWorkspaceFactory = Nothing
        Dim pFeatws As IFeatureWorkspace
        Dim pTable As ITable
        Dim pCursor As ICursor
        Dim pRow As IRow
        Dim err1 As String = ""
        Dim err2 As String = ""
        Try
            Dubl_KadaGIS = True
            sBatKelias = GetAppPath()
            If Dir(sBatKelias & "\KadaGIS.dbf") <> "" Then
                pFact = New ESRI.ArcGIS.DataSourcesFile.ShapefileWorkspaceFactory
                pWorkspace = pFact.OpenFromFile(sBatKelias, 0)
                pFeatws = CType(pWorkspace, IFeatureWorkspace)
                pTable = pFeatws.OpenTable("KadaGIS")
                If pTable.RowCount(Nothing) > 0 Then
                    'lauku(reiksmes); 'Dim sDubl As String; 'Dim sTekstas As String;  'Dim sSavGyv As Short '1 - patikra; 2 - isvada
                    ReDim DublKadaGIS(0) '
                    pCursor = pTable.Search(Nothing, False) 'Imu visus lenteles elementus
                    pRow = pCursor.NextRow
                    While Not pRow Is Nothing
                        DublKadaGIS(UBound(DublKadaGIS)).sDubl = pRow.Value(pRow.Fields.FindField("Dubl"))
                        DublKadaGIS(UBound(DublKadaGIS)).sTekstas = pRow.Value(pRow.Fields.FindField("Tekstas"))
                        DublKadaGIS(UBound(DublKadaGIS)).iSavGyv = pRow.Value(pRow.Fields.FindField("Sav_Gyv"))
                        ReDim Preserve DublKadaGIS(UBound(DublKadaGIS) + 1)
                        pRow = pCursor.NextRow
                    End While
                End If
            Else
                MsgBox("Nëra lauko 'Dubl' lentelës " & sBatKelias & "\KadaGIS.dbf") : Dubl_KadaGIS = False
            End If
        Catch ex As Exception
            Dubl_KadaGIS = False
            MsgBox("Klaida lentelë " & sBatKelias & "\KadaGIS.dbf    " & ex.Message)
        Finally
        End Try
    End Function
    Public Function Dubl_Geomatininkas() As Boolean

        Dim pWorkspace As IWorkspace = Nothing
        Dim pFact As IWorkspaceFactory = Nothing
        Dim pFeatws As IFeatureWorkspace
        Dim pTable As ITable
        Dim pCursor As ICursor
        Dim pRow As IRow
        Dim err1 As String = ""
        Dim err2 As String = ""
        Try
            Dubl_Geomatininkas = True
            sBatKelias = GetAppPath()
            If Dir(sBatKelias & "\GeoMatininkas.dbf") <> "" Then
                pFact = New ESRI.ArcGIS.DataSourcesFile.ShapefileWorkspaceFactory
                pWorkspace = pFact.OpenFromFile(sBatKelias, 0)
                pFeatws = CType(pWorkspace, IFeatureWorkspace)
                pTable = pFeatws.OpenTable("GeoMatininkas")
                If pTable.RowCount(Nothing) > 0 Then
                    'lauku(reiksmes); 'Dim sDubl As String; 'Dim sTekstas As String;  'Dim sSavGyv As Short '1 - patikra; 2 - isvada
                    ReDim DublGeomat(0) '
                    pCursor = pTable.Search(Nothing, False) 'Imu visus lenteles elementus
                    pRow = pCursor.NextRow
                    While Not pRow Is Nothing
                        DublGeomat(UBound(DublGeomat)).sDubl = pRow.Value(pRow.Fields.FindField("Dubl"))
                        DublGeomat(UBound(DublGeomat)).sTekstas = pRow.Value(pRow.Fields.FindField("Tekstas"))
                        DublGeomat(UBound(DublGeomat)).iSavGyv = pRow.Value(pRow.Fields.FindField("Sav_Gyv"))
                        ReDim Preserve DublGeomat(UBound(DublGeomat) + 1)
                        pRow = pCursor.NextRow
                    End While
                End If
            Else
                MsgBox("Nëra lauko 'Dubl' lentelës " & sBatKelias & "\GeoMatininkas.dbf") : Dubl_Geomatininkas = False
            End If
        Catch ex As Exception
            Dubl_Geomatininkas = False
            MsgBox("Klaida lentelë " & sBatKelias & "\GeoMatininkas.dbf    " & ex.Message)
            
        Finally
        End Try
    End Function

    Public Function GetAppPath() As String
        Dim l_intCharPos As Integer = 0, l_intReturnPos As Integer
        Dim l_strAppPath As String

        l_strAppPath = System.Reflection.Assembly.GetExecutingAssembly.Location()

        While (1)
            l_intCharPos = InStr(l_intCharPos + 1, l_strAppPath, "\", CompareMethod.Text)
            If l_intCharPos = 0 Then
                If Right(Mid(l_strAppPath, 1, l_intReturnPos), 1) <> "\" Then
                    Return Mid(l_strAppPath, 1, l_intReturnPos) & "\"
                Else
                    Return Mid(l_strAppPath, 1, l_intReturnPos)
                End If
                Exit Function
            End If
            l_intReturnPos = l_intCharPos
        End While
        Return ""
    End Function

    Public Function ConnectToTransactionalVersionD(ByVal server As String, ByVal instance As String, _
    ByVal user As String, ByVal password As String, ByVal database As String, ByVal version As String, ByRef bOK As Boolean) As IWorkspace2
        Dim propertySet As IPropertySet = New PropertySetClass()
        Dim pWks As IWorkspace2 = Nothing
        propertySet.SetProperty("DB_CONNECTION_PROPERTIES", server)
        '  propertySet.SetProperty("DATABASE", database)
        propertySet.SetProperty("USER", user)
        propertySet.SetProperty("INSTANCE", instance)
        propertySet.SetProperty("PASSWORD", password)
        propertySet.SetProperty("VERSION", version)
        bOK = True
        Try
            Dim factoryType As Type = Type.GetTypeFromProgID("esriDataSourcesGDB.SdeWorkspaceFactory")
            Dim workspaceFactory As IWorkspaceFactory = CType(Activator.CreateInstance(factoryType), IWorkspaceFactory)
            pWks = workspaceFactory.Open(propertySet, 0)
            If pWks Is Nothing Then
                MsgBox("Nepavyko prisijungti prie SDE") : bOK = False : Return pWks
            Else
                bOK = True : Return pWks
            End If
            Exit Function
        Catch ex As Exception
            MsgBox(ex.Message & " Nepavyko prisijungti prie SDE")
            bOK = False : ConnectToTransactionalVersionD = Nothing
        Finally
        End Try

    End Function


End Module



