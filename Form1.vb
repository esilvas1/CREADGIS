'Imports System.Data.OleDb
'Imports Oracle.DataAccess.Client
'Imports System.Data.OracleClient
Imports Oracle.ManagedDataAccess.Client
Public Class Form1


    Dim oradb As String = "" 'SERVIDOR SPARD PRODUCCION
    Dim conexion As New OracleConnection(oradb)
    Dim comandos As New OracleCommand
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Hide()
        Me.ControlBox = True
        Label25.Hide()


        Try
            'verdad
            LoadDB.Show()
            LoadDB.Text = "Conectando... ORACLE"
            conexion.Open()
            LoadDB.Close()

            LoadDB.Show()
            LoadDB.Text = "Listando... Circuitos"
            FillComboBox()
            ComboCONF_RED()
            LoadDB.Close()

            LoadDB.Show()
            LoadDB.Text = "Cargando... DataGridView"
            DataGridAalimentadores_F()
            LoadDB.Close()
            Me.Show()

            'MsgBox("!!Conectado con Exito!!", vbInformation, "DATABASE")
        Catch ex As Exception
            'falso
            ButCrearDGS.Enabled = False
            LoadDB.Close()
            MsgBox("!!Error intentando Conectar al Servidor de datos!!", vbCritical, "DATABASE")

        End Try
        Enabledtext()
    End Sub

    Private Sub ComboCONF_RED()
        ComboCR.Items.Add("Circuito")
        ComboCR.Items.Add("SubEstación")
        ComboCR.SelectedIndex = 0
    End Sub
    Private Sub DataGridAalimentadores_F()
        Dim adaptador As New OracleDataAdapter
        Dim registros As New DataSet
        Dim consulta As String

        If ComboCR.SelectedIndex = 0 Then
            consulta = "SELECT  MVL.FPARENT
                              , MVL.LENGTH
                              , MVL.KVNOM
                              , MVL.CONDUCTOR
                              , MVL.CODE AS CODIGOLINEA
                              , MVL.ELNODE1
                              , MVL.ELNODE2
                              , MVL.XPOS1
                              , MVL.YPOS1
                              , MVL.XPOS2
                              , MVL.YPOS2
                              , T.CODE AS CODIGOTRANSFORMADOR
                              , MVL.PHASES
                              , T.PHASES as FASES_TRF
                              , MVL.ORDER_
                              , NVL(TY.KVA,0) AS STRAP
                              , 1 AS iLoadTrf
                              , NVL(TC.TT2_CLASS_CARGA,0) AS classif
                              , (CASE WHEN(NVL(T.CUSTOMERS,1)=0) THEN(1) ELSE(NVL(T.CUSTOMERS,1)) END) AS NrCust
                              , NVL(T.UCCAP14,'NA') AS chr_name
                              , NVL(TC.TT2_CODE_IUA,0) AS sernum
                              , (CASE WHEN(T.CODE IS NOT NULL) 
                              THEN( 'PROP:'||T.OWNER1||','||
                                    'TIPS:'||T.TIPOSUB||','||
                                    'GR_C:'||T.GRUPO015||','||
                                    'MUN:'||T.MUNICIPIO||','||
                                    'POB:'||T.POBLACION||','||
                                    'FEC_INS:'||T.DATE_INST||','||
                                    'MARK:'||T.MARCA||','||
                                    'TIPR:'||T.TIPO_RED 
                                   ) 
                              ELSE(NULL) 
                              END) AS DESCRIPCION_TRF
                             , MVL.LAT1
                             , MVL.LON1
                             , MVL.LAT2
                             , MVL.LON2
                             , T.LATITUD
                             , T.LONGITUD
                        FROM SPARD.MVLINSEC MVL
                        LEFT OUTER JOIN SPARD.TRANSFOR T
                        ON MVL.ELNODE1=T.ELNODE OR MVL.ELNODE2=T.ELNODE
                        LEFT OUTER JOIN SPARD.TRFTYPES TY 
                        ON TY.CODE = T.TRFTYPE
                        LEFT OUTER JOIN (SELECT TT2_CODIGOELEMENTO,TT2_CLASS_CARGA,TT2_CODE_IUA FROM BRAE.QA_TTT2_REGISTRO) TC 
                        ON TC.TT2_CODIGOELEMENTO = T.CODE
                        WHERE MVL.FPARENT =  '" + ComboBox1.Text + "'
                        ORDER BY MVL.CODE"

            adaptador = New OracleDataAdapter(consulta, conexion)
            registros.Tables.Add("ALIMENTADORES")
            adaptador.Fill(registros.Tables("ALIMENTADORES"))
            DataGridView1.DataSource = registros.Tables("ALIMENTADORES")
            ButCrearDGS.Enabled = True
            DataGridView1.ColumnHeadersVisible = True
            LabelCantidadRegistros.Text = DataGridView1.RowCount - 1 & " Registros"
        End If

        If ComboCR.SelectedIndex = 1 Then
            consulta = "SELECT SRC.SUBSTATION
                      , MVL.LENGTH
                      , MVL.KVNOM
                      , MVL.CONDUCTOR
                      , MVL.CODE
                      , MVL.ELNODE1
                      , MVL.ELNODE2
                      , MVL.XPOS1
                      , MVL.YPOS1
                      , MVL.XPOS2
                      , MVL.YPOS2
                      , T.CODE
                      , MVL.PHASES
                      , T.PHASES
                      , MVL.ORDER_
                      , NVL(TY.KVA,0) AS STRAP
                      , 1 AS iLoadTrf
                      , NVL(TC.TT2_CLASS_CARGA,0) AS classif
                      , (CASE WHEN(NVL(T.CUSTOMERS,1)=0) THEN(1) ELSE(NVL(T.CUSTOMERS,1)) END) AS NrCust
                      , NVL(T.UCCAP14,'NA') AS chr_name
                      , NVL(TC.TT2_CODE_IUA,0) AS sernum
                      , (CASE WHEN(T.CODE IS NOT NULL) 
                              THEN( 'PROP:'||T.OWNER1||','||
                                    'TIPS:'||T.TIPOSUB||','||
                                    'GR_C:'||T.GRUPO015||','||
                                    'MUN:'||T.MUNICIPIO||','||
                                    'POB:'||T.POBLACION||','||
                                    'FEC_INS:'||T.DATE_INST||','||
                                    'MARK:'||T.MARCA||','||
                                    'TIPR:'||T.TIPO_RED 
                                   ) 
                              ELSE(NULL) 
                              END) AS DESCRIPCION_TRF
                        , MVL.LAT1
                        , MVL.LON1
                        , MVL.LAT2
                        , MVL.LON2  
                        , T.LATITUD
                        , T.LONGITUD
                        FROM MVLINSEC MVL
                        LEFT OUTER JOIN TRANSFOR T
                        ON MVL.ELNODE1=T.ELNODE OR MVL.ELNODE2=T.ELNODE
                        LEFT OUTER JOIN SPARD.TRFTYPES TY 
                        ON TY.CODE = T.TRFTYPE
                        LEFT OUTER JOIN SPARD.FEEDERS F
                        ON F.CODE =  MVL.FPARENT
                        LEFT OUTER JOIN SPARD.SRCBUSES SRC
                        ON SRC.CODE = F.SOURCEBUS
                        LEFT OUTER JOIN (SELECT TT2_CODIGOELEMENTO,TT2_CLASS_CARGA,TT2_CODE_IUA FROM BRAE.QA_TTT2_REGISTRO) TC 
                        ON TC.TT2_CODIGOELEMENTO = T.CODE
                        WHERE SRC.SUBSTATION =  '" + ComboBox1.Text + "'
                        AND MVL.KVNOM <> '115'
                       "

            adaptador = New OracleDataAdapter(consulta, conexion)
            registros.Tables.Add("SUBS")
            adaptador.Fill(registros.Tables("SUBS"))
            DataGridView1.DataSource = registros.Tables("SUBS")
            ButCrearDGS.Enabled = True
            DataGridView1.ColumnHeadersVisible = True
            LabelCantidadRegistros.Text = DataGridView1.RowCount - 1 & " Registros"
        End If


        LabelP1.Text = ""
        LabelP2.Text = ""
        LabelP3.Text = ""
        LabelP4.Text = ""
        LabelP5.Text = ""
        LabelP6.Text = ""
        LabelP7.Text = ""
        LabelP8.Text = ""
        LabelP9.Text = ""
        LabelP10.Text = ""
        LabelP11.Text = ""
        LabelP12.Text = ""
        LabelP13.Text = ""

    End Sub
    Private Sub FillComboBox()

        Dim da As OracleDataAdapter
        Dim ds As DataSet
        Dim tables As DataTableCollection
        Dim mySQLStrg As String
        ds = New DataSet
        tables = ds.Tables

        If ComboCR.SelectedIndex = 0 Then

            mySQLStrg = "SELECT DISTINCT FPARENT 
                     FROM MVLINSEC 
                     WHERE FPARENT IS NOT NULL 
                     ORDER BY FPARENT "
            da = New OracleDataAdapter(mySQLStrg, conexion)
            da.Fill(ds, "ALIMENTADORES")
            Dim view1 As New DataView(tables(0))
            With ComboBox1
                .DataSource = ds.Tables("ALIMENTADORES")
                .DisplayMember = "FPARENT"
                .ValueMember = "FPARENT"
                .SelectedIndex = 0
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .AutoCompleteSource = AutoCompleteSource.ListItems
            End With
        End If

        If ComboCR.SelectedIndex = 1 Then

            mySQLStrg = "SELECT DISTINCT SUBSTATION 
                         FROM SPARD.SRCBUSES
                         ORDER BY 1"
            da = New OracleDataAdapter(mySQLStrg, conexion)
            da.Fill(ds, "DATOS")
            Dim view1 As New DataView(tables(0))
            With ComboBox1
                .DataSource = ds.Tables("DATOS")
                .DisplayMember = "SUBSTATION"
                .ValueMember = "SUBSTATION"
                .SelectedIndex = 0
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .AutoCompleteSource = AutoCompleteSource.ListItems
            End With
        End If


    End Sub
    Private Sub Enabledtext()

        TextAlimentador.ReadOnly = True
        TextCalibre.ReadOnly = True
        TextCodPHL.ReadOnly = True
        TextCodSL.ReadOnly = True
        TextCoorXNPHF.ReadOnly = True
        TextCoorXNPHI.ReadOnly = True
        TextCoorYNPHF.ReadOnly = True
        TextCoorYNPHI.ReadOnly = True
        TextLong.ReadOnly = True
        TextNPHF.ReadOnly = True
        TextNPHI.ReadOnly = True


        If CheckBox1.Checked = False Then
            TextBox1.ReadOnly = True
            TextBox2.ReadOnly = True
            TextBox3.ReadOnly = True
            TextBox4.ReadOnly = True
            TextBox5.ReadOnly = True
            TextBox6.ReadOnly = True
        End If
    End Sub

    Private Sub ButFiltrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        DataGridAalimentadores_F()

    End Sub
    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Enabledtext()
        Filltext()
    End Sub

    Private Sub DataGridView1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGridView1.SelectionChanged
        Enabledtext()
        Filltext()
    End Sub

    Private Sub Filltext()
        Dim i As New Integer

        i = DataGridView1.CurrentRow.Index

        If Not String.IsNullOrEmpty(Me.DataGridView1.Item(0, i).Value().ToString) Then
            TextAlimentador.Text = DataGridView1.Item(0, i).Value()
        Else
            TextAlimentador.Text = " "
        End If

        If Not String.IsNullOrEmpty(Me.DataGridView1.Item(1, i).Value().ToString) Then
            TextLong.Text = DataGridView1.Item(1, i).Value()
        Else
            TextLong.Text = " "
        End If

        If Not String.IsNullOrEmpty(Me.DataGridView1.Item(3, i).Value().ToString) Then
            TextCalibre.Text = DataGridView1.Item(3, i).Value()
        Else
            TextCalibre.Text = " "
        End If

        If Not String.IsNullOrEmpty(Me.DataGridView1.Item(2, i).Value().ToString) Then
            TextTension.Text = DataGridView1.Item(2, i).Value()
        Else
            TextTension.Text = " "
        End If

        If Not String.IsNullOrEmpty(Me.DataGridView1.Item(4, i).Value().ToString) Then
            TextCodSL.Text = DataGridView1.Item(4, i).Value()
        Else
            TextCodSL.Text = " "
        End If

        If Not String.IsNullOrEmpty(Me.DataGridView1.Item(5, i).Value().ToString) Then
            TextNPHI.Text = DataGridView1.Item(5, i).Value()
        Else
            TextNPHI.Text = " "
        End If

        If Not String.IsNullOrEmpty(Me.DataGridView1.Item(6, i).Value().ToString) Then
            TextNPHF.Text = DataGridView1.Item(6, i).Value()
        Else
            TextNPHF.Text = " "
        End If

        If Not String.IsNullOrEmpty(Me.DataGridView1.Item(12, i).Value().ToString) Then
            TextCodPHL.Text = DataGridView1.Item(12, i).Value()
        Else
            TextCodPHL.Text = " "
        End If

        If Not String.IsNullOrEmpty(Me.DataGridView1.Item(7, i).Value().ToString) Then
            TextCoorXNPHI.Text = DataGridView1.Item(7, i).Value()
        Else
            TextCoorXNPHI.Text = " "
        End If

        If Not String.IsNullOrEmpty(Me.DataGridView1.Item(8, i).Value().ToString) Then
            TextCoorYNPHI.Text = DataGridView1.Item(8, i).Value()
        Else
            TextCoorYNPHI.Text = " "
        End If

        If Not String.IsNullOrEmpty(Me.DataGridView1.Item(9, i).Value().ToString) Then
            TextCoorXNPHF.Text = DataGridView1.Item(9, i).Value()
        Else
            TextCoorXNPHF.Text = " "
        End If

        If Not String.IsNullOrEmpty(Me.DataGridView1.Item(10, i).Value().ToString) Then
            TextCoorYNPHF.Text = DataGridView1.Item(10, i).Value()
        Else
            TextCoorYNPHF.Text = " "
        End If

    End Sub

    Private Sub ButCrearDGS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButCrearDGS.Click
        System.Threading.Thread.Sleep(1000)
        Application.DoEvents()
        If CheckBox1.Checked = True Then
            If TextBox1.Text = "" Then
                MsgBox("!!Parametros en Vacio!! Desmarque la casilla del tamaño", vbObjectError, "Warning")
                Exit Sub
            End If
        End If

        Dim ApExcel = New Microsoft.Office.Interop.Excel.Application
        Dim libro = ApExcel.Workbooks.Add
        Dim Hoja = ApExcel.Worksheets.Add
        Hoja = ApExcel.Worksheets.Add
        Hoja = ApExcel.Worksheets.Add
        Hoja = ApExcel.Worksheets.Add
        Hoja = ApExcel.Worksheets.Add
        Hoja = ApExcel.Worksheets.Add
        Hoja = ApExcel.Worksheets.Add


        Label25.Text = " "
        ButCrearDGS.Enabled = False



        libro.Sheets(1).Name = "General"
        libro.Sheets(2).Name = "ElmNet"
        libro.Sheets(3).Name = "IntGrfnet"
        libro.Sheets(4).Name = "ElmTerm"
        libro.Sheets(5).Name = "IntGrf"
        libro.Sheets(6).Name = "ElmLod"
        libro.Sheets(7).Name = "ElmLne"
        libro.Sheets(8).Name = "StaCubic"

        'HOJA GENERAL------------------------------------------------------------------------------
        libro.Sheets("General").cells(1, 1).Font.Bold = True
        libro.Sheets("General").cells(1, 1) = "ID(a:40)"
        libro.Sheets("General").cells(1, 2).Font.Bold = True
        libro.Sheets("General").cells(1, 2) = "Descr(a:40)"
        libro.Sheets("General").cells(1, 3).Font.Bold = True
        libro.Sheets("General").cells(1, 3) = "Val(a:40)"
        libro.Sheets("General").cells(2, 1) = "1"
        libro.Sheets("General").cells(2, 2) = "Version"
        libro.Sheets("General").cells(2, 3).NumberFormat = "@"
        libro.Sheets("General").cells(2, 3) = "5.0"
        LabelP1.Text = "ID(a:40)"
        LabelP2.Text = "Descr(a:40)"
        LabelP3.Text = "Val(a:40)"
        LabelP4.Text = "1"
        LabelP5.Text = "Version"
        LabelP6.Text = "5.0"
        LabelP7.Text = "0"
        ProgressBar.Value = 5
        LabelPorcentaje.Text = "5%"


        'HOJA ElmNet------------------------------------------------------------------------------ 
        libro.Sheets("ElmNet").cells(1, 1).font.bold = True
        libro.Sheets("ElmNet").cells(1, 1) = "ID(a:40)"
        libro.Sheets("ElmNet").cells(1, 2).font.bold = True
        libro.Sheets("ElmNet").cells(1, 2) = "loc_name(a:40)"
        libro.Sheets("ElmNet").cells(1, 3).font.bold = True
        libro.Sheets("ElmNet").cells(1, 3) = "fold_id(p)"
        libro.Sheets("ElmNet").cells(1, 4).font.bold = True
        libro.Sheets("ElmNet").cells(1, 4) = "frnom(r)"
        libro.Sheets("ElmNet").cells(1, 5).font.bold = True
        libro.Sheets("ElmNet").cells(1, 5) = "for_name"
        libro.Sheets("ElmNet").cells(1, 6).font.bold = True
        libro.Sheets("ElmNet").cells(1, 6) = "chr_name"
        libro.Sheets("ElmNet").cells(2, 1).NumberFormat = "0.00"
        libro.Sheets("ElmNet").cells(2, 1) = "2.00"
        libro.Sheets("ElmNet").cells(2, 2) = "RED_" & ComboBox1.Text & ""
        libro.Sheets("ElmNet").cells(2, 3) = ""
        libro.Sheets("ElmNet").cells(2, 4) = "60"
        libro.Sheets("ElmNet").cells(2, 5) = "RED_" & ComboBox1.Text & ""
        libro.Sheets("ElmNet").cells(2, 6) = "RED_" & ComboBox1.Text & ""
        LabelP1.Text = "ID(a:40)"
        LabelP2.Text = "loc_name(a:40)"
        LabelP3.Text = "fold_id(p)"
        LabelP4.Text = "frnom(r)"
        LabelP5.Text = "for_name"
        LabelP6.Text = "chr_name"
        LabelP7.Text = "RED_" & ComboBox1.Text & ""
        LabelP8.Text = " "
        LabelP9.Text = " "
        LabelP10.Text = " "
        LabelP11.Text = " "
        LabelP12.Text = " "
        LabelP13.Text = " "
        ProgressBar.Value = 6
        LabelPorcentaje.Text = "6%"



        'HOJA IntGrfnet------------------------------------------------------------------------------ 
        libro.Sheets("IntGrfnet").cells(1, 1).font.bold = True
        libro.Sheets("IntGrfnet").cells(1, 1) = "ID(a:40)"
        libro.Sheets("IntGrfnet").cells(1, 2).font.bold = True
        libro.Sheets("IntGrfnet").cells(1, 2) = "loc_name(a:40)"
        libro.Sheets("IntGrfnet").cells(1, 3).font.bold = True
        libro.Sheets("IntGrfnet").cells(1, 3) = "fold_id(p)"
        libro.Sheets("IntGrfnet").cells(1, 4).font.bold = True
        libro.Sheets("IntGrfnet").cells(1, 4) = "grid_on(i)"
        libro.Sheets("IntGrfnet").cells(1, 5).font.bold = True
        libro.Sheets("IntGrfnet").cells(1, 5) = "ortho_on(i)"
        libro.Sheets("IntGrfnet").cells(1, 6).font.bold = True
        libro.Sheets("IntGrfnet").cells(1, 6) = "snap_on(i)"
        libro.Sheets("IntGrfnet").cells(1, 7).font.bold = True
        libro.Sheets("IntGrfnet").cells(1, 7) = "pDataFolder(p)"
        libro.Sheets("IntGrfnet").cells(1, 8).font.bold = True
        libro.Sheets("IntGrfnet").cells(1, 8) = "for_name"
        libro.Sheets("IntGrfnet").cells(2, 1).NumberFormat = "0.00"
        libro.Sheets("IntGrfnet").cells(2, 1) = "3.00"
        libro.Sheets("IntGrfnet").cells(2, 2) = ComboBox1.Text
        libro.Sheets("IntGrfnet").cells(2, 3) = ""
        libro.Sheets("IntGrfnet").cells(2, 4) = "0"
        libro.Sheets("IntGrfnet").cells(2, 5) = "0"
        libro.Sheets("IntGrfnet").cells(2, 6) = "0"
        libro.Sheets("IntGrfnet").cells(2, 7).NumberFormat = "0.00"
        libro.Sheets("IntGrfnet").cells(2, 7) = "2.00"
        libro.Sheets("IntGrfnet").cells(2, 8) = ComboBox1.Text

        LabelP1.Text = "ID(a:40)"
        LabelP2.Text = "loc_name(a:40)"
        LabelP3.Text = "fold_id(p)"
        LabelP4.Text = "grid_on(i)"
        LabelP5.Text = "ortho_on(i)"
        LabelP6.Text = "snap_on(i)"
        LabelP7.Text = "RED_" & ComboBox1.Text & ""
        LabelP8.Text = " "
        LabelP9.Text = " "
        LabelP10.Text = " "
        LabelP11.Text = " "
        LabelP12.Text = " "
        LabelP13.Text = " "
        ProgressBar.Value = 7
        LabelPorcentaje.Text = "7%"




        'HOJA ElmTerm------------------------------------------------------------------------------
        libro.Sheets("ElmTerm").cells(1, 1).font.bold = True
        libro.Sheets("ElmTerm").cells(1, 1) = "ID(a:40)"
        libro.Sheets("ElmTerm").cells(1, 2).font.bold = True
        libro.Sheets("ElmTerm").cells(1, 2) = "loc_name(a:40)"
        libro.Sheets("ElmTerm").cells(1, 3).font.bold = True
        libro.Sheets("ElmTerm").cells(1, 3) = "fold_id(p)"
        libro.Sheets("ElmTerm").cells(1, 4).font.bold = True
        libro.Sheets("ElmTerm").cells(1, 4) = "typ_id(p)"
        libro.Sheets("ElmTerm").cells(1, 5).font.bold = True
        libro.Sheets("ElmTerm").cells(1, 5) = "iUsage(i)"
        libro.Sheets("ElmTerm").cells(1, 6).font.bold = True
        libro.Sheets("ElmTerm").cells(1, 6) = "phtech"
        libro.Sheets("ElmTerm").cells(1, 7).font.bold = True
        libro.Sheets("ElmTerm").cells(1, 7) = "uknom(r)"
        libro.Sheets("ElmTerm").cells(1, 8).font.bold = True
        libro.Sheets("ElmTerm").cells(1, 8) = "chr_name(a:20)"
        libro.Sheets("ElmTerm").cells(1, 9).font.bold = True
        libro.Sheets("ElmTerm").cells(1, 9) = "outserv(i)"
        libro.Sheets("ElmTerm").cells(1, 10).font.bold = True
        libro.Sheets("ElmTerm").cells(1, 10) = "for_name(a:40)"
        libro.Sheets("ElmTerm").cells(1, 11).font.bold = True
        libro.Sheets("ElmTerm").cells(1, 11) = "GPSlat(r)"
        libro.Sheets("ElmTerm").cells(1, 12).font.bold = True
        libro.Sheets("ElmTerm").cells(1, 12) = "GPSlon(r)"

        'HOJA IntGrf------------------------------------------------------------------------------
        libro.Sheets("IntGrf").cells(1, 1).font.bold = True
        libro.Sheets("IntGrf").cells(1, 1) = "ID(a:40)"
        libro.Sheets("IntGrf").cells(1, 2).font.bold = True
        libro.Sheets("IntGrf").cells(1, 2) = "loc_name(a:40)"
        libro.Sheets("IntGrf").cells(1, 3).font.bold = True
        libro.Sheets("IntGrf").cells(1, 3) = "fold_id(p)"
        libro.Sheets("IntGrf").cells(1, 4).font.bold = True
        libro.Sheets("IntGrf").cells(1, 4) = "iCol(i)"
        libro.Sheets("IntGrf").cells(1, 5).font.bold = True
        libro.Sheets("IntGrf").cells(1, 5) = "iVis(i)"
        libro.Sheets("IntGrf").cells(1, 6).font.bold = True
        libro.Sheets("IntGrf").cells(1, 6) = "iLevel(i)"
        libro.Sheets("IntGrf").cells(1, 7).font.bold = True
        libro.Sheets("IntGrf").cells(1, 7) = "rCenterX(r)"
        libro.Sheets("IntGrf").cells(1, 8).font.bold = True
        libro.Sheets("IntGrf").cells(1, 8) = "rCenterY(r)"
        libro.Sheets("IntGrf").cells(1, 9).font.bold = True
        libro.Sheets("IntGrf").cells(1, 9) = "sSymNam(a:40)"
        libro.Sheets("IntGrf").cells(1, 10).font.bold = True
        libro.Sheets("IntGrf").cells(1, 10) = "pDataObj(p)"
        libro.Sheets("IntGrf").cells(1, 11).font.bold = True
        libro.Sheets("IntGrf").cells(1, 11) = "iRot(i)"
        libro.Sheets("IntGrf").cells(1, 12).font.bold = True
        libro.Sheets("IntGrf").cells(1, 12) = "rSizeX(r)"
        libro.Sheets("IntGrf").cells(1, 13).font.bold = True
        libro.Sheets("IntGrf").cells(1, 13) = "rSizeY(r)"
        libro.Sheets("IntGrf").cells(1, 14).font.bold = True
        libro.Sheets("IntGrf").cells(1, 14) = "sAttr:SIZEROW(i)"
        libro.Sheets("IntGrf").cells(1, 15).font.bold = True
        libro.Sheets("IntGrf").cells(1, 15) = "sAttr:0(a)"
        libro.Sheets("IntGrf").cells(1, 16).font.bold = True
        libro.Sheets("IntGrf").cells(1, 16) = "sAttr:1(a)"
        libro.Sheets("IntGrf").cells(1, 17).font.bold = True
        libro.Sheets("IntGrf").cells(1, 17) = "sAttr:2(a)"
        libro.Sheets("IntGrf").cells(1, 18).font.bold = True
        libro.Sheets("IntGrf").cells(1, 18) = "sAttr:3(a)"
        libro.Sheets("IntGrf").cells(1, 19).font.bold = True
        libro.Sheets("IntGrf").cells(1, 19) = "sAttr:4(a)"
        libro.Sheets("IntGrf").cells(1, 20).font.bold = True
        libro.Sheets("IntGrf").cells(1, 20) = "for_name"


        'HOJA ElmLod------------------------------------------------------------------------------
        libro.Sheets("ElmLod").cells(1, 1).font.bold = True
        libro.Sheets("ElmLod").cells(1, 1) = "ID(a:40)"
        libro.Sheets("ElmLod").cells(1, 2).font.bold = True
        libro.Sheets("ElmLod").cells(1, 2) = "loc_name(a:40)"
        libro.Sheets("ElmLod").cells(1, 3).font.bold = True
        libro.Sheets("ElmLod").cells(1, 3) = "fold_id(p)"
        libro.Sheets("ElmLod").cells(1, 4).font.bold = True
        libro.Sheets("ElmLod").cells(1, 4) = "bus1(p)"
        libro.Sheets("ElmLod").cells(1, 5).font.bold = True
        libro.Sheets("ElmLod").cells(1, 5) = "typ_id(p)"
        libro.Sheets("ElmLod").cells(1, 6).font.bold = True
        libro.Sheets("ElmLod").cells(1, 6) = "mode_inp"
        libro.Sheets("ElmLod").cells(1, 7).font.bold = True
        libro.Sheets("ElmLod").cells(1, 7) = "chr_name(a:20)"
        libro.Sheets("ElmLod").cells(1, 8).font.bold = True
        libro.Sheets("ElmLod").cells(1, 8) = "plini(r)"
        libro.Sheets("ElmLod").cells(1, 9).font.bold = True
        libro.Sheets("ElmLod").cells(1, 9) = "qlini(r)"
        libro.Sheets("ElmLod").cells(1, 10).font.bold = True
        libro.Sheets("ElmLod").cells(1, 10) = "scale0(r)"
        libro.Sheets("ElmLod").cells(1, 11).font.bold = True
        libro.Sheets("ElmLod").cells(1, 11) = "desc:0(a)"
        libro.Sheets("ElmLod").cells(1, 12).font.bold = True
        libro.Sheets("ElmLod").cells(1, 12) = "iLoadTrf(i)"
        libro.Sheets("ElmLod").cells(1, 13).font.bold = True
        libro.Sheets("ElmLod").cells(1, 13) = "Strat(r)"
        libro.Sheets("ElmLod").cells(1, 14).font.bold = True
        libro.Sheets("ElmLod").cells(1, 14) = "classif(a:20)"
        libro.Sheets("ElmLod").cells(1, 15).font.bold = True
        libro.Sheets("ElmLod").cells(1, 15) = "NrCust(i)"
        libro.Sheets("ElmLod").cells(1, 16).font.bold = True
        libro.Sheets("ElmLod").cells(1, 16) = "sernum(a:20)"
        libro.Sheets("ElmLod").cells(1, 17).font.bold = True
        libro.Sheets("ElmLod").cells(1, 17) = "for_name(a:20)"
        libro.Sheets("ElmLod").cells(1, 18).font.bold = True
        libro.Sheets("ElmLod").cells(1, 18) = "GPSlat(r)"
        libro.Sheets("ElmLod").cells(1, 19).font.bold = True
        libro.Sheets("ElmLod").cells(1, 19) = "GPSlon(r)"
        libro.Sheets("ElmLod").cells(1, 20).font.bold = True
        libro.Sheets("ElmLod").cells(1, 20) = "for_name"




        'HOJA ElmLne------------------------------------------------------------------------------
        libro.Sheets("ElmLne").cells(1, 1).font.bold = True
        libro.Sheets("ElmLne").cells(1, 1) = "ID(a:40)"
        libro.Sheets("ElmLne").cells(1, 2).font.bold = True
        libro.Sheets("ElmLne").cells(1, 2) = "loc_name(a:40)"
        libro.Sheets("ElmLne").cells(1, 3).font.bold = True
        libro.Sheets("ElmLne").cells(1, 3) = "fold_id(p)"
        libro.Sheets("ElmLne").cells(1, 4).font.bold = True
        libro.Sheets("ElmLne").cells(1, 4) = "bus1(p)"
        libro.Sheets("ElmLne").cells(1, 5).font.bold = True
        libro.Sheets("ElmLne").cells(1, 5) = "bus2(p)"
        libro.Sheets("ElmLne").cells(1, 6).font.bold = True
        libro.Sheets("ElmLne").cells(1, 6) = "typ_id(p)"
        libro.Sheets("ElmLne").cells(1, 7).font.bold = True
        libro.Sheets("ElmLne").cells(1, 7) = "dline(r)"
        libro.Sheets("ElmLne").cells(1, 8).font.bold = True
        libro.Sheets("ElmLne").cells(1, 8) = "chr_name(a:20)"
        libro.Sheets("ElmLne").cells(1, 9).font.bold = True
        libro.Sheets("ElmLne").cells(1, 9) = "for_name(a:20)"




        'HOJA StaCubic------------------------------------------------------------------------------
        libro.Sheets("StaCubic").cells(1, 1).font.bold = True
        libro.Sheets("StaCubic").cells(1, 1) = "ID(a:40)"
        libro.Sheets("StaCubic").cells(1, 2).font.bold = True
        libro.Sheets("StaCubic").cells(1, 2) = "loc_name(a:40)"
        libro.Sheets("StaCubic").cells(1, 3).font.bold = True
        libro.Sheets("StaCubic").cells(1, 3) = "fold_id(p)"
        libro.Sheets("StaCubic").cells(1, 4).font.bold = True
        libro.Sheets("StaCubic").cells(1, 4) = "obj_id(p)"
        libro.Sheets("StaCubic").cells(1, 5).font.bold = True
        libro.Sheets("StaCubic").cells(1, 5) = "chr_name(a:20)"
        libro.Sheets("StaCubic").cells(1, 6).font.bold = True
        libro.Sheets("StaCubic").cells(1, 6) = "for_name"
        libro.Sheets("StaCubic").cells(1, 7).font.bold = True
        libro.Sheets("StaCubic").cells(1, 7) = "it2p1"
        libro.Sheets("StaCubic").cells(1, 8).font.bold = True
        libro.Sheets("StaCubic").cells(1, 8) = "it2p2"

        '______________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________________
        'operaciones de vectores y matrices con el objetivo de crear campos de columnas en excel

        Dim i, j, k, m, contador, size, consecutivo1, consecutivo2, consecutivo3, consecutivo4, consecutivo5, consecutivoTRF, consecutivoTRF2, consecutivoTRF3 As New Integer
        Dim var1 As Integer      'variable para el progreso de la barra de espera
        Dim var2 As Double       'variable para el progreso de la barra de espera
        Dim coordX1() As Double ' vector que contiene coordenadas en x
        Dim coordX() As Double  ' vector que contiene coordenadas en x ya validado sin espacios y sin duplicados, con propositos de hallar parametros de medida de coordenadas digsilent
        Dim coordY1() As Double  'vector que contiene coordenadas en y
        Dim coordY() As Double   'vector que contiene coordenadas en y ya validado sin espacion y sin duplicados, con propositos de hallar parametros de medida de coordenadas digsilent
        Dim x() As String        'vector que contiene los todos nodos del circuito
        Dim y() As String        'vector que contiene los todos nodos del circuito ya validado sin repeticiones y sin espacios vacios
        Dim TempX() As Double    'vector que contiene coordenadas en x de manera temporal, con propositos de sincronizar con vector de nodos y()
        Dim TempY() As Double    'vector que contiene coordenadas en y de manera temporal, con propositos de sincronizar con vector de nodos y()
        Dim TRF_TEMP() As String 'vector que contiene todos los nombres de los trafos del circuito seleccionado en el momento del filtro, no necesita estar sincronizado con el vector de nodos y(), de manera temporal antes de filtrar
        Dim TRF() As String      'vector que contiene todos los nombres de los trafos del circuito seleccionado en el momento del filtro, no necesita estar sincronizado con el vector de nodos y(), despues de haberse validado
        Dim NodoI() As String    'vector que contiene todos los nombres de los nodos iniciales correspondientes a cada linea
        Dim NodoF() As String    'vector que contiene todos los nombres de los nodos finales correspondientes a cadad linea
        Dim Line_Temp() As String 'vector que contiene todas las lineas "TYPLIN_NODOI-NODOF" Complementado o creado a partir de los vectores nodoi y nodof, vector temporal sin filtrar duplicados, y bajantes.
        Dim line() As String      'vector que contiene todas las lineas filtradas sin espacio en blanco, duplicados ni bajantes, contenido del vector Line_Temp(), con tamaño al contador de lineas resultantes
        Dim Sline_Temp() As String 'vector que contiene todos los calibres de las conductor  
        Dim Sline() As String     'vector que contiene todos los calibres de los conductores, equal line()
        Dim SlineAmp() As String  'vector que contiene la capacidad en amperios de cada conductor
        Dim longitud_temp() As Double 'vector que contiene las logitud de cada linea de manera temporal
        Dim longitud() As Double   ' vector que contiene las longitudes de todas las lineas ya filtradas y validadas
        Dim nodo1() As String     'contiene los nodos en la posicion inicial en el orden en que se encuentra la organizacion de las lineas, el cual indica el nodo inicial sin importar repeticiones 
        Dim nodo2() As String     'contiene los nodos en la posicion final   en el orden en que se encuentar la organizacion de las lineas, el cual indica el nodo inicial sin importar repeticiones
        Dim PositionsNodo1() As String ' contiene las posiciones de los nodos iniciales en el orden de las lineas
        Dim PositionsNodo2() As String ' contiene las posiciones de los nodos finales en el orden de las lineas
        Dim NodoTRF_Temp() As String        'contiene los nodos que corresponden a los trafos asignados, de manera temporal antes de validar espacios vacios y tamaño del vector contenedor
        Dim NodoTRF() As String             'contiene los nodos que corresponden a los trafos asignados, ya validados sin espacion ni repeticiones
        Dim positionsNTRF() As String       'contiene las posiciones de los nodos que contienen trafos conectados en sus cubiculos
        Dim PhaseLine_Temp() As Integer          'vector que contiene las fases abc de las lineas, de manaera temporal mientras se seleccionan las lineas a graficar excluyendo los bajantes y las duplicadas
        Dim PhaseLine() As Integer                'vector que contiene las fases abc de las lineas ya validadas
        Dim TipoTRF1() As String         'contien los tipos de los trafos
        Dim TipoTRF2() As String         'contiene los tipos de transformadores
        Dim TipoTRF3() As String   'contiene los tipos de los transformadores
        Dim it2p1_TRF() As String  'almacena los valores de la matriz del controlo de flujo de los transformadodres conectados
        Dim it2p2_TRF() As String  'almacena los valores de la matriz del control de flujo de los transformadores conectados
        Dim it2p1_Line() As String 'almacena los valores de la matriz del control de flujo de las lineas 
        Dim it2p2_Line() As String 'almacena los valores de la matriz del control de flujo de las lineas
        Dim nlnph() As String  'contiene el numero de lineas por cada tramo de linea dependiendo la fases contenidas en el vecotor PhaseLine
        Dim nneutral() As String 'contiene datos de la linea con respecto a si esta contiene neutro, 0 no contiene neutro en sus fases, 1 contiene un neutro en sus fases
        Dim TensionNominalNodos() As String 'contiene las tension nominal de todos los nodos
        Dim TensionNominalLines() As String 'contiene las tensiones nominal de todas las lineas
        Dim TensionNominalNodos1() As String 'contiene la tension nominal de los nodos resultantes del filtro
        Dim TensionNominalLines1() As String 'contiene la tenions nominal de las lineas resultantes del filtro
        Dim phtech() As String 'contienen las tencnologias de los nodos como ABC,BI,1PH
        Dim PhaseNodo_1() As String 'contiene las fases correspondiente a la linea que pertenecen los nodos 
        Dim PhaseNodo_2() As String 'contiene las fases correspondiente a la linea que pertenecen los nodos 
        Dim Order_() As String 'contine los valores de linea inversa o en direccion con respecto al flujo
        Dim CodeLine_temp() As String ' contiene todos los nombres de secciones de lineas mvl
        Dim CodeLine() As String
        Dim desc_temp() As String
        Dim iLoadTrf_temp() As String
        Dim Strat_temp() As String
        Dim classif_temp() As String
        Dim NrCust_temp() As String
        Dim chr_name_temp() As String
        Dim sernum_temp() As String
        Dim desc() As String
        Dim iLoadTrf() As String
        Dim Strat() As String
        Dim classif() As String
        Dim NrCust() As String
        Dim chr_name() As String
        Dim sernum() As String
        Dim LAT1_temp() As String
        Dim LON1_temp() As String
        Dim LAT1() As String
        Dim LON1() As String
        Dim LAT_TRF_TEMP() As String
        Dim LON_TRF_TEMP() As String
        Dim LAT_TRF() As String
        Dim LON_TRF() As String





        Dim RowCount As New Integer
        RowCount = DataGridView1.RowCount
        size = RowCount * 2
        ReDim x(size)
        ReDim coordX1(size)
        ReDim coordX(size - 3)
        ReDim coordY1(size)
        ReDim coordY(size - 3)
        ReDim TRF_TEMP(RowCount)
        ReDim NodoI(RowCount)
        ReDim NodoF(RowCount)
        ReDim LAT1_temp(RowCount * 2)
        ReDim LON1_temp(RowCount * 2)
        ReDim LAT_TRF_TEMP(RowCount)
        ReDim LON_TRF_TEMP(RowCount)
        ReDim Line_Temp(RowCount - 2)
        ReDim Sline_Temp(RowCount - 2)
        ReDim longitud_temp(RowCount)
        ReDim PositionsNodo1(RowCount)
        ReDim PositionsNodo2(RowCount)
        ReDim positionsNTRF(RowCount - 2)
        ReDim PhaseLine_Temp(RowCount)
        ReDim PhaseLine(RowCount)
        ReDim NodoTRF_Temp(RowCount)
        ReDim TipoTRF1(RowCount)
        ReDim TensionNominalNodos(size)
        ReDim TensionNominalLines(RowCount - 2)
        ReDim PhaseNodo_1(size)
        ReDim Order_(RowCount)
        ReDim CodeLine_temp(RowCount)

        ReDim desc_temp(RowCount)
        ReDim iLoadTrf_temp(RowCount)
        ReDim Strat_temp(RowCount)
        ReDim classif_temp(RowCount)
        ReDim NrCust_temp(RowCount)
        ReDim chr_name_temp(RowCount)
        ReDim sernum_temp(RowCount)


        j = 0
        i = 0
        k = 0


        For i = 0 To RowCount - 1
            If DataGridView1.Item(14, j).Value() = -1 Then
                x(i) = DataGridView1.Item(6, j).Value() 'nodos inicial
                NodoI(i) = DataGridView1.Item(6, j).Value() ' nodos iniciales nuevamente para el vector nodoi
                NodoF(i) = DataGridView1.Item(5, j).Value() ' nodos finales nuevamente para el vector nodof
                coordX1(i) = DataGridView1.Item(9, j).Value() 'coordenadas x iniciales
                coordY1(i) = DataGridView1.Item(10, j).Value() 'coordenadas y iniciales 
                LAT1_temp(i) = DataGridView1.Item(24, j).Value()
                LON1_temp(i) = DataGridView1.Item(25, j).Value()

            Else
                x(i) = DataGridView1.Item(5, j).Value() 'nodos inicial
                NodoI(i) = DataGridView1.Item(5, j).Value() ' nodos iniciales nuevamente para el vector nodoi
                NodoF(i) = DataGridView1.Item(6, j).Value() ' nodos finales nuevamente para el vector nodof
                coordX1(i) = DataGridView1.Item(7, j).Value() 'coordenadas x iniciales
                coordY1(i) = DataGridView1.Item(8, j).Value() 'coordenadas y iniciales 
                LAT1_temp(i) = DataGridView1.Item(22, j).Value()
                LON1_temp(i) = DataGridView1.Item(23, j).Value()
            End If
            PhaseNodo_1(i) = DataGridView1.Item(12, j).Value() 'fase de nodos inicial
            TensionNominalNodos(i) = DataGridView1.Item(2, j).Value()
            CodeLine_temp(i) = DataGridView1.Item(4, j).Value()
            j = j + 1
            'x(i) = DataGridView1.Item(5, j).Value() 'nodos inicial
            'NodoI(i) = DataGridView1.Item(5, j).Value() ' nodos iniciales nuevamente para el vector nodoi
            'NodoF(i) = DataGridView1.Item(6, j).Value() ' nodos finales nuevamente para el vector nodof
            'coordX1(i) = DataGridView1.Item(7, j).Value() 'coordenadas x iniciales
            'coordY1(i) = DataGridView1.Item(8, j).Value() 'coordenadas y iniciales 
        Next

        For i = j - 1 To size - 2
            If DataGridView1.Item(14, k).Value() = -1 Then
                x(i) = DataGridView1.Item(5, k).Value() 'nodos finales
                coordX1(i) = DataGridView1.Item(7, k).Value() 'coordenadas x iniciales
                coordY1(i) = DataGridView1.Item(8, k).Value() 'coordenadas y iniciales 
                LAT1_temp(i) = DataGridView1.Item(22, k).Value()
                LON1_temp(i) = DataGridView1.Item(23, k).Value()
            Else
                x(i) = DataGridView1.Item(6, k).Value() 'nodos finales
                coordX1(i) = DataGridView1.Item(9, k).Value() 'coordenadas x iniciales
                coordY1(i) = DataGridView1.Item(10, k).Value() 'coordenadas y iniciales
                LAT1_temp(i) = DataGridView1.Item(24, k).Value()
                LON1_temp(i) = DataGridView1.Item(25, k).Value()
            End If
            PhaseNodo_1(i) = DataGridView1.Item(12, k).Value() 'fase de nodos final
            TensionNominalNodos(i) = DataGridView1.Item(2, k).Value()
            k = k + 1
            'x(i) = DataGridView1.Item(6, k).Value() 'nodos finales
            'coordX1(i) = DataGridView1.Item(9, k).Value() 'coordenadas x finales
            'coordY1(i) = DataGridView1.Item(10, k).Value() 'coordenadas y finales
        Next

        For i = 0 To DataGridView1.RowCount - 2
            If Not String.IsNullOrEmpty(Me.DataGridView1.Item(11, i).Value().ToString) Then
                TRF_TEMP(i) = DataGridView1.Item(11, i).Value() 'inserta trafos en el vector TRF
                NodoTRF_Temp(i) = DataGridView1.Item(6, i).Value()
                desc_temp(i) = DataGridView1.Item(21, i).Value()
                iLoadTrf_temp(i) = DataGridView1.Item(16, i).Value()
                Strat_temp(i) = DataGridView1.Item(15, i).Value()
                classif_temp(i) = DataGridView1.Item(17, i).Value()
                NrCust_temp(i) = DataGridView1.Item(18, i).Value()
                chr_name_temp(i) = DataGridView1.Item(19, i).Value()
                sernum_temp(i) = DataGridView1.Item(20, i).Value()
                LAT_TRF_TEMP(i) = DataGridView1.Item(26, i).Value()
                LON_TRF_TEMP(i) = DataGridView1.Item(27, i).Value()
            Else
                TRF_TEMP(i) = " "
                NodoTRF_Temp(i) = " "
                desc_temp(i) = " "
                iLoadTrf_temp(i) = " "
                Strat_temp(i) = " "
                classif_temp(i) = " "
                NrCust_temp(i) = " "
                chr_name_temp(i) = " "
                sernum_temp(i) = " "
                LAT_TRF_TEMP(i) = " "
                LON_TRF_TEMP(i) = " "
            End If
        Next

        For i = 0 To DataGridView1.RowCount - 2
            If Not String.IsNullOrEmpty(Me.DataGridView1.Item(13, i).Value().ToString) Then
                TipoTRF1(i) = DataGridView1.Item(13, i).Value()
            Else
                TipoTRF1(i) = " "
            End If
        Next

        'Quitar Duplicados del vector x, lo cual genera espacios en blanco en el vector por los eliminados

        j = 0
        i = 0

        For i = 0 To size - 2
            If x(i) <> " " Then
                For j = i + 1 To size - 2
                    If x(i) = x(j) Then
                        x(j) = " "
                        TensionNominalNodos(j) = " "
                        PhaseNodo_1(j) = " "
                    End If
                Next
            End If
        Next
        'contar cuantos nodos quedan luego de la validacion de duplos
        contador = 0
        For i = 0 To size - 2
            If x(i) <> " " Then
                contador = contador + 1
            End If
        Next


        ' nuevo vector con los nodos existentes sin espacios vacios en el nuevo vector y()
        j = 0
        i = 0

        ReDim y(contador)
        ReDim PhaseNodo_2(contador)
        ReDim phtech(contador)
        ReDim TempX(contador)
        ReDim TempY(contador)
        ReDim LAT1(contador)
        ReDim LON1(contador)
        ReDim TensionNominalNodos1(contador)
        For i = 0 To size - 2
            If x(i) <> " " Then
                y(j) = x(i)
                PhaseNodo_2(j) = PhaseNodo_1(i)
                TempX(j) = coordX1(i)
                TempY(j) = coordY1(i)
                LAT1(j) = LAT1_temp(i)
                LON1(j) = LON1_temp(i)
                TensionNominalNodos1(j) = TensionNominalNodos(i)
                j = j + 1
            End If
        Next
        'volver a rectificar Phasenodo()-------------------------------------------PROBAR
        For i = 0 To y.Count - 1
            For j = 0 To NodoF.Count - 1
                If y(i) = NodoF(j) Then
                    PhaseNodo_2(i) = DataGridView1.Item(12, j).Value()
                    Exit For
                End If
            Next
        Next

        'eliminar vacios o ceros en el vector coordx1
        i = 0
        j = 0
        k = 0

        For i = 0 To size - 3
            If coordX1(i) <> 0 Then
                coordX(i) = coordX1(i)
            End If
        Next

        'eliminar vacios o ceros en el vector coordy1
        i = 0
        j = 0
        k = 0
        For i = 0 To size - 3
            If coordY1(i) <> 0 Then
                coordY(i) = coordY1(i)
            End If
        Next


        'valores para realizar cambio de coordenadas SPARD a DIGSILENT
        Dim Xmax As Double 'valor de coordenada maxmima en x
        Dim Xmin As Double 'valor de coordenada minima en x
        Dim Ymax As Double 'valor de coordenada maxmima en y
        Dim Ymin As Double 'valor de coordenada minima en y
        Dim Ex As Double   'coeficiente de regulacion en coordenadas en x
        Dim Ey As Double   'coeficiente de regulacion en coordenadas en y
        Dim Exy As Double  ' coeficiente que contiene el mayor entre Ex y Ey
        Dim Xdig As Double 'punto extremo de coordenada maxima en digrama digsilent en x, para este caso 835 en tamaño A0
        Dim Ydig As Double 'punto extremo de coordenada maxima en digrama digsilent en y, para este caso 1185 en tamaño A0

        'Xdig = 6400
        'Ydig = 8400

        'XMAX_SPARD=1218639
        'XMIN_SPARD=1006236
        'YMAX_SPARD=1501263
        'YMIN_SPARD=1209486

        If CheckBox1.Checked = False Then

            Dim Xmax_T As Double = 1218688.3
            Dim Xmin_T As Double = 1006236
            Dim Ymax_T As Double = 1501263.5
            Dim Ymin_T As Double = 1209486

            Xmax = ((99000 - 100) / (Xmax_T - Xmin_T)) * (coordX.Max - Xmin_T) + 100
            Xmin = ((99000 - 100) / (Xmax_T - Xmin_T)) * (coordX.Min - Xmin_T) + 100
            Ymax = ((99000 - 100) / (Ymax_T - Ymin_T)) * (coordY.Max - Ymin_T) + 100
            Ymin = ((99000 - 100) / (Ymax_T - Ymin_T)) * (coordY.Min - Ymin_T) + 100

            'Xmax = coordX.Max
            'Xmin = coordX.Min
            'Ymax = coordY.Max
            'Ymin = coordY.Min

            Ex = (Xmax - Xmin)
            Ey = (Ymax - Ymin)

            If Ex > Ey Then
                Exy = Ex
            Else
                Exy = Ey
            End If
            TextBox1.Text = Xmax
            TextBox2.Text = Xmin
            TextBox3.Text = Ymax
            TextBox4.Text = Ymin
            TextBox5.Text = Ex
            TextBox6.Text = Ey

            'Console.WriteLine(Xdig & "XDIG")
            'Console.WriteLine(Ydig & "YDIG")

        Else
            Xmax = TextBox1.Text
            Xmin = TextBox2.Text
            Ymax = TextBox3.Text
            Ymin = TextBox4.Text

            Ex = (Xmax - Xmin) / (0.8 * Xdig)
            Ey = (Ymax - Ymin) / (0.8 * Ydig)


            If Ex > Ey Then
                Exy = Ex
            Else
                Exy = Ey
            End If
            TextBox5.Text = Ex
            TextBox6.Text = Ey
        End If


        'eliminar los trafos duplicados del vector TRF_TEMP()
        j = 0
        i = 0

        For i = 0 To RowCount - 2
            If TRF_TEMP(i) <> " " Then
                For j = i + 1 To RowCount - 2
                    If TRF_TEMP(i) = TRF_TEMP(j) Then
                        TRF_TEMP(j) = " "
                        NodoTRF_Temp(j) = " "
                        TipoTRF1(j) = " "
                    End If
                Next
            End If
        Next


        'contar cuantos trafos existen en el circuito y pasar los trafos existentes al vector nuevo TRF() con el tamaño de contador 
        contador = 0
        For i = 0 To RowCount - 2
            If TRF_TEMP(i) <> " " And TRF_TEMP(i) <> "0" Then
                contador = contador + 1
            End If
        Next

        ReDim TRF(contador - 1)
        ReDim NodoTRF(contador - 1)
        ReDim TipoTRF2(contador - 1)
        ReDim TipoTRF3(contador - 1)
        ReDim it2p1_TRF(contador - 1)
        ReDim it2p2_TRF(contador - 1)

        ReDim desc(contador - 1)
        ReDim iLoadTrf(contador - 1)
        ReDim Strat(contador - 1)
        ReDim classif(contador - 1)
        ReDim NrCust(contador - 1)
        ReDim chr_name(contador - 1)
        ReDim sernum(contador - 1)

        ReDim LAT_TRF(contador - 1)
        ReDim LON_TRF(contador - 1)

        j = 0
        For i = 0 To RowCount - 2
            If TRF_TEMP(i) <> " " And TRF_TEMP(i) <> "0" Then
                TRF(j) = TRF_TEMP(i)
                NodoTRF(j) = NodoTRF_Temp(i)
                TipoTRF2(j) = TipoTRF1(i)

                desc(j) = desc_temp(i)
                iLoadTrf(j) = iLoadTrf_temp(i)
                Strat(j) = Strat_temp(i)
                classif(j) = classif_temp(i)
                NrCust(j) = NrCust_temp(i)
                chr_name(j) = chr_name_temp(i)
                sernum(j) = sernum_temp(i)

                LAT_TRF(j) = LAT_TRF_TEMP(i)
                LON_TRF(j) = LON_TRF_TEMP(i)


                j = j + 1
            End If
        Next
        'llenar el vector Line_Temp() con los vectores NodoI NodoF para obtener los strings de lines TYPLIN_NODOI-NODOF
        For i = 0 To RowCount - 2
            If NodoI(i) <> NodoF(i) Then
                Line_Temp(i) = NodoI(i) & "-" & NodoF(i)
                Sline_Temp(i) = DataGridView1.Item(3, i).Value()
                longitud_temp(i) = DataGridView1.Item(1, i).Value()
                PhaseLine_Temp(i) = DataGridView1.Item(12, i).Value()
                TensionNominalLines(i) = DataGridView1.Item(2, i).Value()
            Else
                Line_Temp(i) = " "
                Sline_Temp(i) = " "
                longitud_temp(i) = 0.0
            End If
        Next
        'eliminar duplicados de lineas 
        j = 0
        i = 0

        For i = 0 To RowCount - 2
            If Line_Temp(i) <> " " Then
                For j = i + 1 To RowCount - 2
                    If Line_Temp(i) = Line_Temp(j) Then
                        Line_Temp(j) = " "
                        Sline_Temp(j) = " "
                        longitud_temp(j) = 0.0
                        NodoI(j) = " "
                        NodoF(j) = " "
                        CodeLine_temp(i) = " "
                    End If
                Next
            End If
        Next
        'eliminamos espacios vacios y llenamos vector Line() que seria el vector filtrado y validado de lineas. 
        contador = 0
        For i = 0 To RowCount - 2
            If Line_Temp(i) <> " " Then
                contador = contador + 1
            End If
        Next
        ReDim line(contador)
        ReDim Sline(contador)
        ReDim SlineAmp(contador)
        ReDim longitud(contador)
        ReDim nodo1(contador)
        ReDim nodo2(contador)
        ReDim PhaseLine(contador)
        ReDim it2p1_Line(contador)
        ReDim it2p2_Line(contador)
        ReDim nlnph(contador)
        ReDim nneutral(contador)
        ReDim TensionNominalLines1(contador)
        ReDim CodeLine(contador)



        j = 0
        For i = 0 To RowCount - 2
            If Line_Temp(i) <> " " Then
                line(j) = Line_Temp(i)
                Sline(j) = Sline_Temp(i)
                longitud(j) = longitud_temp(i)
                nodo1(j) = NodoI(i)
                nodo2(j) = NodoF(i)
                PhaseLine(j) = PhaseLine_Temp(i)
                TensionNominalLines1(j) = TensionNominalLines(i)
                CodeLine(j) = CodeLine_temp(i)
                j = j + 1
            End If
        Next



        'vectoes que contienen las posicines de los nodos en la pestaña ElmNet
        k = 0
        For i = 0 To nodo1.Count - 1
            For j = 0 To y.Count - 1
                If nodo1(i) = y(j) Then
                    PositionsNodo1(k) = j + 4
                    k = k + 1
                End If
            Next
        Next
        k = 0
        For i = 0 To nodo2.Count - 1
            For j = 0 To y.Count - 1
                If nodo2(i) = y(j) Then
                    PositionsNodo2(k) = j + 4
                    k = k + 1
                End If
            Next
        Next


        'llenar vector que contiene las fases de la linea nlnph() y asignar el control de flujo de las lineas segun vectores it2p1 e it2p2

        For i = 0 To PhaseLine.Count - 1
            If PhaseLine(i) = "1" Then
                nlnph(i) = "1"
                nneutral(i) = "0"
                it2p1_Line(i) = "0"
                it2p2_Line(i) = ""
            Else
                If PhaseLine(i) = "2" Then
                    nlnph(i) = "1"
                    nneutral(i) = "0"
                    it2p1_Line(i) = "1"
                    it2p2_Line(i) = ""
                Else
                    If PhaseLine(i) = "3" Then
                        nlnph(i) = "2"
                        nneutral(i) = "0"
                        it2p1_Line(i) = "0"
                        it2p2_Line(i) = "1"
                    Else
                        If PhaseLine(i) = "4" Then
                            nlnph(i) = "1"
                            nneutral(i) = "0"
                            it2p1_Line(i) = "2"
                            it2p2_Line(i) = ""
                        Else
                            If PhaseLine(i) = "5" Then
                                nlnph(i) = "2"
                                nneutral(i) = "0"
                                it2p1_Line(i) = "0"
                                it2p2_Line(i) = "2"
                            Else
                                If PhaseLine(i) = "6" Then
                                    nlnph(i) = "2"
                                    nneutral(i) = "0"
                                    it2p1_Line(i) = "1"
                                    it2p2_Line(i) = "2"
                                Else
                                    If PhaseLine(i) = "7" Then
                                        nlnph(i) = "3"
                                        nneutral(i) = "0"
                                        it2p1_Line(i) = ""
                                        it2p2_Line(i) = ""
                                    Else
                                        If PhaseLine(i) = "8" Then
                                            nlnph(i) = "1"
                                            nneutral(i) = "0"
                                            it2p1_Line(i) = "0"
                                            it2p2_Line(i) = ""
                                        Else
                                            If PhaseLine(i) = "9" Then
                                                nlnph(i) = "1"
                                                nneutral(i) = "0"
                                                it2p1_Line(i) = "0"
                                                it2p2_Line(i) = ""
                                            Else
                                                If PhaseLine(i) = "10" Then
                                                    nlnph(i) = "1"
                                                    nneutral(i) = "0"
                                                    it2p1_Line(i) = "1"
                                                    it2p2_Line(i) = ""
                                                Else
                                                    If PhaseLine(i) = "11" Then
                                                        nlnph(i) = "1"
                                                        nneutral(i) = "0"
                                                        it2p1_Line(i) = "0"
                                                        it2p2_Line(i) = "1"
                                                    Else
                                                        If PhaseLine(i) = "12" Then
                                                            nlnph(i) = "1"
                                                            nneutral(i) = "0"
                                                            it2p1_Line(i) = "2"
                                                            it2p2_Line(i) = ""
                                                        Else
                                                            If PhaseLine(i) = "13" Then
                                                                nlnph(i) = "2"
                                                                nneutral(i) = "0"
                                                                it2p1_Line(i) = "0"
                                                                it2p2_Line(i) = "2"
                                                            Else
                                                                If PhaseLine(i) = "14" Then
                                                                    nlnph(i) = "2"
                                                                    nneutral(i) = "0"
                                                                    it2p1_Line(i) = "1"
                                                                    it2p2_Line(i) = "2"
                                                                Else
                                                                    If PhaseLine(i) = "15" Then
                                                                        nlnph(i) = "3"
                                                                        nneutral(i) = "0"
                                                                        it2p1_Line(i) = ""
                                                                        it2p2_Line(i) = ""
                                                                    Else
                                                                        If PhaseLine(i) = "16" Then
                                                                            nlnph(i) = "1"
                                                                            nneutral(i) = "0"
                                                                            it2p1_Line(i) = "0"
                                                                            it2p2_Line(i) = ""
                                                                        Else
                                                                            If PhaseLine(i) = "17" Then
                                                                                nlnph(i) = "1"
                                                                                nneutral(i) = "0"
                                                                                it2p1_Line(i) = "0"
                                                                                it2p2_Line(i) = ""
                                                                            Else
                                                                                If PhaseLine(i) = "18" Then
                                                                                    nlnph(i) = "1"
                                                                                    nneutral(i) = "0"
                                                                                    it2p1_Line(i) = "1"
                                                                                    it2p2_Line(i) = ""
                                                                                Else
                                                                                    If PhaseLine(i) = "19" Then
                                                                                        nlnph(i) = "2"
                                                                                        nneutral(i) = "0"
                                                                                        it2p1_Line(i) = "0"
                                                                                        it2p2_Line(i) = "1"
                                                                                    Else
                                                                                        If PhaseLine(i) = "20" Then
                                                                                            nlnph(i) = "1"
                                                                                            nneutral(i) = "0"
                                                                                            it2p1_Line(i) = "2"
                                                                                            it2p2_Line(i) = ""
                                                                                        Else
                                                                                            If PhaseLine(i) = "21" Then
                                                                                                nlnph(i) = "2"
                                                                                                nneutral(i) = "0"
                                                                                                it2p1_Line(i) = "0"
                                                                                                it2p2_Line(i) = "2"
                                                                                            Else
                                                                                                If PhaseLine(i) = "22" Then
                                                                                                    nlnph(i) = "2"
                                                                                                    nneutral(i) = "0"
                                                                                                    it2p1_Line(i) = "1"
                                                                                                    it2p2_Line(i) = "2"
                                                                                                Else
                                                                                                    If PhaseLine(i) = "23" Then
                                                                                                        nlnph(i) = "3"
                                                                                                        nneutral(i) = "0"
                                                                                                        it2p1_Line(i) = ""
                                                                                                        it2p2_Line(i) = ""
                                                                                                    Else
                                                                                                        If PhaseLine(i) = "24" Then
                                                                                                            nlnph(i) = "1"
                                                                                                            nneutral(i) = "0"
                                                                                                            it2p1_Line(i) = "0"
                                                                                                            it2p2_Line(i) = ""
                                                                                                        Else
                                                                                                            If PhaseLine(i) = "25" Then
                                                                                                                nlnph(i) = "1"
                                                                                                                nneutral(i) = "0"
                                                                                                                it2p1_Line(i) = "0"
                                                                                                                it2p2_Line(i) = ""
                                                                                                            Else
                                                                                                                If PhaseLine(i) = "26" Then
                                                                                                                    nlnph(i) = "1"
                                                                                                                    nneutral(i) = "0"
                                                                                                                    it2p1_Line(i) = "1"
                                                                                                                    it2p2_Line(i) = ""
                                                                                                                Else
                                                                                                                    If PhaseLine(i) = "27" Then
                                                                                                                        nlnph(i) = "2"
                                                                                                                        nneutral(i) = "0"
                                                                                                                        it2p1_Line(i) = "0"
                                                                                                                        it2p2_Line(i) = "1"
                                                                                                                    Else
                                                                                                                        If PhaseLine(i) = "28" Then
                                                                                                                            nlnph(i) = "1"
                                                                                                                            nneutral(i) = "0"
                                                                                                                            it2p1_Line(i) = "2"
                                                                                                                            it2p2_Line(i) = ""
                                                                                                                        Else
                                                                                                                            If PhaseLine(i) = "29" Then
                                                                                                                                nlnph(i) = "2"
                                                                                                                                nneutral(i) = "0"
                                                                                                                                it2p1_Line(i) = "0"
                                                                                                                                it2p2_Line(i) = "2"
                                                                                                                            Else
                                                                                                                                If PhaseLine(i) = "30" Then
                                                                                                                                    nlnph(i) = "2"
                                                                                                                                    nneutral(i) = "0"
                                                                                                                                    it2p1_Line(i) = "1"
                                                                                                                                    it2p2_Line(i) = "2"
                                                                                                                                Else
                                                                                                                                    If PhaseLine(i) = "31" Then
                                                                                                                                        nlnph(i) = "3"
                                                                                                                                        nneutral(i) = "0"
                                                                                                                                        it2p1_Line(i) = ""
                                                                                                                                        it2p2_Line(i) = ""
                                                                                                                                    End If
                                                                                                                                End If
                                                                                                                            End If
                                                                                                                        End If
                                                                                                                    End If
                                                                                                                End If
                                                                                                            End If
                                                                                                        End If
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next


        'llenar vector que contiene las fases del trafo en vector de tipos de trafos en la biblioteca digsilent:  tipoTRF2 a TipoTRF3


        If TextTension.Text = "13,8" Then
            TextTension.Text = "13.8"
        Else
            If TextTension.Text = "34,5" Then
                TextTension.Text = "34.5"
            End If
        End If

        If TextTension.Text = "13.8" Then
            For i = 0 To TipoTRF2.Count - 1
                If TipoTRF2(i) = "1" Then
                    TipoTRF3(i) = "##LOD_13.8kV-1F"
                    it2p1_TRF(i) = "0"
                    it2p2_TRF(i) = ""
                Else
                    If TipoTRF2(i) = "2" Then
                        TipoTRF3(i) = "##LOD_13.8kV-1F"
                        it2p1_TRF(i) = "1"
                        it2p2_TRF(i) = ""
                    Else
                        If TipoTRF2(i) = "3" Then
                            TipoTRF3(i) = "##LOD_13.8kV-2F"
                            it2p1_TRF(i) = "0"
                            it2p2_TRF(i) = "1"
                        Else
                            If TipoTRF2(i) = "4" Then
                                TipoTRF3(i) = "##LOD_13.8kV-1F"
                                it2p1_TRF(i) = "2"
                                it2p2_TRF(i) = ""
                            Else
                                If TipoTRF2(i) = "5" Then
                                    TipoTRF3(i) = "##LOD_13.8kV-2F"
                                    it2p1_TRF(i) = "0"
                                    it2p2_TRF(i) = "2"
                                Else
                                    If TipoTRF2(i) = "6" Then
                                        TipoTRF3(i) = "##LOD_13.8kV-2F"
                                        it2p1_TRF(i) = "1"
                                        it2p2_TRF(i) = "2"
                                    Else
                                        If TipoTRF2(i) = "7" Then
                                            TipoTRF3(i) = "##LOD_13.8kV-3F"
                                            it2p1_TRF(i) = ""
                                            it2p2_TRF(i) = ""
                                        Else
                                            If TipoTRF2(i) = "8" Then
                                                TipoTRF3(i) = "##LOD_13.8kV-1F"
                                                it2p1_TRF(i) = "0"
                                                it2p2_TRF(i) = ""
                                            Else
                                                If TipoTRF2(i) = "9" Then
                                                    TipoTRF3(i) = "##LOD_13.8kV-1F"
                                                    it2p1_TRF(i) = "0"
                                                    it2p2_TRF(i) = ""
                                                Else
                                                    If TipoTRF2(i) = "10" Then
                                                        TipoTRF3(i) = "##LOD_13.8kV-1F"
                                                        it2p1_TRF(i) = "1"
                                                        it2p2_TRF(i) = ""
                                                    Else
                                                        If TipoTRF2(i) = "11" Then
                                                            TipoTRF3(i) = "##LOD_13.8kV-1F"
                                                            it2p1_TRF(i) = "0"
                                                            it2p2_TRF(i) = "1"
                                                        Else
                                                            If TipoTRF2(i) = "12" Then
                                                                TipoTRF3(i) = "##LOD_13.8kV-1F"
                                                                it2p1_TRF(i) = "2"
                                                                it2p2_TRF(i) = ""
                                                            Else
                                                                If TipoTRF2(i) = "13" Then
                                                                    TipoTRF3(i) = "##LOD_13.8kV-2F"
                                                                    it2p1_TRF(i) = "0"
                                                                    it2p2_TRF(i) = "2"
                                                                Else
                                                                    If TipoTRF2(i) = "14" Then
                                                                        TipoTRF3(i) = "##LOD_13.8kV-2F"
                                                                        it2p1_TRF(i) = "1"
                                                                        it2p2_TRF(i) = "2"
                                                                    Else
                                                                        If TipoTRF2(i) = "15" Then
                                                                            TipoTRF3(i) = "##LOD_13.8kV-3F"
                                                                            it2p1_TRF(i) = ""
                                                                            it2p2_TRF(i) = ""
                                                                        Else
                                                                            If TipoTRF2(i) = "16" Then
                                                                                TipoTRF3(i) = "##LOD_13.8kV-1F"
                                                                                it2p1_TRF(i) = "0"
                                                                                it2p2_TRF(i) = ""
                                                                            Else
                                                                                If TipoTRF2(i) = "17" Then
                                                                                    TipoTRF3(i) = "##LOD_13.8kV-1F"
                                                                                    it2p1_TRF(i) = "0"
                                                                                    it2p2_TRF(i) = ""
                                                                                Else
                                                                                    If TipoTRF2(i) = "18" Then
                                                                                        TipoTRF3(i) = "##LOD_13.8kV-1F"
                                                                                        it2p1_TRF(i) = "1"
                                                                                        it2p2_TRF(i) = ""
                                                                                    Else
                                                                                        If TipoTRF2(i) = "19" Then
                                                                                            TipoTRF3(i) = "##LOD_13.8kV-2F"
                                                                                            it2p1_TRF(i) = "0"
                                                                                            it2p2_TRF(i) = "1"
                                                                                        Else
                                                                                            If TipoTRF2(i) = "20" Then
                                                                                                TipoTRF3(i) = "##LOD_13.8kV-1F"
                                                                                                it2p1_TRF(i) = "2"
                                                                                                it2p2_TRF(i) = ""
                                                                                            Else
                                                                                                If TipoTRF2(i) = "21" Then
                                                                                                    TipoTRF3(i) = "##LOD_13.8kV-2F"
                                                                                                    it2p1_TRF(i) = "0"
                                                                                                    it2p2_TRF(i) = "2"
                                                                                                Else
                                                                                                    If TipoTRF2(i) = "22" Then
                                                                                                        TipoTRF3(i) = "##LOD_13.8kV-2F"
                                                                                                        it2p1_TRF(i) = "1"
                                                                                                        it2p2_TRF(i) = "2"
                                                                                                    Else
                                                                                                        If TipoTRF2(i) = "23" Then
                                                                                                            TipoTRF3(i) = "##LOD_13.8kV-3F"
                                                                                                            it2p1_TRF(i) = ""
                                                                                                            it2p2_TRF(i) = ""
                                                                                                        Else
                                                                                                            If TipoTRF2(i) = "24" Then
                                                                                                                TipoTRF3(i) = "##LOD_13.8kV-1F"
                                                                                                                it2p1_TRF(i) = "0"
                                                                                                                it2p2_TRF(i) = ""
                                                                                                            Else
                                                                                                                If TipoTRF2(i) = "25" Then
                                                                                                                    TipoTRF3(i) = "##LOD_13.8kV-1F"
                                                                                                                    it2p1_TRF(i) = "0"
                                                                                                                    it2p2_TRF(i) = ""
                                                                                                                Else
                                                                                                                    If TipoTRF2(i) = "26" Then
                                                                                                                        TipoTRF3(i) = "##LOD_13.8kV-1F"
                                                                                                                        it2p1_TRF(i) = "1"
                                                                                                                        it2p2_TRF(i) = ""
                                                                                                                    Else
                                                                                                                        If TipoTRF2(i) = "27" Then
                                                                                                                            TipoTRF3(i) = "##LOD_13.8kV-2F"
                                                                                                                            it2p1_TRF(i) = "0"
                                                                                                                            it2p2_TRF(i) = "1"
                                                                                                                        Else
                                                                                                                            If TipoTRF2(i) = "28" Then
                                                                                                                                TipoTRF3(i) = "##LOD_13.8kV-1F"
                                                                                                                                it2p1_TRF(i) = "2"
                                                                                                                                it2p2_TRF(i) = "0"
                                                                                                                            Else
                                                                                                                                If TipoTRF2(i) = "29" Then
                                                                                                                                    TipoTRF3(i) = "##LOD_13.8kV-2F"
                                                                                                                                    it2p1_TRF(i) = "0"
                                                                                                                                    it2p2_TRF(i) = "2"
                                                                                                                                Else
                                                                                                                                    If TipoTRF2(i) = "30" Then
                                                                                                                                        TipoTRF3(i) = "##LOD_13.8kV-2F"
                                                                                                                                        it2p1_TRF(i) = "1"
                                                                                                                                        it2p2_TRF(i) = "2"
                                                                                                                                    Else
                                                                                                                                        If TipoTRF2(i) = "31" Then
                                                                                                                                            TipoTRF3(i) = "##LOD_13.8kV-3F"
                                                                                                                                            it2p1_TRF(i) = ""
                                                                                                                                            it2p2_TRF(i) = ""
                                                                                                                                        End If
                                                                                                                                    End If
                                                                                                                                End If
                                                                                                                            End If
                                                                                                                        End If
                                                                                                                    End If
                                                                                                                End If
                                                                                                            End If
                                                                                                        End If
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If

        If TextTension.Text = "34.5" Then
            For i = 0 To TipoTRF2.Count - 1
                TipoTRF3(i) = "##LOD_34.5kV-3F"
            Next
        End If

        If TextTension.Text = "115" Then
            For i = 0 To TipoTRF2.Count - 1
                TipoTRF3(i) = "##LOD_115kV-3F"
            Next
        End If




        'llenar vector phtech (tecnologia del nodo) al convertir las fases de la linea de cada nodo en cada linea

        For i = 0 To PhaseNodo_2.Count - 1
            If PhaseNodo_2(i) = "1" Then
                phtech(i) = "6"
            Else
                If PhaseNodo_2(i) = "2" Then
                    phtech(i) = "6"
                Else
                    If PhaseNodo_2(i) = "3" Then
                        phtech(i) = "2"
                    Else
                        If PhaseNodo_2(i) = "4" Then
                            phtech(i) = "6"
                        Else
                            If PhaseNodo_2(i) = "5" Then
                                phtech(i) = "2"
                            Else
                                If PhaseNodo_2(i) = "6" Then
                                    phtech(i) = "2"
                                Else
                                    If PhaseNodo_2(i) = "7" Then
                                        phtech(i) = "0"
                                    Else
                                        If PhaseNodo_2(i) = "8" Then
                                            phtech(i) = "6"
                                        Else
                                            If PhaseNodo_2(i) = "9" Then
                                                phtech(i) = "6"
                                            Else
                                                If PhaseNodo_2(i) = "10" Then
                                                    phtech(i) = "6"
                                                Else
                                                    If PhaseNodo_2(i) = "11" Then
                                                        phtech(i) = "6"
                                                    Else
                                                        If PhaseNodo_2(i) = "12" Then
                                                            phtech(i) = "6"
                                                        Else
                                                            If PhaseNodo_2(i) = "13" Then
                                                                phtech(i) = "2"
                                                            Else
                                                                If PhaseNodo_2(i) = "14" Then
                                                                    phtech(i) = "2"
                                                                Else
                                                                    If PhaseNodo_2(i) = "15" Then
                                                                        phtech(i) = "0"
                                                                    Else
                                                                        If PhaseNodo_2(i) = "16" Then
                                                                            phtech(i) = "6"
                                                                        Else
                                                                            If PhaseNodo_2(i) = "17" Then
                                                                                phtech(i) = "6"
                                                                            Else
                                                                                If PhaseNodo_2(i) = "18" Then
                                                                                    phtech(i) = "6"
                                                                                Else
                                                                                    If PhaseNodo_2(i) = "19" Then
                                                                                        phtech(i) = "2"
                                                                                    Else
                                                                                        If PhaseNodo_2(i) = "20" Then
                                                                                            phtech(i) = "6"
                                                                                        Else
                                                                                            If PhaseNodo_2(i) = "21" Then
                                                                                                phtech(i) = "2"
                                                                                            Else
                                                                                                If PhaseNodo_2(i) = "22" Then
                                                                                                    phtech(i) = "2"
                                                                                                Else
                                                                                                    If PhaseNodo_2(i) = "23" Then
                                                                                                        phtech(i) = "0"
                                                                                                    Else
                                                                                                        If PhaseNodo_2(i) = "24" Then
                                                                                                            phtech(i) = "6"
                                                                                                        Else
                                                                                                            If PhaseNodo_2(i) = "25" Then
                                                                                                                phtech(i) = "6"
                                                                                                            Else
                                                                                                                If PhaseNodo_2(i) = "26" Then
                                                                                                                    phtech(i) = "6"
                                                                                                                Else
                                                                                                                    If PhaseNodo_2(i) = "27" Then
                                                                                                                        phtech(i) = "2"
                                                                                                                    Else
                                                                                                                        If PhaseNodo_2(i) = "28" Then
                                                                                                                            phtech(i) = "6"
                                                                                                                        Else
                                                                                                                            If PhaseNodo_2(i) = "29" Then
                                                                                                                                phtech(i) = "2"
                                                                                                                            Else
                                                                                                                                If PhaseNodo_2(i) = "30" Then
                                                                                                                                    phtech(i) = "2"
                                                                                                                                Else
                                                                                                                                    If PhaseNodo_2(i) = "31" Then
                                                                                                                                        phtech(i) = "0"
                                                                                                                                    End If
                                                                                                                                End If
                                                                                                                            End If
                                                                                                                        End If
                                                                                                                    End If
                                                                                                                End If
                                                                                                            End If
                                                                                                        End If
                                                                                                    End If
                                                                                                End If
                                                                                            End If
                                                                                        End If
                                                                                    End If
                                                                                End If
                                                                            End If
                                                                        End If
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next






        'obtener las posiciones de los noodos con trafos conectados y almacenarlos en el vector positionsntrf
        k = 0
        For i = 0 To NodoTRF.Count - 1
            For j = 0 To y.Count - 1
                If NodoTRF(i) = y(j) Then
                    positionsNTRF(k) = j + 4
                    k = k + 1
                End If
            Next
        Next


        ' llenar los campos de la hoja ElmTerm
        var1 = 1
        var2 = y.Count / 20
        m = 2
        consecutivo1 = 3
        For i = 0 To y.Count - 3
            Application.DoEvents()
            consecutivo1 = consecutivo1 + 1
            libro.Sheets("ElmTerm").cells(m, 1).NumberFormat = "0.00"
            libro.Sheets("ElmTerm").cells(m, 1) = consecutivo1
            libro.Sheets("ElmTerm").cells(m, 2) = "Node_" & y(i)
            libro.Sheets("ElmTerm").cells(m, 3) = 2
            libro.Sheets("ElmTerm").cells(m, 5) = 1
            libro.Sheets("ElmTerm").cells(m, 6) = phtech(i)
            libro.Sheets("ElmTerm").cells(m, 7) = TensionNominalNodos1(i)
            libro.Sheets("ElmTerm").cells(m, 8) = y(i) 'espacion unidad constructiva'
            libro.Sheets("ElmTerm").cells(m, 9) = 0
            libro.Sheets("ElmTerm").cells(m, 10) = "_" & y(i)
            libro.Sheets("ElmTerm").cells(m, 11) = LAT1(i)
            libro.Sheets("ElmTerm").cells(m, 12) = LON1(i)
            LabelP1.Text = "ID(a:40)  " & consecutivo1
            LabelP2.Text = "loc_name(a:40)  " & y(i)
            LabelP3.Text = "fold_id(p)  " & 2
            LabelP4.Text = "typ_id(p)  " & ""
            LabelP5.Text = "iUsage(i)  " & 1
            LabelP6.Text = "phtech  " & phtech(i)
            LabelP7.Text = "uknom(r)  " & TensionNominalNodos1(i)
            LabelP8.Text = "chr_name(a:20)  " & y(i)
            LabelP9.Text = "outserv(i)  " & 0
            LabelP10.Text = "for_name  " & y(i)
            LabelP11.Text = "GPSlat(r) " & LAT1(i)
            LabelP12.Text = "GPSlon(r) " & LON1(i)
            LabelP13.Text = " "

            m = m + 1
            If i >= var1 * var2 And ProgressBar.Value < 26 Then
                ProgressBar.Value = ProgressBar.Value + 1
                LabelPorcentaje.Text = ProgressBar.Value & "%"
                var1 = var1 + 1
            End If
        Next
        ProgressBar.Value = 26
        LabelPorcentaje.Text = ProgressBar.Value & "%"


        'llenar los campos de la hoja IntGrf 
        var1 = 1
        var2 = y.Count / 20
        m = 2
        For i = 0 To y.Count - 3
            Application.DoEvents()
            consecutivo1 = consecutivo1 + 1
            libro.Sheets("IntGrf").cells(m, 2) = "Gra_" & y(i) & ""
            libro.Sheets("IntGrf").cells(m, 1).NumberFormat = "0.00"
            libro.Sheets("IntGrf").cells(m, 1) = consecutivo1
            libro.Sheets("IntGrf").cells(m, 3).NumberFormat = "0.00"
            libro.Sheets("IntGrf").cells(m, 3) = 3.0
            libro.Sheets("IntGrf").cells(m, 4).NumberFormat = "0.00"
            libro.Sheets("IntGrf").cells(m, 4) = 1.0
            libro.Sheets("IntGrf").cells(m, 5).NumberFormat = "0.00"
            libro.Sheets("IntGrf").cells(m, 5) = 1.0
            libro.Sheets("IntGrf").cells(m, 6).NumberFormat = "0.00"
            libro.Sheets("IntGrf").cells(m, 6) = 1.0
            libro.Sheets("IntGrf").cells(m, 7).NumberFormat = "0.000"
            libro.Sheets("IntGrf").cells(m, 7) = 100 + ((((99000 / 212452.3) * (TempX(i) - 1006236) + 100) - Xmin))
            libro.Sheets("IntGrf").cells(m, 8).NumberFormat = "0.000"
            libro.Sheets("IntGrf").cells(m, 8) = 100 + ((((99000 / 291777.5) * (TempY(i) - 1209486) + 100) - Ymin))
            libro.Sheets("IntGrf").cells(m, 9) = "PointTerm"
            libro.Sheets("IntGrf").cells(m, 10).NumberFormat = "0.00"
            libro.Sheets("IntGrf").cells(m, 10) = i + 4
            libro.Sheets("IntGrf").cells(m, 11).NumberFormat = "0.00"
            libro.Sheets("IntGrf").cells(m, 11) = 0.0
            libro.Sheets("IntGrf").cells(m, 12).NumberFormat = "0.00"
            libro.Sheets("IntGrf").cells(m, 12) = 1.0
            libro.Sheets("IntGrf").cells(m, 13).NumberFormat = "0.00"
            libro.Sheets("IntGrf").cells(m, 13) = 1.0
            libro.Sheets("IntGrf").cells(m, 20) = "Gra_" & y(i) & ""
            LabelP1.Text = "ID(a:40)  " & consecutivo1
            LabelP2.Text = "loc_name(a:40)  " & "Gra_" & y(i) & ""
            LabelP3.Text = "fold_id(p)  " & 3.0
            LabelP4.Text = "iCol(i)  " & 1.0
            LabelP5.Text = "iVis(i)  " & 1.0
            LabelP6.Text = "iLevel(i)  " & 1.0
            LabelP7.Text = "rCenterX(r)  " & 0.2 * Xdig + ((TempX(i) - Xmin) / Exy)
            LabelP8.Text = "rCenterY(r)  " & 0.2 * Ydig + ((TempY(i) - Ymin) / Exy)
            LabelP9.Text = "sSymNam(a:40)  " & "PointTerm"
            LabelP10.Text = "pDataObj(p)  " & i + 4
            LabelP11.Text = "iRot(i)  " & "0.00"
            LabelP12.Text = "rSizeX(r)  " & "1.00"
            LabelP13.Text = "rSizeY(r)  " & "1.00"
            m = m + 1
            If i >= var1 * var2 And ProgressBar.Value < 36 Then
                ProgressBar.Value = ProgressBar.Value + 1
                LabelPorcentaje.Text = ProgressBar.Value & "%"
                var1 = var1 + 1
            End If
        Next
        ProgressBar.Value = 36
        LabelPorcentaje.Text = ProgressBar.Value & "%"


        'llenar los campos de la hoja ElmLod
        var1 = 1
        var2 = TRF.Count / 20

        consecutivoTRF2 = consecutivo1
        m = 2
        consecutivoTRF = 4
        For i = 0 To TRF.Count - 1
            Application.DoEvents()
            consecutivo1 = consecutivo1 + 1
            consecutivoTRF = consecutivoTRF + 1
            libro.Sheets("ElmLod").cells(m, 2) = "Lod_" & TRF(i)
            libro.Sheets("ElmLod").cells(m, 1).NumberFormat = "0.00"
            libro.Sheets("ElmLod").cells(m, 1) = consecutivo1
            libro.Sheets("ElmLod").cells(m, 3) = 2
            libro.Sheets("ElmLod").cells(m, 4) = positionsNTRF(i)
            libro.Sheets("ElmLod").cells(m, 5) = TipoTRF3(i)
            libro.Sheets("ElmLod").cells(m, 6) = "PQ"
            libro.Sheets("ElmLod").cells(m, 7) = chr_name(i)
            libro.Sheets("ElmLod").cells(m, 8) = 0
            libro.Sheets("ElmLod").cells(m, 9) = 0
            libro.Sheets("ElmLod").cells(m, 10) = 1
            libro.Sheets("ElmLod").cells(m, 11) = desc(i)
            libro.Sheets("ElmLod").cells(m, 12) = 0 'iLoadTrf(i)
            libro.Sheets("ElmLod").cells(m, 13) = Strat(i)
            libro.Sheets("ElmLod").cells(m, 14) = classif(i)
            libro.Sheets("ElmLod").cells(m, 15) = NrCust(i)
            libro.Sheets("ElmLod").cells(m, 16) = sernum(i)
            libro.Sheets("ElmLod").cells(m, 17) = TRF(i)
            libro.Sheets("ElmLod").cells(m, 18) = LAT_TRF(i)
            libro.Sheets("ElmLod").cells(m, 19) = LON_TRF(i)
            libro.Sheets("ElmLod").cells(m, 20) = TRF(i)

            LabelP1.Text = "ID(a:40)  " & consecutivo1
            LabelP2.Text = "loc_name(a:40)  " & TRF(i)
            LabelP3.Text = "fold_id(p)  " & 2
            LabelP4.Text = "bus1(p)  " & positionsNTRF(i)
            LabelP5.Text = "typ_id(p)  " & TipoTRF3(i)
            LabelP6.Text = "mode_inp  " & "PQ"
            LabelP7.Text = "chr_name(a:20)  " & TRF(i)
            LabelP8.Text = "plini(r)  " & 0
            LabelP9.Text = "qlini(r)  " & 0
            LabelP10.Text = "scale0(r)  " & 1
            LabelP11.Text = "desc:0(a) " & desc(i)
            LabelP12.Text = "iLoadTrf(i) " & iLoadTrf(i)
            LabelP13.Text = "Strat(r) " & Strat(i)
            m = m + 1
            If i >= var1 * var2 And ProgressBar.Value < 46 Then
                ProgressBar.Value = ProgressBar.Value + 1
                LabelPorcentaje.Text = ProgressBar.Value & "%"
                var1 = var1 + 1
            End If
        Next
        ProgressBar.Value = 46
        LabelPorcentaje.Text = ProgressBar.Value & "%"


        'llenar los campos de la hoja TypLne
        var1 = 1
        var2 = line.Count / 20
        m = 2
        consecutivoTRF3 = consecutivo1
        consecutivo1 = consecutivo1 + consecutivoTRF - 4
        consecutivo3 = consecutivo1



        ProgressBar.Value = 56
        LabelPorcentaje.Text = ProgressBar.Value & "%"

        consecutivo4 = consecutivo1
        consecutivo5 = consecutivo1

        Application.DoEvents()
        'llenar los campos de la hoja ElmLne
        var1 = 1
        var2 = line.Count / 20
        m = 2
        For i = 0 To line.Count - 2
            Application.DoEvents()
            consecutivo1 = consecutivo1 + 1
            consecutivo3 = consecutivo3 + 1
            libro.Sheets("ElmLne").cells(m, 2) = "LIN_" & CodeLine(i)
            libro.Sheets("ElmLne").cells(m, 1) = consecutivo1
            libro.Sheets("ElmLne").cells(m, 3) = "2"
            libro.Sheets("ElmLne").cells(m, 4) = PositionsNodo1(i)
            libro.Sheets("ElmLne").cells(m, 5) = PositionsNodo2(i)
            libro.Sheets("ElmLne").cells(m, 6) = "##Line" & "(" & Sline(i) & ")_" & TensionNominalLines1(i) & "-" & nlnph(i) & "F"
            libro.Sheets("ElmLne").cells(m, 7) = longitud(i) / 1000
            libro.Sheets("ElmLne").cells(m, 8) = CodeLine(i)
            libro.Sheets("ElmLne").cells(m, 9) = CodeLine(i)

            LabelP1.Text = "ID(a:40)  " & consecutivo1
            LabelP2.Text = "loc_name(a:40)  " & "LIN_" & CodeLine(i)
            LabelP3.Text = "fold_id(p)  " & "##Line" & "(" & Sline(i) & ")_" & TensionNominalLines1(i) & "-" & nlnph(i) & "F"
            LabelP4.Text = "bus1(p)  " & PositionsNodo1(i)
            LabelP5.Text = "bus2(p)  " & PositionsNodo2(i)
            LabelP6.Text = "typ_id(p)  " & consecutivo3
            LabelP7.Text = "dline(r)  " & longitud(i) / 1000
            LabelP8.Text = "chr_name(a:20)  " & " "
            LabelP10.Text = " "
            LabelP11.Text = " "
            LabelP12.Text = " "
            LabelP13.Text = " "
            m = m + 1
            If i >= var1 * var2 And ProgressBar.Value < 67 Then
                ProgressBar.Value = ProgressBar.Value + 1
                LabelPorcentaje.Text = ProgressBar.Value & "%"
                var1 = var1 + 1
            End If
        Next
        ProgressBar.Value = 67
        LabelPorcentaje.Text = ProgressBar.Value & "%"



        'llenar los campos de la hoja StaCubic
        'ASIGNACION DE CUBICULOS A LAS CARGAS (LOD)
        var1 = 1
        var2 = NodoTRF.Count / 20
        m = 2
        For i = 0 To NodoTRF.Count - 1
            Application.DoEvents()
            consecutivoTRF3 = consecutivoTRF3 + 1
            consecutivoTRF2 = consecutivoTRF2 + 1
            libro.Sheets("StaCubic").cells(m, 2) = "CubLOD_" & NodoTRF(i) & "_" & TRF(i)
            libro.Sheets("StaCubic").cells(m, 1).NumberFormat = "0.00"
            libro.Sheets("StaCubic").cells(m, 1) = consecutivoTRF3
            libro.Sheets("StaCubic").cells(m, 3) = positionsNTRF(i)
            libro.Sheets("StaCubic").cells(m, 4).NumberFormat = "0.00"
            libro.Sheets("StaCubic").cells(m, 4) = consecutivoTRF2
            libro.Sheets("StaCubic").cells(m, 6) = NodoTRF(i) & "_" & TRF(i)
            libro.Sheets("StaCubic").cells(m, 7) = it2p1_TRF(i)
            libro.Sheets("StaCubic").cells(m, 8) = it2p2_TRF(i)
            LabelP1.Text = "ID(a:40)  " & consecutivoTRF3
            LabelP2.Text = "loc_name(a:40)  " & "CubLOD_" & NodoTRF(i) & "_" & TRF(i)
            LabelP3.Text = "fold_id(p)  " & positionsNTRF(i)
            LabelP4.Text = "obj_id(p)  " & consecutivoTRF2
            LabelP5.Text = "chr_name(a:20)  " & " "
            LabelP6.Text = "for_name  " & " "
            LabelP7.Text = "it2p1  " & it2p1_TRF(i)
            LabelP8.Text = "it2p2  " & it2p2_TRF(i)
            LabelP9.Text = " "
            LabelP10.Text = " "
            LabelP11.Text = " "
            LabelP12.Text = " "
            LabelP13.Text = " "
            m = m + 1
            If i >= var1 * var2 And ProgressBar.Value < 78 Then
                ProgressBar.Value = ProgressBar.Value + 1
                LabelPorcentaje.Text = ProgressBar.Value & "%"
                var1 = var1 + 1
            End If

        Next
        ProgressBar.Value = 78
        LabelPorcentaje.Text = ProgressBar.Value & "%"


        'ASIGNACION DE CUBICULOS A LAS LINEAS EN EL NODO1 (LINE)
        var1 = 1
        var2 = nodo1.Count / 20
        For i = 0 To nodo1.Count - 2
            Application.DoEvents()
            consecutivo1 = consecutivo1 + 1
            consecutivo4 = consecutivo4 + 1
            libro.Sheets("StaCubic").cells(m, 2) = "CubLINE_" & "N1_" & nodo1(i)
            libro.Sheets("StaCubic").cells(m, 1).NumberFormat = "0.00"
            libro.Sheets("StaCubic").cells(m, 1) = consecutivo1
            libro.Sheets("StaCubic").cells(m, 3) = PositionsNodo1(i)
            libro.Sheets("StaCubic").cells(m, 4).NumberFormat = "0.00"
            libro.Sheets("StaCubic").cells(m, 4) = consecutivo4
            libro.Sheets("StaCubic").cells(m, 6) = "N1_" & nodo1(i)
            libro.Sheets("StaCubic").cells(m, 7) = it2p1_Line(i)
            libro.Sheets("StaCubic").cells(m, 8) = it2p2_Line(i)
            LabelP1.Text = "ID(a:40)  " & consecutivo1
            LabelP2.Text = "loc_name(a:40)  " & "CubLINE_" & "N1_" & nodo1(i)
            LabelP3.Text = "fold_id(p)  " & PositionsNodo1(i)
            LabelP4.Text = "obj_id(p)  " & consecutivo4
            LabelP5.Text = "chr_name(a:20)  " & " "
            LabelP6.Text = "for_name  " & "N1_" & nodo1(i)
            LabelP7.Text = "it2p1  " & it2p1_Line(i)
            LabelP8.Text = "it2p2  " & it2p2_Line(i)
            LabelP9.Text = " "
            LabelP10.Text = " "
            LabelP11.Text = " "
            LabelP12.Text = " "
            LabelP13.Text = " "
            m = m + 1
            If i > var1 * var2 And ProgressBar.Value < 89 Then
                ProgressBar.Value = ProgressBar.Value + 1
                LabelPorcentaje.Text = ProgressBar.Value & "%"
                var1 = var1 + 1
            End If
        Next
        ProgressBar.Value = 89
        LabelPorcentaje.Text = ProgressBar.Value & "%"


        'ASIGNACION DE CUBICULOS A LAS LINEAS EN EL NODO2 (LINE)
        var1 = 1
        var2 = nodo1.Count / 16
        For i = 0 To nodo1.Count - 2
            Application.DoEvents()
            consecutivo1 = consecutivo1 + 1
            consecutivo5 = consecutivo5 + 1
            libro.Sheets("StaCubic").cells(m, 2) = "CubLINE_" & "N2_" & nodo2(i)
            libro.Sheets("StaCubic").cells(m, 1).NumberFormat = "0.00"
            libro.Sheets("StaCubic").cells(m, 1) = consecutivo1
            libro.Sheets("StaCubic").cells(m, 3) = PositionsNodo2(i)
            libro.Sheets("StaCubic").cells(m, 4).NumberFormat = "0.00"
            libro.Sheets("StaCubic").cells(m, 4) = consecutivo5
            libro.Sheets("StaCubic").cells(m, 6) = "N2_" & nodo2(i)
            libro.Sheets("StaCubic").cells(m, 7) = it2p1_Line(i)
            libro.Sheets("StaCubic").cells(m, 8) = it2p2_Line(i)
            LabelP1.Text = "ID(a:40)  " & consecutivo1
            LabelP2.Text = "loc_name(a:40)  " & "CubLINE_" & "N2_" & nodo2(i)
            LabelP3.Text = "fold_id(p)  " & PositionsNodo2(i)
            LabelP4.Text = "obj_id(p)  " & consecutivo5
            LabelP5.Text = "chr_name(a:20)  " & " "
            LabelP6.Text = "for_name  " & "N2_" & nodo2(i)
            LabelP7.Text = "it2p1  " & it2p1_Line(i)
            LabelP8.Text = "it2p2  " & it2p2_Line(i)
            LabelP9.Text = " "
            LabelP10.Text = " "
            LabelP11.Text = " "
            LabelP12.Text = " "
            LabelP13.Text = " "
            m = m + 1
            If i > var1 * var2 And ProgressBar.Value < 98 Then
                ProgressBar.Value = ProgressBar.Value + 1
                LabelPorcentaje.Text = ProgressBar.Value & "%"
                var1 = var1 + 1
            End If

        Next
        ProgressBar.Value = 98
        LabelPorcentaje.Text = ProgressBar.Value & "%"


        ProgressBar.Value = 99
        LabelPorcentaje.Text = "99%"


        libro.Sheets("ElmNet").cells(1, 3).EntireColumn.AutoFit()
        libro.Sheets("IntGrfnet").cells(1, 2).EntireColumn.AutoFit()
        libro.Sheets("IntGrfnet").cells(1, 3).EntireColumn.AutoFit()
        libro.Sheets("IntGrf").cells(1, 2).EntireColumn.AutoFit()
        libro.Sheets("ElmLod").cells(1, 2).EntireColumn.AutoFit()
        libro.Sheets("ElmLod").cells(1, 7).EntireColumn.AutoFit()
        libro.Sheets("ElmLne").cells(1, 2).EntireColumn.AutoFit()
        libro.Sheets("StaCubic").cells(1, 2).EntireColumn.AutoFit()
        LabelPorcentaje.Text = "100%"
        ProgressBar.Value = 100


        Application.DoEvents()
        '-----------------------------------------------------------------------------to save Excel
        SaveFileDialog1.DefaultExt = "*.xlsx"
        SaveFileDialog1.FileName = ComboBox1.Text
        SaveFileDialog1.Filter = "Archivos de Excel (*.xls)|*.xlsx"
        SaveFileDialog1.ShowDialog()

        Label25.Show()
        libro.SaveAs(SaveFileDialog1.FileName)
        Label25.Text = "El Archivo se creó y se guardó en: " & SaveFileDialog1.FileName
        LabelPorcentaje.Text = "00%"
        ProgressBar.Value = 0
        LabelP1.Text = ""
        LabelP2.Text = ""
        LabelP3.Text = ""
        LabelP4.Text = ""
        LabelP5.Text = ""
        LabelP6.Text = ""
        LabelP7.Text = ""
        LabelP8.Text = ""
        LabelP9.Text = ""
        LabelP10.Text = ""
        LabelP11.Text = ""
        LabelP12.Text = ""
        LabelP13.Text = ""


        ApExcel.Quit()
        libro = Nothing
        ApExcel = Nothing
        Hoja = Nothing
        Hoja = Nothing
        Hoja = Nothing
        Hoja = Nothing
        Hoja = Nothing

        ButCrearDGS.Enabled = True

    End Sub

    Private Sub ComboBox1_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedValueChanged
        LoadDB.Show()
        DataGridAalimentadores_F()
        Label25.Text = " "
        LoadDB.Close()
    End Sub


    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

        If CheckBox1.Checked = True Then
            TextBox1.ReadOnly = False
            TextBox2.ReadOnly = False
            TextBox3.ReadOnly = False
            TextBox4.ReadOnly = False

        Else
            TextBox1.ReadOnly = True
            TextBox2.ReadOnly = True
            TextBox3.ReadOnly = True
            TextBox4.ReadOnly = True
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox6.Text = ""
        End If



    End Sub



    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        conexion.Close()
        conexion.Close()
    End Sub

    Private Sub SalirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalirToolStripMenuItem.Click
        conexion.Close()
        conexion.Close()
        Me.Close()
    End Sub

    Private Sub AcercaDeCREADGISToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AcercaDeCREADGISToolStripMenuItem.Click
        Acercade.Show()
    End Sub

    Private Sub CargarPeriodo_Click(sender As Object, e As EventArgs)
        ButCrearDGS.Enabled = False
        Try
            ''verdad
            'LoadDB.Show()
            'LoadDB.Text = "Conectando... ORACLE"
            'conexion.Open()
            'LoadDB.Close()

            LoadDB.Show()
            LoadDB.Text = "Listando... Circuitos"
            FillComboBox()
            LoadDB.Close()

            LoadDB.Show()
            LoadDB.Text = "Cargando... DataGridView"
            DataGridAalimentadores_F()
            LoadDB.Close()
            Me.Show()

            'MsgBox("!!Conectado con Exito!!", vbInformation, "DATABASE")
        Catch ex As Exception
            'falso
            ButCrearDGS.Enabled = False
            LoadDB.Close()
            MsgBox("!!Error intentando Conectar al Servidor de datos!!", vbCritical, "DATABASE")

        End Try
        Enabledtext()

        ButCrearDGS.Enabled = True
    End Sub

    Private Sub ComboCR_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboCR.SelectedIndexChanged
        FillComboBox()
    End Sub
End Class
