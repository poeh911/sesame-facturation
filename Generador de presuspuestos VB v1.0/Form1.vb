Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data

Public Class Form1

    '****************** I N I C I O *******************************************************************

    '***************** Definición de variables*********************

    ' Datos de entrada ********************************************

    Dim wh As Double
    Dim hh As Double
    Dim pl As Double

    'Medidas de hojas completas y vidrios *************************

    Dim AHF As Integer
    Dim HHF As Integer
    Dim AHM As Integer
    Dim HHM As Integer
    Dim AVF As Integer
    Dim HVF As Integer
    Dim AVM As Integer
    Dim HVM As Integer
    Dim AHML As Integer
    Dim HHML As Integer
    Dim AVML As Integer
    Dim HVML As Integer

    ' Solapes******************************************************

    Dim solMI As Integer = 50
    Dim solFG As Integer = 30
    Dim solPL As Integer = 50
    Dim solBW52 As Integer = 74

    'Descuentos****************************************************

    Dim deshMI As Integer = 67
    Dim desvMI As Integer = 124
    Dim deshFG As Integer = 30
    Dim desvFG As Integer = 85
    Dim desv2PLF As Integer = 95
    Dim desv2PLM As Integer = 80
    Dim desv1PLF As Integer = (desv2PLF / 2) + 3.5
    Dim desv1PLM As Integer = desv2PLM / 2
    Dim deshfBW52 As Integer = 120
    Dim deshmBW52 As Integer = 130
    Dim desvBW52 As Integer = 133

    'Variables para los componentes de la carpintería ***********************************************

    ' Variables para la carpintería MI **************************************************************

    Dim a1 As Integer
    Dim a2 As Integer
    Dim a3 As Integer
    Dim a4 As Integer
    Dim a5 As Integer
    Dim a6 As Integer
    Dim a7 As Integer
    Dim a8 As Integer
    Dim a9 As Integer
    Dim a10 As Integer
    Dim a11 As Integer
    Dim a12 As Integer
    Dim a13 As Integer
    Dim a14 As Integer
    Dim a15 As Integer
    Dim a16 As Integer
    Dim a17 As Integer

    Dim au1 As Double
    Dim au2 As Double
    Dim au3 As Double
    Dim au4 As Double
    Dim au5 As Double
    Dim au6 As Double
    Dim au7 As Double
    Dim au8 As Double
    Dim au9 As Double
    Dim au10 As Double
    Dim au11 As Double
    Dim au12 As Double
    Dim au13 As Double
    Dim au14 As Double
    Dim au15 As Double
    Dim au16 As Double
    Dim au17 As Double

    ' Variables para la carpintería FG **************************************************************

    Dim b1 As Integer
    Dim b2 As Integer
    Dim b3 As Integer
    Dim b4 As Integer
    Dim b5 As Integer
    Dim b6 As Integer
    Dim b7 As Integer
    Dim b8 As Integer
    Dim b9 As Integer
    Dim b10 As Integer
    Dim b11 As Integer
    Dim b12 As Integer
    Dim b13 As Integer
    Dim b14 As Integer
    Dim b15 As Integer

    Dim bu1 As Double
    Dim bu2 As Double
    Dim bu3 As Double
    Dim bu4 As Double
    Dim bu5 As Double
    Dim bu6 As Double
    Dim bu7 As Double
    Dim bu8 As Double
    Dim bu9 As Double
    Dim bu10 As Double
    Dim bu11 As Double
    Dim bu12 As Double
    Dim bu13 As Double
    Dim bu14 As Double
    Dim bu15 As Double

    ' Variables para la carpintería PL **************************************************************

    Dim c1 As Integer
    Dim c2 As Integer
    Dim c3 As Integer
    Dim c4 As Integer
    Dim c5 As Integer
    Dim c6 As Integer
    Dim c7 As Integer
    Dim c8 As Integer

    Dim cu1 As Double
    Dim cu2 As Double
    Dim cu3 As Double
    Dim cu4 As Double
    Dim cu5 As Double
    Dim cu6 As Double
    Dim cu7 As Double
    Dim cu8 As Double

    ' Variables para la carpintería BW52 **************************************************************

    Dim d1 As Integer
    Dim d2 As Integer
    Dim d3 As Integer
    Dim d4 As Integer
    Dim d5 As Integer
    Dim d6 As Integer
    Dim d7 As Integer
    Dim d8 As Integer
    Dim d9 As Integer
    Dim d10 As Integer
    Dim d11 As Integer
    Dim d12 As Integer
    Dim d13 As Integer
    Dim d14 As Integer
    Dim d15 As Integer
    Dim d16 As Integer
    Dim d17 As Integer
    Dim d18 As Integer
    Dim d19 As Integer
    Dim d20 As Integer
    Dim d21 As Integer
    Dim d22 As Integer
    Dim d23 As Integer
    Dim d24 As Integer
    Dim d25 As Integer
    Dim d26 As Integer
    Dim d27 As Integer
    Dim d28 As Integer
    Dim d29 As Integer
    Dim d30 As Integer
    Dim d31 As Integer
    Dim d32 As Integer

    Dim du1 As Double
    Dim du2 As Double
    Dim du3 As Double
    Dim du4 As Double
    Dim du5 As Double
    Dim du6 As Double
    Dim du7 As Double
    Dim du8 As Double
    Dim du9 As Double
    Dim du10 As Double
    Dim du11 As Double
    Dim du12 As Double
    Dim du13 As Double
    Dim du14 As Double
    Dim du15 As Double
    Dim du16 As Double
    Dim du17 As Double
    Dim du18 As Double
    Dim du19 As Double
    Dim du20 As Double
    Dim du21 As Double
    Dim du22 As Double
    Dim du23 As Double
    Dim du24 As Double
    Dim du25 As Double
    Dim du26 As Double
    Dim du27 As Double
    Dim du28 As Double
    Dim du29 As Double
    Dim du30 As Double
    Dim du31 As Double
    Dim du32 As Double

    ' Variables para la carpintería AP94 **************************************************************

    'Dim e1 As Double
    'Dim e2 As Double
    'Dim e3 As Double
    'Dim e4 As Double
    'Dim e5 As Double
    'Dim e6 As Double
    'Dim e7 As Double
    'Dim e8 As Double
    'Dim e9 As Double
    'Dim e10 As Double
    'Dim e11 As Double
    'Dim e12 As Double
    'Dim e13 As Double
    'Dim e14 As Double
    'Dim e15 As Double
    'Dim e16 As Double
    'Dim e17 As Double
    'Dim e18 As Double
    'Dim e19 As Double
    'Dim e20 As Double
    'Dim e21 As Double
    'Dim e22 As Double
    'Dim e23 As Double
    'Dim e24 As Double
    'Dim e25 As Double
    'Dim e26 As Double
    'Dim e27 As Double
    'Dim e28 As Double
    'Dim e29 As Double

    'Dim eu1 As Double
    'Dim eu2 As Double
    'Dim eu3 As Double
    'Dim eu4 As Double
    'Dim eu5 As Double
    'Dim eu6 As Double
    'Dim eu7 As Double
    'Dim eu8 As Double
    'Dim eu9 As Double
    'Dim eu10 As Double
    'Dim eu11 As Double
    'Dim eu12 As Double
    'Dim eu13 As Double
    'Dim eu14 As Double
    'Dim eu15 As Double
    'Dim eu16 As Double
    'Dim eu17 As Double
    'Dim eu18 As Double
    'Dim eu19 As Double
    'Dim eu20 As Double
    'Dim eu21 As Double
    'Dim eu22 As Double
    'Dim eu23 As Double
    'Dim eu24 As Double
    'Dim eu25 As Double
    'Dim eu26 As Double
    'Dim eu27 As Double
    'Dim eu28 As Double
    'Dim eu29 As Double

    ' Variables fila y columna para los presupuestos ***********************************************************

    Dim fil As Integer
    Dim col As Integer

    ' Variables de coste de los lacados *************************************************************************

    Dim pg As Double
    Dim cpl As Double

    ' Variable de coste de los grupos de anodizados *************************************************************

    Dim pa As Double
    Dim pgr As Double = 0.08

    ' Variable bandera para marcar el final de los eventos ******************************************************

    Dim bFlag As Boolean = False


#Region "Creación de los items en los ComboBox"

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ' Añadir barra de scroll automática 

        Me.VerticalScroll.Visible = True
        Me.VerticalScroll.Enabled = True

        Me.HorizontalScroll.Visible = True
        Me.HorizontalScroll.Enabled = True

        ' No visualización de la pestaña de la carpintería AP94  ****************************

        Panel7.Visible = False

        ' Esconder tabpage de presupuestos y ral *************************************************

        Presupuesto.Parent = Nothing
        TabPage6.Parent = Nothing

        ' Esconder inicialmente el panel con los artículos opcionales para la carpintería *******

        Panel8.Visible = False
        Panel9.Visible = False

        ' Etiquetas no visibles para las configuraciones especiales (No sabemos si se van a incluir) ******** 

        Label26.Visible = False
        CheckBox1.Visible = False
        Label157.Visible = False

        ' Ocultar panel de las hojas móviles lentas hasta que se seleccione ******************

        Panel11.Visible = False
        Panel12.Visible = False
        Panel13.Visible = False
        Panel14.Visible = False

        ' Ocultar letrero "Rápida" de la hoja móvil para cuando no sea telescópica ***********

        Label189.Visible = False
        Label198.Visible = False
        Label204.Visible = False
        Label209.Visible = False

        ' Ocultar cortavientos TES para obra ***************************

        Label222.Visible = False
        a16text.Visible = False
        au16text.Visible = False

        '****************** INTRODUCCIÓN DE DATOS EN LOS COMBO BOX *******************************

        ' COMBO BOX MODELOS DE OPERADORES ********************************************************
        'Creación de la variable tabla para los diferentes modelos de operadores'

        Dim dt As DataTable = New DataTable("Operadores")
        dt.Columns.Add("Código")
        dt.Columns.Add("Modelo")
        Dim dr As DataRow

        ' Adición de cada uno de los modelos de operador'

        dr = dt.NewRow()
        dr("Código") = ""
        dr("Modelo") = ""
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Código") = 2001201
        dr("Modelo") = "MI14 D. Corredera Doble. 2 Hojas, 140 Kg máx. Incluye batería y selector de 4 posiciones"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Código") = 2001202
        dr("Modelo") = "MI14 S. Corredera Simple. 1 Hoja, 100 Kg máx. Incluye batería y selector de 4 posiciones"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Código") = 2001075
        dr("Modelo") = "MI50 D. Corredera Doble. 2 Hojas, 180 Kg máx. Incluye batería y selector de 4 posiciones"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Código") = 2001077
        dr("Modelo") = "MI50 S. Corredera Simple. 1 Hoja, 110 Kg máx. Incluye batería y selector de 4 posiciones"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Código") = 2001120
        dr("Modelo") = "MI75 D. Corredera Doble. 2 Hojas, 250 Kg máx. Incluye batería y selector de 4 posiciones"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Código") = 2001121
        dr("Modelo") = "MI75 S. Corredera Simple. 1 Hoja, 180 Kg máx. Incluye batería y selector de 4 posiciones"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Código") = 2001105
        dr("Modelo") = "MI100 D. Corredera Doble. 2 Hojas, 350 Kg máx. Incluye batería y selector de 4 posiciones"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Código") = 2001106
        dr("Modelo") = "MI100 S. Corredera Simple. 1 Hojas, 250 Kg máx. Incluye batería y selector de 4 posiciones"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Código") = 2001143
        dr("Modelo") = "MI50 TES TH4. Telescópica. 4 Hojas, 45 x 4 Kg máx. Incluye batería y selector de 4 posiciones"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Código") = 2001142
        dr("Modelo") = "MI50 TES TH2. Telescópica. 2 Hojas, 45 x 2 Kg máx. Incluye batería y selector de 4 posiciones"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Código") = 2002279
        dr("Modelo") = "MI SW Push. Operador para puertas batientes con brazo articulado"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Código") = 2002280
        dr("Modelo") = "MI SW Pull. Operador para puertas batientes con brazo deslizante"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Código") = 2002281
        dr("Modelo") = "MI SWSP Push. Operador para puertas batientes con brazo articulado y cierre por muelle"
        dt.Rows.Add(dr)

        dr = dt.NewRow()
        dr("Código") = 2002282
        dr("Modelo") = "MI SWSP Pull. Operador para puertas batientes con brazo deslizante y cierre por muelle"
        dt.Rows.Add(dr)

        ' Asociar la tabla al combobox'

        cmbmodope.DataSource = dt
        cmbmodope.ValueMember = "Código"
        cmbmodope.DisplayMember = "Modelo"
        cmbmodope.Update()

        ' COMBO BOX MODELOS DE CARPINTERÍAS ********************************************************
        'Creación de la variable tabla con los diferentes modelos de carpinterías'

        Dim dt1 As DataTable = New DataTable("Carpinterías")
        dt1.Columns.Add("Código")
        dt1.Columns.Add("Modelo")
        Dim dr1 As DataRow

        'Añadir los modelos de carpintería'

        dr1 = dt1.NewRow()
        dr1("Modelo") = ""
        dt1.Rows.Add(dr1)

        dr1 = dt1.NewRow()
        dr1("Código") = 9980001
        dr1("Modelo") = "MI"
        dt1.Rows.Add(dr1)

        dr1 = dt1.NewRow()
        dr1("Código") = 9980002
        dr1("Modelo") = "Full Glass"
        dt1.Rows.Add(dr1)

        dr1 = dt1.NewRow()
        dr1("Código") = 9900001
        dr1("Modelo") = "Plintos silicona superior e inferior"
        dt1.Rows.Add(dr1)

        dr1 = dt1.NewRow()
        dr1("Código") = 9900002
        dr1("Modelo") = "Plinto silicona. Sólo superior"
        dt1.Rows.Add(dr1)

        dr1 = dt1.NewRow()
        dr1("Código") = 9900003
        dr1("Modelo") = "Plinto superior de pinza e inferior de silicona"
        dt1.Rows.Add(dr1)

        dr1 = dt1.NewRow()
        dr1("Código") = 9900004
        dr1("Modelo") = "Plinto de pinza. Sólo superior"
        dt1.Rows.Add(dr1)

        dr1 = dt1.NewRow()
        dr1("Código") = 9980005
        dr1("Modelo") = "BW52"
        dt1.Rows.Add(dr1)

        'dr1 = dt1.NewRow()
        'dr1("Modelo") = "AP94"
        'dt1.Rows.Add(dr1)

        ' Asociar los datos de la tabla al combobox'

        cmbmodcar.DataSource = dt1
        cmbmodcar.ValueMember = "Código"
        cmbmodcar.DisplayMember = "Modelo"
        cmbmodcar.Update()

        ' COMBO BOX CONFIGURACIONES DE LAS HOJAS ********************************************************
        'Creación de la variable tabla con las diferentes configuraciones de las hojas'

        Dim dt2 As DataTable = New DataTable("Configuración")
        dt2.Columns.Add("Modelo")
        Dim dr2 As DataRow

        'Añadir los modelos de carpintería'

        dr2 = dt2.NewRow()
        dr2("Modelo") = ""
        dt2.Rows.Add(dr2)

        dr2 = dt2.NewRow()
        dr2("Modelo") = "2 Hojas Fijas + 2 Hojas Móviles"
        dt2.Rows.Add(dr2)

        dr2 = dt2.NewRow()
        dr2("Modelo") = "2 Hojas Móviles"
        dt2.Rows.Add(dr2)

        dr2 = dt2.NewRow()
        dr2("Modelo") = "1 Hoja Fija + 1 Hoja Móvil"
        dt2.Rows.Add(dr2)

        dr2 = dt2.NewRow()
        dr2("Modelo") = "1 Hoja Móvil"
        dt2.Rows.Add(dr2)

        dr2 = dt2.NewRow()
        dr2("Modelo") = "2 Hojas Fijas + 4 Hojas Móviles"
        dt2.Rows.Add(dr2)

        dr2 = dt2.NewRow()
        dr2("Modelo") = "4 Hojas Móviles"
        dt2.Rows.Add(dr2)

        dr2 = dt2.NewRow()
        dr2("Modelo") = "1 Hoja Fija + 2 Hojas Móviles"
        dt2.Rows.Add(dr2)

        dr2 = dt2.NewRow()
        dr2("Modelo") = "2 Hojas Móviles TES"
        dt2.Rows.Add(dr2)

        ' Asociar los datos de la tabla al combobox

        cmbconf.DataSource = dt2
        cmbconf.DisplayMember = "Modelo"
        cmbconf.Update()

        ' COMBO BOX RADARES Y SEGURIDAD 1 ********************************************************
        'Creación de la variable tabla para los diferentes radares

        Dim dt3 As DataTable = New DataTable("Selectores")
        dt3.Columns.Add("Código")
        dt3.Columns.Add("Modelo")
        Dim dr3 As DataRow

        'Añadir los radares

        dr3 = dt3.NewRow()
        dr3("Modelo") = ""
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001087
        dr3("Modelo") = "Radar modelo LOBO-I"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001058
        dr3("Modelo") = "Radar modelo LOBO-II"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001104
        dr3("Modelo") = "Radar a efecto Doppler mod. MW2, bidireccional"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001063
        dr3("Modelo") = "Radar mod. OA-AXIS II detección por movimiento y presencia, con cortina de seguridad"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001089
        dr3("Modelo") = "Radar mod. OA-AXIS T monitoring"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001100
        dr3("Modelo") = "Radar mod. OA-203-2 con dos líneas, detección y presencia"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001116
        dr3("Modelo") = "Radar de detección mod. COLIBRÍ-I, unidireccional"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001117
        dr3("Modelo") = "Radar de detección mod. COLIBRÍ-II, bidireccional"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001118
        dr3("Modelo") = "Radar mod. VIO-DT unidireccional"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001119
        dr3("Modelo") = "Radar mod. VIO-DT2 bidireccional"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001156
        dr3("Modelo") = "Radar mod. IXIO-DT3, detección unidireccional"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001061
        dr3("Modelo") = "Radar modelo LOBO SLIM"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001088
        dr3("Modelo") = "Kit Radar Modelo ACTIV-8 ONE ON y SLIM ONE"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001045
        dr3("Modelo") = "Radar mod. ACTIV-8 bidireccional, cortina de seguridad de presencia"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001086
        dr3("Modelo") = "Radar mod. ACTIV-8 unidireccional, cortina de seguridad de presencia"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001109
        dr3("Modelo") = "Radar mod. ACTIV-8 on unidireccional, cortina de seguridad de presencia"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001025
        dr3("Modelo") = "Radar modelo CRYSTAL, por infrarrojos activos"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001052
        dr3("Modelo") = "Radar digital mod. S-87 para puerta industrial"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001011
        dr3("Modelo") = "Radar modelo RBN-4300 a efecto Doppler"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001041
        dr3("Modelo") = "Radar modelo AGUILA-I a efecto Doppler, unidireccional"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001039
        dr3("Modelo") = "Radar modelo AGUILA-II a efecto Doppler"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001057
        dr3("Modelo") = "Buscador de infrarrojo para AXIS"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001046
        dr3("Modelo") = "Buscador de infrarrojo para ACTIV-8"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001044
        dr3("Modelo") = "Telemando ajuste a distancia radar AGUILA"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001042
        dr3("Modelo") = "Porta radar AGUILA para cielo raso mod. ECA"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001064
        dr3("Modelo") = "Porta radar ACTIV-8 para cielo raso mod. ACA"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001065
        dr3("Modelo") = "Porta radar ACTIV-8 para techo  mod.ACA"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001043
        dr3("Modelo") = "Protector de agua para radar AGUILA mod. ERA"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001051
        dr3("Modelo") = "Detector seguridad mod. EYE puertas abatibles"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001122
        dr3("Modelo") = "Borde de seguridad mod. 4SAFE monitorizado, para puertas abatibles"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001126
        dr3("Modelo") = "Borde de seguridad OA-EDGE T M 700 MASTER, monitorizado para puertas abatibles"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001127
        dr3("Modelo") = "Borde de seguridad OA-EDGE T S 700 ESCLAVO, monitorizado para puertas abatibles"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001124
        dr3("Modelo") = "Detector de seguridad mod. 1SAFE"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001047
        dr3("Modelo") = "Célula de seguridad mod. MICROCELL ONE-S"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001048
        dr3("Modelo") = "Célula de seguridad mod. MICROCELL ONE-SL"
        dt3.Rows.Add(dr3)

        dr3 = dt3.NewRow()
        dr3("Código") = 2001080
        dr3("Modelo") = "Soporte de superficie para MICROCELL"
        dt3.Rows.Add(dr3)

        ' Asociar los datos de la tabla a los combobox'

        cmbrad1.DataSource = dt3
        cmbrad1.ValueMember = "Código"
        cmbrad1.DisplayMember = "Modelo"
        cmbrad1.Update()

        ' COMBO BOX RADARES Y SEGURIDAD 2 ********************************************************
        'Creación de la variable tabla para los diferentes radares

        Dim dt4 As DataTable = New DataTable("Selectores")
        dt4.Columns.Add("Código")
        dt4.Columns.Add("Modelo")
        Dim dr4 As DataRow

        'Añadir los radares

        dr4 = dt4.NewRow()
        dr4("Modelo") = ""
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001087
        dr4("Modelo") = "Radar modelo LOBO-I"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001058
        dr4("Modelo") = "Radar modelo LOBO-II"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001104
        dr4("Modelo") = "Radar a efecto Doppler mod. MW2, bidireccional"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001063
        dr4("Modelo") = "Radar mod. OA-AXIS II detección por movimiento y presencia, con cortina de seguridad"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001089
        dr4("Modelo") = "Radar mod. OA-AXIS T monitoring"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001100
        dr4("Modelo") = "Radar mod. OA-203-2 con dos líneas, detección y presencia"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001116
        dr4("Modelo") = "Radar de detección mod. COLIBRÍ-I, unidireccional"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001117
        dr4("Modelo") = "Radar de detección mod. COLIBRÍ-II, bidireccional"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001118
        dr4("Modelo") = "Radar mod. VIO-DT unidireccional"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001119
        dr4("Modelo") = "Radar mod. VIO-DT2 bidireccional"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001156
        dr4("Modelo") = "Radar mod. IXIO-DT3, detección unidireccional"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001061
        dr4("Modelo") = "Radar modelo LOBO SLIM"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001088
        dr4("Modelo") = "Kit Radar Modelo ACTIV-8 ONE ON y SLIM ONE"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001045
        dr4("Modelo") = "Radar mod. ACTIV-8 bidireccional, cortina de seguridad de presencia"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001086
        dr4("Modelo") = "Radar mod. ACTIV-8 unidireccional, cortina de seguridad de presencia"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001109
        dr4("Modelo") = "Radar mod. ACTIV-8 on unidireccional, cortina de seguridad de presencia"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001025
        dr4("Modelo") = "Radar modelo CRYSTAL, por infrarrojos activos"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001052
        dr4("Modelo") = "Radar digital mod. S-87 para puerta industrial"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001011
        dr4("Modelo") = "Radar modelo RBN-4300 a efecto Doppler"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001041
        dr4("Modelo") = "Radar modelo AGUILA-I a efecto Doppler, unidireccional"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001039
        dr4("Modelo") = "Radar modelo AGUILA-II a efecto Doppler"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001057
        dr4("Modelo") = "Buscador de infrarrojo para AXIS"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001046
        dr4("Modelo") = "Buscador de infrarrojo para ACTIV-8"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001044
        dr4("Modelo") = "Telemando ajuste a distancia radar AGUILA"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001042
        dr4("Modelo") = "Porta radar AGUILA para cielo raso mod. ECA"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001064
        dr4("Modelo") = "Porta radar ACTIV-8 para cielo raso mod. ACA"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001065
        dr4("Modelo") = "Porta radar ACTIV-8 para techo  mod.ACA"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001043
        dr4("Modelo") = "Protector de agua para radar AGUILA mod. ERA"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001051
        dr4("Modelo") = "Detector seguridad mod. EYE puertas abatibles"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001122
        dr4("Modelo") = "Borde de seguridad mod. 4SAFE monitorizado, para puertas abatibles"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001126
        dr4("Modelo") = "Borde de seguridad OA-EDGE T M 700 MASTER, monitorizado para puertas abatibles"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001127
        dr4("Modelo") = "Borde de seguridad OA-EDGE T S 700 ESCLAVO, monitorizado para puertas abatibles"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001124
        dr4("Modelo") = "Detector de seguridad mod. 1SAFE"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001047
        dr4("Modelo") = "Célula de seguridad mod. MICROCELL ONE-S"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001048
        dr4("Modelo") = "Célula de seguridad mod. MICROCELL ONE-SL"
        dt4.Rows.Add(dr4)

        dr4 = dt4.NewRow()
        dr4("Código") = 2001080
        dr4("Modelo") = "Soporte de superficie para MICROCELL"
        dt4.Rows.Add(dr4)

        ' Asociar los datos de la tabla a los combobox'

        cmbrad2.DataSource = dt4
        cmbrad2.ValueMember = "Código"
        cmbrad2.DisplayMember = "Modelo"
        cmbrad2.Update()

        ' COMBO BOX RADARES Y SEGURIDAD 3 ********************************************************
        'Creación de la variable tabla para los diferentes radares

        Dim dt5 As DataTable = New DataTable("Selectores")
        dt5.Columns.Add("Código")
        dt5.Columns.Add("Modelo")
        Dim dr5 As DataRow

        'Añadir los radares

        dr5 = dt5.NewRow()
        dr5("Modelo") = ""
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001087
        dr5("Modelo") = "Radar modelo LOBO-I"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001058
        dr5("Modelo") = "Radar modelo LOBO-II"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001104
        dr5("Modelo") = "Radar a efecto Doppler mod. MW2, bidireccional"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001063
        dr5("Modelo") = "Radar mod. OA-AXIS II detección por movimiento y presencia, con cortina de seguridad"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001089
        dr5("Modelo") = "Radar mod. OA-AXIS T monitoring"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001100
        dr5("Modelo") = "Radar mod. OA-203-2 con dos líneas, detección y presencia"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001116
        dr5("Modelo") = "Radar de detección mod. COLIBRÍ-I, unidireccional"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001117
        dr5("Modelo") = "Radar de detección mod. COLIBRÍ-II, bidireccional"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001118
        dr5("Modelo") = "Radar mod. VIO-DT unidireccional"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001119
        dr5("Modelo") = "Radar mod. VIO-DT2 bidireccional"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001156
        dr5("Modelo") = "Radar mod. IXIO-DT3, detección unidireccional"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001061
        dr5("Modelo") = "Radar modelo LOBO SLIM"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001088
        dr5("Modelo") = "Kit Radar Modelo ACTIV-8 ONE ON y SLIM ONE"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001045
        dr5("Modelo") = "Radar mod. ACTIV-8 bidireccional, cortina de seguridad de presencia"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001086
        dr5("Modelo") = "Radar mod. ACTIV-8 unidireccional, cortina de seguridad de presencia"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001109
        dr5("Modelo") = "Radar mod. ACTIV-8 on unidireccional, cortina de seguridad de presencia"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001025
        dr5("Modelo") = "Radar modelo CRYSTAL, por infrarrojos activos"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001052
        dr5("Modelo") = "Radar digital mod. S-87 para puerta industrial"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001011
        dr5("Modelo") = "Radar modelo RBN-4300 a efecto Doppler"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001041
        dr5("Modelo") = "Radar modelo AGUILA-I a efecto Doppler, unidireccional"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001039
        dr5("Modelo") = "Radar modelo AGUILA-II a efecto Doppler"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001057
        dr5("Modelo") = "Buscador de infrarrojo para AXIS"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001046
        dr5("Modelo") = "Buscador de infrarrojo para ACTIV-8"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001044
        dr5("Modelo") = "Telemando ajuste a distancia radar AGUILA"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001042
        dr5("Modelo") = "Porta radar AGUILA para cielo raso mod. ECA"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001064
        dr5("Modelo") = "Porta radar ACTIV-8 para cielo raso mod. ACA"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001065
        dr5("Modelo") = "Porta radar ACTIV-8 para techo  mod.ACA"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001043
        dr5("Modelo") = "Protector de agua para radar AGUILA mod. ERA"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001051
        dr5("Modelo") = "Detector seguridad mod. EYE puertas abatibles"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001122
        dr5("Modelo") = "Borde de seguridad mod. 4SAFE monitorizado, para puertas abatibles"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001126
        dr5("Modelo") = "Borde de seguridad OA-EDGE T M 700 MASTER, monitorizado para puertas abatibles"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001127
        dr5("Modelo") = "Borde de seguridad OA-EDGE T S 700 ESCLAVO, monitorizado para puertas abatibles"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001124
        dr5("Modelo") = "Detector de seguridad mod. 1SAFE"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001047
        dr5("Modelo") = "Célula de seguridad mod. MICROCELL ONE-S"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001048
        dr5("Modelo") = "Célula de seguridad mod. MICROCELL ONE-SL"
        dt5.Rows.Add(dr5)

        dr5 = dt5.NewRow()
        dr5("Código") = 2001080
        dr5("Modelo") = "Soporte de superficie para MICROCELL"
        dt5.Rows.Add(dr5)

        ' Asociar los datos de la tabla a los combobox

        cmbrad3.DataSource = dt5
        cmbrad3.ValueMember = "Código"
        cmbrad3.DisplayMember = "Modelo"
        cmbrad3.Update()

        ' COMBOBOX PULSADORES 1 *******************************************************************************************************
        'Creación de la variable tabla para los diferentes pulsadores

        Dim dt6 As DataTable = New DataTable("Pulsadores")
        dt6.Columns.Add("Código")
        dt6.Columns.Add("Modelo")
        Dim dr6 As DataRow

        'Añadir los pulsadores

        dr6 = dt6.NewRow()
        dr6("Modelo") = ""
        dt6.Rows.Add(dr6)

        dr6 = dt6.NewRow()
        dr6("Código") = 2001095
        dr6("Modelo") = "Pulsador verde Francia"
        dt6.Rows.Add(dr6)

        dr6 = dt6.NewRow()
        dr6("Código") = 2001055
        dr6("Modelo") = "Pulsador de codo en acero inoxidable"
        dt6.Rows.Add(dr6)

        dr6 = dt6.NewRow()
        dr6("Código") = 2001020
        dr6("Modelo") = "Pulsador manos libres mod. MAGIC, se activa por presencia"
        dt6.Rows.Add(dr6)

        ' Asociar los datos de la tabla a los combobox'

        cmbpul1.DataSource = dt6
        cmbpul1.ValueMember = "Código"
        cmbpul1.DisplayMember = "Modelo"
        cmbpul1.Update()

        ' COMBOBOX PULSADORES 2 *******************************************************************************************************
        'Creación de la variable tabla para los diferentes pulsadores

        Dim dt7 As DataTable = New DataTable("Pulsadores")
        dt7.Columns.Add("Código")
        dt7.Columns.Add("Modelo")
        Dim dr7 As DataRow

        'Añadir los pulsadores'

        dr7 = dt7.NewRow()
        dr7("Modelo") = ""
        dt7.Rows.Add(dr7)

        dr7 = dt7.NewRow()
        dr7("Código") = 2001095
        dr7("Modelo") = "Pulsador verde Francia"
        dt7.Rows.Add(dr7)

        dr7 = dt7.NewRow()
        dr7("Código") = 2001055
        dr7("Modelo") = "Pulsador de codo en acero inoxidable"
        dt7.Rows.Add(dr7)

        dr7 = dt7.NewRow()
        dr7("Código") = 2001020
        dr7("Modelo") = "Pulsador manos libres mod. MAGIC, se activa por presencia"
        dt7.Rows.Add(dr7)

        ' Asociar los datos de la tabla a los combobox'

        cmbpul2.DataSource = dt7
        cmbpul2.ValueMember = "Código"
        cmbpul2.DisplayMember = "Modelo"
        cmbpul2.Update()

        ' COMBOBOX SELECTORES *******************************************************************************************************
        'Creación de la variable tabla para los diferentes selectores'

        Dim dt8 As DataTable = New DataTable("Selectores de maniobra")
        dt8.Columns.Add("Código")
        dt8.Columns.Add("Modelo")
        Dim dr8 As DataRow

        'Añadir los selectores de maniobra'

        dr8 = dt8.NewRow()
        dr8("Modelo") = ""
        dt8.Rows.Add(dr8)

        dr8 = dt8.NewRow()
        dr8("Código") = 2002285
        dr8("Modelo") = "Selector B-6 Digital SWD5"
        dt8.Rows.Add(dr8)

        dr8 = dt8.NewRow()
        dr8("Código") = 2002119
        dr8("Modelo") = "Selector de maniobra rotativo A4"
        dt8.Rows.Add(dr8)

        dr8 = dt8.NewRow()
        dr8("Código") = 2002109
        dr8("Modelo") = "Selector de maniobra digital mod. SLD5"
        dt8.Rows.Add(dr8)

        dr8 = dt8.NewRow()
        dr8("Código") = 2001013
        dr8("Modelo") = "Selector de maniobra mod. SMI-4P, de 4 posiciones"
        dt8.Rows.Add(dr8)

        dr8 = dt8.NewRow()
        dr8("Código") = 6904076
        dr8("Modelo") = "Cerradura de contacto para empotrar, mod. MOSEU-NICE"
        dt8.Rows.Add(dr8)

        dr8 = dt8.NewRow()
        dr8("Código") = 6904074
        dr8("Modelo") = "Cerradura de contacto de superficie, mod. MOSEU-NICE"
        dt8.Rows.Add(dr8)

        ' Asociar los datos de la tabla a los combobox'

        cmbsel.DataSource = dt8
        cmbsel.ValueMember = "Código"
        cmbsel.DisplayMember = "Modelo"
        cmbsel.Update()

        ' COMBOBOX CERROJOS  *******************************************************************************************************
        'Creación de la variable tabla para los diferentes cerrojos'

        Dim dt9 As DataTable = New DataTable("Cerrojos")
        dt9.Columns.Add("Código")
        dt9.Columns.Add("Modelo")
        Dim dr9 As DataRow

        'Añadir los diferentes modelos de cerrojos'

        dr9 = dt9.NewRow()
        dr9("Modelo") = ""
        dt9.Rows.Add(dr9)

        dr9 = dt9.NewRow()
        dr9("Código") = 2001101
        dr9("Modelo") = "Cerrojo MCR-1 TES (Kit)"
        dt9.Rows.Add(dr9)

        dr9 = dt9.NewRow()
        dr9("Código") = 2002309
        dr9("Modelo") = "Cerrojo MI100/2015 (Kit)"
        dt9.Rows.Add(dr9)

        dr9 = dt9.NewRow()
        dr9("Código") = 2001097
        dr9("Modelo") = "Cerrojo electromagnético mod. MCR-1, bloquea sin tensión"
        dt9.Rows.Add(dr9)

        dr9 = dt9.NewRow()
        dr9("Código") = 2002292
        dr9("Modelo") = "Cerrojo MCR VSLIM (Kit)"
        dt9.Rows.Add(dr9)

        dr9 = dt9.NewRow()
        dr9("Código") = 2002265
        dr9("Modelo") = "Cerrojo electromagnético mod. MCR-1 MI-100"
        dt9.Rows.Add(dr9)

        dr9 = dt9.NewRow()
        dr9("Código") = 2001098
        dr9("Modelo") = "Cerrojo electromagnético MCR-2"
        dt9.Rows.Add(dr9)

        dr9 = dt9.NewRow()
        dr9("Código") = 2002361
        dr9("Modelo") = "Cerrojo MCR-14"
        dt9.Rows.Add(dr9)

        dr9 = dt9.NewRow()
        dr9("Código") = 2001123
        dr9("Modelo") = "Patín para cerrojo MCR"
        dt9.Rows.Add(dr9)

        dr9 = dt9.NewRow()
        dr9("Código") = 2002284
        dr9("Modelo") = "MCR-1 con desbloqueo exterior"
        dt9.Rows.Add(dr9)

        ' Asociar los datos de la tabla a los combobox'

        cmbcerr.DataSource = dt9
        cmbcerr.ValueMember = "Código"
        cmbcerr.DisplayMember = "Modelo"
        cmbcerr.Update()

        ' COMBOBOX OTROS 1 *******************************************************************************************************
        'Creación de la variable tabla para los diferentes accesorios

        Dim dt10 As DataTable = New DataTable("Otros")
        dt10.Columns.Add("Código")
        dt10.Columns.Add("Modelo")
        Dim dr10 As DataRow

        'Añadir los diferentes modelos de cerrojos

        dr10 = dt10.NewRow()
        dr10("Modelo") = ""
        dt10.Rows.Add(dr10)

        dr10 = dt10.NewRow()
        dr10("Código") = 2002158
        dr10("Modelo") = "Patín de seguridad mod. MGNY"
        dt10.Rows.Add(dr10)

        dr10 = dt10.NewRow()
        dr10("Código") = 2002069
        dr10("Modelo") = "Patín sólo vidrio 8/12 mm, regulable"
        dt10.Rows.Add(dr10)

        dr10 = dt10.NewRow()
        dr10("Código") = 2501023
        dr10("Modelo") = "Guía de seguridad de aluminio mod. MGO"
        dt10.Rows.Add(dr10)

        dr10 = dt10.NewRow()
        dr10("Código") = 2002001
        dr10("Modelo") = "Patin inferior (MI/Full-Glass)"
        dt10.Rows.Add(dr10)

        dr10 = dt10.NewRow()
        dr10("Código") = 2002223
        dr10("Modelo") = "Patín de suelo o ángulo para BW52"
        dt10.Rows.Add(dr10)

        'dr10 = dt10.NewRow()
        'dr10("Código") = 2002051
        'dr10("Modelo") = "Patín inferior AP94"
        'dt10.Rows.Add(dr10)

        'dr10 = dt10.NewRow()
        'dr10("Código") = 2501004
        'dr10("Modelo") = "Guía inferior para puerta antipánico integral"
        'dt10.Rows.Add(dr10)

        dr10 = dt10.NewRow()
        dr10("Código") = 5001030
        dr10("Modelo") = "Receptor 2 canales 12/24V Base 30-2B"
        dt10.Rows.Add(dr10)

        dr10 = dt10.NewRow()
        dr10("Código") = 5001033
        dr10("Modelo") = "Receptor 1 canal 12/24V Base 30-1B"
        dt10.Rows.Add(dr10)

        dr10 = dt10.NewRow()
        dr10("Código") = 3501058
        dr10("Modelo") = "Emisor doble, código cambiante, Mod. 60-02-mini"
        dt10.Rows.Add(dr10)

        ' Asociar los datos de la tabla a los combobox'

        cmbotr1.DataSource = dt10
        cmbotr1.ValueMember = "Código"
        cmbotr1.DisplayMember = "Modelo"
        cmbotr1.Update()

        ' COMBOBOX OTROS 2 *******************************************************************************************************
        'Creación de la variable tabla para los diferentes accesorios

        Dim dt11 As DataTable = New DataTable("Otros")
        dt11.Columns.Add("Código")
        dt11.Columns.Add("Modelo")
        Dim dr11 As DataRow

        'Añadir los diferentes modelos de cerrojos

        dr11 = dt11.NewRow()
        dr11("Modelo") = ""
        dt11.Rows.Add(dr11)

        dr11 = dt11.NewRow()
        dr11("Código") = 2002158
        dr11("Modelo") = "Patín de seguridad mod. MGNY"
        dt11.Rows.Add(dr11)

        dr11 = dt11.NewRow()
        dr11("Código") = 2002069
        dr11("Modelo") = "Patín sólo vidrio 8/12 mm, regulable"
        dt11.Rows.Add(dr11)

        dr11 = dt11.NewRow()
        dr11("Código") = 2501023
        dr11("Modelo") = "Guía de seguridad de aluminio mod. MGO"
        dt11.Rows.Add(dr11)

        dr11 = dt11.NewRow()
        dr11("Código") = 2002001
        dr11("Modelo") = "Patin inferior (MI/Full-Glass)"
        dt11.Rows.Add(dr11)

        dr11 = dt11.NewRow()
        dr11("Código") = 2002223
        dr11("Modelo") = "Patín de suelo o ángulo para BW52"
        dt11.Rows.Add(dr11)

        'dr11 = dt11.NewRow()
        'dr11("Código") = 2002051
        'dr11("Modelo") = "Patín inferior AP94"
        'dt11.Rows.Add(dr11)

        'dr11 = dt11.NewRow()
        'dr11("Código") = 2501004
        'dr11("Modelo") = "Guía inferior para puerta antipánico integral"
        'dt11.Rows.Add(dr11)

        dr11 = dt11.NewRow()
        dr11("Código") = 5001030
        dr11("Modelo") = "Receptor 2 canales 12/24V Base 30-2B"
        dt11.Rows.Add(dr11)

        dr11 = dt11.NewRow()
        dr11("Código") = 5001033
        dr11("Modelo") = "Receptor 1 canal 12/24V Base 30-1B"
        dt11.Rows.Add(dr11)

        dr11 = dt11.NewRow()
        dr11("Código") = 3501058
        dr11("Modelo") = "Emisor doble, código cambiante, Mod. 60-02-mini"
        dt11.Rows.Add(dr11)

        ' Asociar los datos de la tabla a los combobox'

        cmbotr2.DataSource = dt11
        cmbotr2.ValueMember = "Código"
        cmbotr2.DisplayMember = "Modelo"
        cmbotr2.Update()

        ' COMBOBOX FORRO VIGA PORTA-OPERADOR *******************************************************************************************
        'Creación de la variable tabla para los diferentes forros para las vigas porta-operador

        Dim dt12 As DataTable = New DataTable("Forros viga porta-operador")
        dt12.Columns.Add("Código")
        dt12.Columns.Add("Modelo")
        Dim dr12 As DataRow

        'Añadir los forros para vigas porta-operador

        dr12 = dt12.NewRow()
        dr12("Modelo") = ""
        dt12.Rows.Add(dr12)

        dr12 = dt12.NewRow()
        dr12("Código") = 2251055
        dr12("Modelo") = "Forro de viga de 160, hasta 4000 mm, en chapa de aluminio lacado o anodizado"
        dt12.Rows.Add(dr12)

        dr12 = dt12.NewRow()
        dr12("Código") = 2251056
        dr12("Modelo") = "Forro de viga de 160 a 250 mm, hasta 6000 mm, en chapa de aluminio lacado o anodizado"
        dt12.Rows.Add(dr12)

        dr12 = dt12.NewRow()
        dr12("Código") = 2251069
        dr12("Modelo") = "Forro acero inox. - 304 para viga de 160 hasta 4000 mm."
        dt12.Rows.Add(dr12)

        dr12 = dt12.NewRow()
        dr12("Código") = 2251070
        dr12("Modelo") = "Forro acero inox. - 316 para viga de 160 hasta 4000 mm."
        dt12.Rows.Add(dr12)

        dr12 = dt12.NewRow()
        dr12("Código") = 2251071
        dr12("Modelo") = "Forro acero inox. - 304 para viga de 160 a 250 mm hasta 6000 mm."
        dt12.Rows.Add(dr12)

        dr12 = dt12.NewRow()
        dr12("Código") = 2501072
        dr12("Modelo") = "Forro acero inox. - 316 para viga de 160 a 250 mm hasta 6000 mm."
        dt12.Rows.Add(dr12)

        ' Asociar los datos de la tabla a los combobox'

        cmbfvig.DataSource = dt12
        cmbfvig.ValueMember = "Código"
        cmbfvig.DisplayMember = "Modelo"
        cmbfvig.Update()

        ' COMBOBOX FORRO POSTE VERTICAL *******************************************************************************************
        'Creación de la variable tabla para los diferentes forros para postes verticales

        Dim dt13 As DataTable = New DataTable("Forros postes verticales")
        dt13.Columns.Add("Código")
        dt13.Columns.Add("Modelo")
        Dim dr13 As DataRow

        'Añadir los forros para los postes verticales

        dr13 = dt13.NewRow()
        dr13("Modelo") = ""
        dt13.Rows.Add(dr13)

        dr13 = dt13.NewRow()
        dr13("Código") = 2251057
        dr13("Modelo") = "Forro para poste vertical  en aluminio lacado o anodizado de 40/60/80 x 40/60/80 hasta 3000 mm."
        dt13.Rows.Add(dr13)

        dr13 = dt13.NewRow()
        dr13("Código") = 2251073
        dr13("Modelo") = "Forro de acero inox. - 304 para poste vertical de 40/60/80 x 40/60/80, hasta 3000 mm."
        dt13.Rows.Add(dr13)

        dr13 = dt13.NewRow()
        dr13("Código") = 2251074
        dr13("Modelo") = "Forro de acero inox. - 316 para poste vertical de 40/60/80 x 40/60/80, hasta 3000 mm."
        dt13.Rows.Add(dr13)

        ' Asociar los datos de la tabla a los combobox'

        cmbfpv.DataSource = dt13
        cmbfpv.ValueMember = "Código"
        cmbfpv.DisplayMember = "Modelo"
        cmbfpv.Update()

        ' COMBOBOX FORROS TAPAS OPERADORES *************************************************************************************
        'Creación de la variable tabla para los diferentes forros para tapas de operador

        Dim dt14 As DataTable = New DataTable("Forros tapas operador")
        dt14.Columns.Add("Código")
        dt14.Columns.Add("Modelo")
        Dim dr14 As DataRow

        'Añadir los forros para tapas de operador

        dr14 = dt14.NewRow()
        dr14("Modelo") = ""
        dt14.Rows.Add(dr14)

        dr14 = dt14.NewRow()
        dr14("Código") = 2251021
        dr14("Modelo") = "Forro para tapa, de acero inox. - 304 para MD-45 hasta 4000 mm."
        dt14.Rows.Add(dr14)

        dr14 = dt14.NewRow()
        dr14("Código") = 2251033
        dr14("Modelo") = "Forro para tapa, de acero inox. - 316 para MD-45 hasta 4000 mm."
        dt14.Rows.Add(dr14)

        dr14 = dt14.NewRow()
        dr14("Código") = 2251062
        dr14("Modelo") = "Forro para tapa, de acero inox. - 304 para MD-45 hasta 6000 mm."
        dt14.Rows.Add(dr14)

        dr14 = dt14.NewRow()
        dr14("Código") = 2251063
        dr14("Modelo") = "Forro para tapa, de acero inox. - 316 para MD-45 hasta 6000 mm."
        dt14.Rows.Add(dr14)

        dr14 = dt14.NewRow()
        dr14("Código") = 2251064
        dr14("Modelo") = "Forro para tapa, de acero inox. - 304 para MD60/45H hasta 4000 mm."
        dt14.Rows.Add(dr14)

        dr14 = dt14.NewRow()
        dr14("Código") = 2251066
        dr14("Modelo") = "Forro para tapa, de acero inox. - 316 para MD60/45H hasta 4000 mm."
        dt14.Rows.Add(dr14)

        dr14 = dt14.NewRow()
        dr14("Código") = 2251067
        dr14("Modelo") = "Forro para tapa, de acero inox. - 304 para MD60/45H hasta 6000 mm."
        dt14.Rows.Add(dr14)

        dr14 = dt14.NewRow()
        dr14("Código") = 2251068
        dr14("Modelo") = "Forro para tapa, de acero inox. - 316 para MD60/45H hasta 6000 mm."
        dt14.Rows.Add(dr14)

        ' Asociar los datos de la tabla a los combobox'

        cmbftap.DataSource = dt14
        cmbftap.ValueMember = "Código"
        cmbftap.DisplayMember = "Modelo"
        cmbftap.Update()

        ' COMBOBOX FORROS PERFILERÍA *************************************************************************************
        'Creación de la variable tabla para los diferentes forros para las perfilerías

        Dim dt15 As DataTable = New DataTable("Forros perfilería")
        dt15.Columns.Add("Código")
        dt15.Columns.Add("Modelo")
        Dim dr15 As DataRow

        'Añadir los forros para perfilería

        dr15 = dt15.NewRow()
        dr15("Modelo") = ""
        dt15.Rows.Add(dr15)

        dr15 = dt15.NewRow()
        dr15("Código") = 2251058
        dr15("Modelo") = "Forro hoja perfilería MI, en acero inoxidable - 304."
        dt15.Rows.Add(dr15)

        dr15 = dt15.NewRow()
        dr15("Código") = 2251059
        dr15("Modelo") = "Forro hoja perfilería MI, en acero inoxidable - 316."
        dt15.Rows.Add(dr15)

        dr15 = dt15.NewRow()
        dr15("Código") = 2251060
        dr15("Modelo") = "Forro hoja perfilerías BW52/AP94, en acero inoxidable - 304."
        dt15.Rows.Add(dr15)

        dr15 = dt15.NewRow()
        dr15("Código") = 2251061
        dr15("Modelo") = "Forro hoja perfilerías BW52/AP94, en acero inoxidable - 316."
        dt15.Rows.Add(dr15)

        dr15 = dt15.NewRow()
        dr15("Código") = 2251022
        dr15("Modelo") = "Forro Plinto superior, en acero inoxidable - 304."
        dt15.Rows.Add(dr15)

        dr15 = dt15.NewRow()
        dr15("Código") = 2251023
        dr15("Modelo") = "Forro Plinto inferior, en acero inoxidable - 304."
        dt15.Rows.Add(dr15)

        dr15 = dt15.NewRow()
        dr15("Código") = 2251065
        dr15("Modelo") = "Forro para fijo superior, en acero inoxidable - 304."
        dt15.Rows.Add(dr15)

        dr15 = dt15.NewRow()
        dr15("Código") = 2251080
        dr15("Modelo") = "Forro para fijo superior, en acero inoxidable - 316."
        dt15.Rows.Add(dr15)

        ' Asociar los datos de la tabla a los combobox'

        cmbfper.DataSource = dt15
        cmbfper.ValueMember = "Código"
        cmbfper.DisplayMember = "Modelo"
        cmbfper.Update()

        ' COMBO BOX TRATAMIENTO DE PERFILERÍA ***************************************************Creación de la variable tabla para los diferentes tratamientos

        Dim dt16 As DataTable = New DataTable("Tratamiento")
        dt16.Columns.Add("Código")
        dt16.Columns.Add("Modelo")
        Dim dr16 As DataRow

        ' Adición de cada uno de los modelos de operador'

        dr16 = dt16.NewRow()
        dr16("Modelo") = ""
        dt16.Rows.Add(dr16)

        dr16 = dt16.NewRow()
        dr16("Código") = 11111
        dr16("Modelo") = "Grupo 1"
        dt16.Rows.Add(dr16)

        dr16 = dt16.NewRow()
        dr16("Código") = 11112
        dr16("Modelo") = "Grupo 2"
        dt16.Rows.Add(dr16)

        dr16 = dt16.NewRow()
        dr16("Código") = 11113
        dr16("Modelo") = "Grupo 3"
        dt16.Rows.Add(dr16)

        dr16 = dt16.NewRow()
        dr16("Código") = 11114
        dr16("Modelo") = "Grupo 4"
        dt16.Rows.Add(dr16)

        ' Asociar los datos de la tabla a los combobox'

        cmblac.DataSource = dt16
        cmblac.ValueMember = "Código"
        cmblac.DisplayMember = "Modelo"
        cmblac.Update()

        ' COMBOBOX ANODIZADO *******************************************************************************************************
        'Creación de la variable tabla para los diferentes tipos de anodizado

        Dim dt17 As DataTable = New DataTable("Grupos de Anodizados")
        dt17.Columns.Add("Código")
        dt17.Columns.Add("Modelo")
        Dim dr17 As DataRow

        'Añadir los pulsadores'

        dr17 = dt17.NewRow()
        dr17("Modelo") = ""
        dt17.Rows.Add(dr17)

        dr17 = dt17.NewRow()
        dr17("Código") = 11115
        dr17("Modelo") = "Plata 15 micras"
        dt17.Rows.Add(dr17)

        dr17 = dt17.NewRow()
        dr17("Código") = 11116
        dr17("Modelo") = "Plata 20 micras"
        dt17.Rows.Add(dr17)

        dr17 = dt17.NewRow()
        dr17("Código") = 11117
        dr17("Modelo") = "Bronce 15 micras"
        dt17.Rows.Add(dr17)

        dr17 = dt17.NewRow()
        dr17("Código") = 11118
        dr17("Modelo") = "Inox 15 micras"
        dt17.Rows.Add(dr17)

        ' Asociar los datos de la tabla a los combobox'

        cmbano.DataSource = dt17
        cmbano.ValueMember = "Código"
        cmbano.DisplayMember = "Modelo"
        cmbano.Update()

        ' NO VISIBILIDAD DE LOS LABELS QUE MUESTRAN LA INFORMACIÓN EN PRESUPUESTO ******************************************

        lblcod1.Visible = False
        lblcod2.Visible = False
        lbldes3.Visible = False
        lblcod4.Visible = False
        lblcod5.Visible = False
        lblcod6.Visible = False
        lblcod7.Visible = False
        lblcod8.Visible = False
        lblcod9.Visible = False
        lblcod10.Visible = False
        lblcod11.Visible = False
        lblcod12.Visible = False
        lblcod13.Visible = False
        lblcod14.Visible = False
        lblcod15.Visible = False
        lblcod16.Visible = False
        lblcod17.Visible = False
        lblcod18.Visible = False
        lblcod19.Visible = False
        lblcod20.Visible = False
        lblcod21.Visible = False

        lbldes1.Visible = False
        lbldes2.Visible = False
        lbluds3.Visible = False
        lbldes4.Visible = False
        lbldes5.Visible = False
        lbldes6.Visible = False
        lbldes7.Visible = False
        lbldes8.Visible = False
        lbldes9.Visible = False
        lbldes10.Visible = False
        lbldes11.Visible = False
        lbldes12.Visible = False
        lbldes13.Visible = False
        lbldes14.Visible = False
        lbldes15.Visible = False
        lbldes16.Visible = False
        lbldes17.Visible = False
        lbldes18.Visible = False
        lbldes19.Visible = False
        lbldes20.Visible = False
        lbldes21.Visible = False

        lbluds1.Visible = False
        lbluds2.Visible = False
        lblcod3.Visible = False
        lbluds4.Visible = False
        lbluds5.Visible = False
        lbluds6.Visible = False
        lbluds7.Visible = False
        lbluds8.Visible = False
        lbluds9.Visible = False
        lbluds10.Visible = False
        lbluds11.Visible = False
        lbluds12.Visible = False
        lbluds13.Visible = False
        lbluds14.Visible = False
        lbluds15.Visible = False
        lbluds16.Visible = False
        lbluds17.Visible = False
        lbluds18.Visible = False
        lbluds19.Visible = False
        lbluds20.Visible = False
        lbluds21.Visible = False

        lblpre1.Visible = False
        lblpre2.Visible = False
        lblpre3.Visible = False
        lblpre4.Visible = False
        lblpre5.Visible = False
        lblpre6.Visible = False
        lblpre7.Visible = False
        lblpre8.Visible = False
        lblpre9.Visible = False
        lblpre10.Visible = False
        lblpre11.Visible = False
        lblpre12.Visible = False
        lblpre13.Visible = False
        lblpre14.Visible = False
        lblpre15.Visible = False
        lblpre16.Visible = False
        lblpre17.Visible = False
        lblpre18.Visible = False
        lblpre19.Visible = False
        lblpre20.Visible = False
        lblpre21.Visible = False

        lblimp1.Visible = False
        lblimp2.Visible = False
        lblimp3.Visible = False
        lblimp4.Visible = False
        lblimp5.Visible = False
        lblimp6.Visible = False
        lblimp7.Visible = False
        lblimp8.Visible = False
        lblimp9.Visible = False
        lblimp10.Visible = False
        lblimp11.Visible = False
        lblimp12.Visible = False
        lblimp13.Visible = False
        lblimp14.Visible = False
        lblimp15.Visible = False
        lblimp16.Visible = False
        lblimp17.Visible = False
        lblimp18.Visible = False
        lblimp19.Visible = False
        lblimp20.Visible = False
        lblimp21.Visible = False

        Label160.Visible = False
        Label161.Visible = False
        Label162.Visible = False
        Label163.Visible = False

    End Sub

#End Region

    Private Sub tb_pl_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tb_pl.TextChanged

        bFlag = False

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        cargaExcel()

    End Sub

    Private Sub cmbfper_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbfper.SelectionChangeCommitted

        If ((cmbfper.SelectedIndex = 5) Or (cmbfper.SelectedIndex = 6)) Then

            cmbfperu.Visible = False
            tb_metros.Visible = True
            lblmetros1.Visible = True

        Else

            cmbfperu.Visible = True
            tb_metros.Visible = False
            lblmetros1.Visible = False

        End If

    End Sub

    Private Sub cmbotr1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbotr1.SelectionChangeCommitted

        If ((cmbotr1.SelectedIndex = 3)) Then

            cmbotru1.Visible = False
            tb_otr1.Visible = True
            lblmetros2.Visible = True

        Else

            cmbotru1.Visible = True
            tb_otr1.Visible = False
            lblmetros2.Visible = False

        End If

    End Sub

    Private Sub cmbotr2_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbotr2.SelectionChangeCommitted

        If ((cmbotr2.SelectedIndex = 3)) Then

            cmbotru2.Visible = False
            tb_otr2.Visible = True
            lblmetros2.Visible = True

        Else

            cmbotru2.Visible = True
            tb_otr2.Visible = False
            lblmetros2.Visible = False

        End If

    End Sub

    Private Sub cmbmodcar_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbmodcar.SelectionChangeCommitted

        If (cmbmodcar.SelectedIndex = 1) Then

            Panel8.Visible = True

            'Creación de la variable tabla para los diferentes artículos opcionales para perfilería MI'

            Dim dt16 As DataTable = New DataTable("Artículos MI")
            dt16.Columns.Add("Código")
            dt16.Columns.Add("Modelo")
            Dim dr16 As DataRow

            ' Adición de cada uno de los artículos

            dr16 = dt16.NewRow()
            dr16("Código") = ""
            dr16("Modelo") = ""
            dt16.Rows.Add(dr16)

            dr16 = dt16.NewRow()
            dr16("Código") = "2251084"
            dr16("Modelo") = "Perfil de cierre para MI"
            dt16.Rows.Add(dr16)

            ' Asociar la tabla al combobox'

            cmbart1.DataSource = dt16
            cmbart1.ValueMember = "Código"
            cmbart1.DisplayMember = "Modelo"
            cmbart1.Update()

            cmbart2.Visible = False
            cmbart3.Visible = False
            cmbart4.Visible = False
            cmbartu2.Visible = False
            cmbartu3.Visible = False
            cmbartu4.Visible = False
            tb_art2.Visible = False
            tb_art3.Visible = False
            tb_art4.Visible = False

        ElseIf (cmbmodcar.SelectedIndex = 7) Then

            Panel8.Visible = True

            cmbart2.Visible = True
            cmbart3.Visible = True
            cmbart4.Visible = True
            cmbartu2.Visible = True
            cmbartu3.Visible = True
            cmbartu4.Visible = True
            tb_art2.Visible = True
            tb_art3.Visible = True
            tb_art4.Visible = True

            'Creación de la variable tabla para los diferentes artículos opcionales para perfilería BW52'

            Dim dt16 As DataTable = New DataTable("Artículos BW52")
            dt16.Columns.Add("Código")
            dt16.Columns.Add("Modelo")
            Dim dr16 As DataRow

            ' Adición de cada uno de los artículos

            dr16 = dt16.NewRow()
            dr16("Código") = ""
            dr16("Modelo") = ""
            dt16.Rows.Add(dr16)

            dr16 = dt16.NewRow()
            dr16("Código") = "2251040"
            dr16("Modelo") = "Perfil pilastra en aluminio BW52"
            dt16.Rows.Add(dr16)

            dr16 = dt16.NewRow()
            dr16("Código") = "2251044"
            dr16("Modelo") = "Tocho para pilastra y puerta batiente"
            dt16.Rows.Add(dr16)

            dr16 = dt16.NewRow()
            dr16("Código") = "2251027"
            dr16("Modelo") = "Fijo superior en perfilería BW52 en aluminio extrusionado, completamente acabado"
            dt16.Rows.Add(dr16)

            ' Asociar la tabla al combobox'

            cmbart1.DataSource = dt16
            cmbart1.ValueMember = "Código"
            cmbart1.DisplayMember = "Modelo"
            cmbart1.Update()

            'Creación de la variable tabla para los diferentes artículos opcionales para perfilería BW52'

            Dim dt17 As DataTable = New DataTable("Artículos BW52")
            dt17.Columns.Add("Código")
            dt17.Columns.Add("Modelo")
            Dim dr17 As DataRow

            ' Adición de cada uno de los artículos

            dr17 = dt17.NewRow()
            dr17("Código") = ""
            dr17("Modelo") = ""
            dt17.Rows.Add(dr17)

            dr17 = dt17.NewRow()
            dr17("Código") = "2251040"
            dr17("Modelo") = "Perfil pilastra en aluminio BW52"
            dt17.Rows.Add(dr17)

            dr17 = dt17.NewRow()
            dr17("Código") = "2251044"
            dr17("Modelo") = "Tocho para pilastra y puerta batiente"
            dt17.Rows.Add(dr17)

            dr17 = dt17.NewRow()
            dr17("Código") = "2251027"
            dr17("Modelo") = "Fijo superior en perfilería BW52 en aluminio extrusionado, completamente acabado"
            dt17.Rows.Add(dr17)

            ' Asociar la tabla al combobox'

            cmbart2.DataSource = dt17
            cmbart2.ValueMember = "Código"
            cmbart2.DisplayMember = "Modelo"
            cmbart2.Update()

            'Creación de la variable tabla para los diferentes artículos opcionales para perfilería BW52'

            Dim dt18 As DataTable = New DataTable("Artículos BW52")
            dt18.Columns.Add("Código")
            dt18.Columns.Add("Modelo")
            Dim dr18 As DataRow

            ' Adición de cada uno de los artículos

            dr18 = dt18.NewRow()
            dr18("Código") = ""
            dr18("Modelo") = ""
            dt18.Rows.Add(dr18)

            dr18 = dt18.NewRow()
            dr18("Código") = "2251040"
            dr18("Modelo") = "Perfil pilastra en aluminio BW52"
            dt18.Rows.Add(dr18)

            dr18 = dt18.NewRow()
            dr18("Código") = "2251044"
            dr18("Modelo") = "Tocho para pilastra y puerta batiente"
            dt18.Rows.Add(dr18)

            dr18 = dt18.NewRow()
            dr18("Código") = "2251027"
            dr18("Modelo") = "Fijo superior en perfilería BW52 en aluminio extrusionado, completamente acabado"
            dt18.Rows.Add(dr18)

            ' Asociar la tabla al combobox'

            cmbart3.DataSource = dt18
            cmbart3.ValueMember = "Código"
            cmbart3.DisplayMember = "Modelo"
            cmbart3.Update()

            'Creación de la variable tabla para los diferentes artículos opcionales para perfilería BW52'

            Dim dt19 As DataTable = New DataTable("Artículos BW52")
            dt19.Columns.Add("Código")
            dt19.Columns.Add("Modelo")
            Dim dr19 As DataRow

            ' Adición de cada uno de los artículos

            dr19 = dt19.NewRow()
            dr19("Código") = ""
            dr19("Modelo") = ""
            dt19.Rows.Add(dr19)

            dr19 = dt19.NewRow()
            dr19("Código") = "2251040"
            dr19("Modelo") = "Perfil pilastra en aluminio BW52"
            dt19.Rows.Add(dr19)

            dr19 = dt19.NewRow()
            dr19("Código") = "2251044"
            dr19("Modelo") = "Tocho para pilastra y puerta batiente"
            dt19.Rows.Add(dr19)

            dr19 = dt19.NewRow()
            dr19("Código") = "2251027"
            dr19("Modelo") = "Fijo superior en perfilería BW52 en aluminio extrusionado, completamente acabado"
            dt19.Rows.Add(dr19)

            ' Asociar la tabla al combobox'

            cmbart4.DataSource = dt19
            cmbart4.ValueMember = "Código"
            cmbart4.DisplayMember = "Modelo"
            cmbart4.Update()

        End If

    End Sub

    Private Sub cmbart1_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbart1.SelectionChangeCommitted

        If (cmbart1.SelectedIndex = cmbart1.FindStringExact("Perfil pilastra en aluminio BW52") Or (cmbart1.SelectedIndex = cmbart1.FindStringExact("Tocho para pilastra y puerta batiente"))) Then

            tb_art1.Visible = False
            cmbartu1.Visible = True

        ElseIf (cmbart1.SelectedIndex = cmbart1.FindStringExact("Fijo superior en perfilería BW52 en aluminio extrusionado, completamente acabado")) Then

            cmbartu1.Visible = False
            tb_art1.Visible = True

        ElseIf (cmbart1.SelectedIndex = cmbart1.FindStringExact("")) Then

            tb_art1.Visible = True
            cmbartu1.Visible = True

        End If

    End Sub

    Private Sub cmbart2_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbart2.SelectionChangeCommitted

        If (cmbart2.SelectedIndex = cmbart2.FindStringExact("Perfil pilastra en aluminio BW52") Or (cmbart2.SelectedIndex = cmbart2.FindStringExact("Tocho para pilastra y puerta batiente"))) Then

            tb_art2.Visible = False
            cmbartu2.Visible = True

        ElseIf (cmbart2.SelectedIndex = cmbart2.FindStringExact("Fijo superior en perfilería BW52 en aluminio extrusionado, completamente acabado")) Then

            cmbartu2.Visible = False
            tb_art2.Visible = True

        ElseIf (cmbart2.SelectedIndex = cmbart2.FindStringExact("")) Then

            tb_art2.Visible = True
            cmbartu2.Visible = True

        End If

    End Sub

    Private Sub cmbart3_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbart3.SelectionChangeCommitted

        If (cmbart3.SelectedIndex = cmbart3.FindStringExact("Perfil pilastra en aluminio BW52") Or (cmbart3.SelectedIndex = cmbart3.FindStringExact("Tocho para pilastra y puerta batiente"))) Then

            tb_art3.Visible = False
            cmbartu3.Visible = True

        ElseIf (cmbart3.SelectedIndex = cmbart3.FindStringExact("Fijo superior en perfilería BW52 en aluminio extrusionado, completamente acabado")) Then

            cmbartu3.Visible = False
            tb_art3.Visible = True

        ElseIf (cmbart3.SelectedIndex = cmbart3.FindStringExact("")) Then

            tb_art3.Visible = True
            cmbartu3.Visible = True

        End If

    End Sub

    Private Sub cmbart4_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbart4.SelectionChangeCommitted

        If (cmbart4.SelectedIndex = cmbart4.FindStringExact("Perfil pilastra en aluminio BW52") Or (cmbart4.SelectedIndex = cmbart4.FindStringExact("Tocho para pilastra y puerta batiente"))) Then

            tb_art4.Visible = False
            cmbartu4.Visible = True

        ElseIf (cmbart4.SelectedIndex = cmbart4.FindStringExact("Fijo superior en perfilería BW52 en aluminio extrusionado, completamente acabado")) Then

            cmbartu4.Visible = False
            tb_art4.Visible = True

        ElseIf (cmbart4.SelectedIndex = cmbart4.FindStringExact("")) Then

            tb_art4.Visible = True
            cmbartu4.Visible = True

        End If

    End Sub

    Private Sub calcular_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles calcular.Click

        tb_nombre3.Text = tb_nombre.Text
        tb_dni3.Text = tb_dni.Text
        tb_tlf3.Text = tb_tlf.Text
        tb_cif3.Text = tb_cif.Text
        tb_direccion3.Text = tb_direccion.Text
        tb_email3.Text = tb_email.Text

        carpinteria()

    End Sub

    Private Sub btnimprimir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnimprimir.Click

        Button3.Visible = False
        btnimprimir.Visible = False

        imprimir()

        Button3.Visible = True
        btnimprimir.Visible = True

    End Sub

    Private Sub btnimprimir2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnimprimir2.Click

        calcular.Visible = False
        btnimprimir2.Visible = False

        imprimir()

        calcular.Visible = True
        btnimprimir2.Visible = True

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If Me.CheckBox1.Checked = True Then
            Panel9.Visible = True
        Else
            Panel9.Visible = False
        End If
    End Sub

    Private Sub carpinteria()

        ' Variable Flag en Falso al pulsar calcular para marcar el final de los bucles while ******************************************

        bFlag = False

        ' Bucles While para indicar las configuraciones que no son posibles ***********************************************************

        While ((cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI50 D. Corredera Doble. 2 Hojas, 180 Kg máx. Incluye batería y selector de 4 posiciones") Or cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI50 S. Corredera Simple. 1 Hoja, 110 Kg máx. Incluye batería y selector de 4 posiciones") Or cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI14 D. Corredera Doble. 2 Hojas, 140 Kg máx. Incluye batería y selector de 4 posiciones") Or cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI14 S. Corredera Simple. 1 Hoja, 100 Kg máx. Incluye batería y selector de 4 posiciones") Or cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI75 D. Corredera Doble. 2 Hojas, 250 Kg máx. Incluye batería y selector de 4 posiciones") Or cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI75 S. Corredera Simple. 1 Hoja, 180 Kg máx. Incluye batería y selector de 4 posiciones") Or cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI100 D. Corredera Doble. 2 Hojas, 350 Kg máx. Incluye batería y selector de 4 posiciones") Or cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI100 S. Corredera Simple. 1 Hojas, 250 Kg máx. Incluye batería y selector de 4 posiciones")) And bFlag = False And (cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 4 Hojas Móviles") Or cmbconf.SelectedIndex = cmbconf.FindStringExact("4 Hojas Móviles") Or cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 2 Hojas Móviles") Or cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Móviles TES")))

            bFlag = True

            MsgBox("La configuración de hojas elegida no es compatible para el modelo de operador seleccionado", MsgBoxStyle.Exclamation)

        End While

        While ((cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI50 D. Corredera Doble. 2 Hojas, 180 Kg máx. Incluye batería y selector de 4 posiciones") Or cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI14 D. Corredera Doble. 2 Hojas, 140 Kg máx. Incluye batería y selector de 4 posiciones") Or cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI75 D. Corredera Doble. 2 Hojas, 250 Kg máx. Incluye batería y selector de 4 posiciones") Or cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI100 D. Corredera Doble. 2 Hojas, 350 Kg máx. Incluye batería y selector de 4 posiciones")) And bFlag = False And (cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 1 Hoja Móvil") Or cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Móvil")))

            bFlag = True

            MsgBox("La configuración de hojas elegida no es compatible para el modelo de operador seleccionado", MsgBoxStyle.Exclamation)

        End While

        While ((cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI50 S. Corredera Simple. 1 Hoja, 110 Kg máx. Incluye batería y selector de 4 posiciones") Or cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI14 S. Corredera Simple. 1 Hoja, 100 Kg máx. Incluye batería y selector de 4 posiciones") Or cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI75 S. Corredera Simple. 1 Hoja, 180 Kg máx. Incluye batería y selector de 4 posiciones") Or cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI100 S. Corredera Simple. 1 Hojas, 250 Kg máx. Incluye batería y selector de 4 posiciones")) And bFlag = False And (cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 4 Hojas Móviles") Or cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 2 Hojas Móviles") Or cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Móviles")))

            bFlag = True

            MsgBox("La configuración de hojas elegida no es compatible para el modelo de operador seleccionado", MsgBoxStyle.Exclamation)

        End While

        While ((cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI50 TES TH4. Telescópica. 4 Hojas, 45 x 4 Kg máx. Incluye batería y selector de 4 posiciones") Or cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI50 TES TH2. Telescópica. 2 Hojas, 45 x 2 Kg máx. Incluye batería y selector de 4 posiciones")) And bFlag = False And (cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 2 Hojas Móviles") Or cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Móviles") Or cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 1 Hoja Móvil") Or cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hojas Móvil")))

            bFlag = True

            MsgBox("La configuración de hojas elegida no es compatible para el modelo de operador seleccionado", MsgBoxStyle.Exclamation)

        End While

        While ((cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI50 TES TH4. Telescópica. 4 Hojas, 45 x 4 Kg máx. Incluye batería y selector de 4 posiciones") Or cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI50 TES TH2. Telescópica. 2 Hojas, 45 x 2 Kg máx. Incluye batería y selector de 4 posiciones")) And (bFlag = False) And (cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto de pinza. Sólo superior") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto silicona. Sólo superior")))

            bFlag = True

            MsgBox("La carpintería elegida no es compatible para el modelo de operador seleccionado", MsgBoxStyle.Exclamation)

        End While

        While ((cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI50 TES TH4. Telescópica. 4 Hojas, 45 x 4 Kg máx. Incluye batería y selector de 4 posiciones") And bFlag = False And (cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 2 Hojas Móviles") Or cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Móviles TES"))))

            bFlag = True

            MsgBox("La configuración de hojas elegida no es compatible para el modelo de operador seleccionado", MsgBoxStyle.Exclamation)

        End While

        While (cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI50 TES TH2. Telescópica. 2 Hojas, 45 x 2 Kg máx. Incluye batería y selector de 4 posiciones") And bFlag = False And (cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 4 Hojas Móviles") Or cmbconf.SelectedIndex = cmbconf.FindStringExact("4 Hojas Móviles")))

            bFlag = True

            MsgBox("La configuración de hojas elegida no es compatible para el modelo de operador seleccionado", MsgBoxStyle.Exclamation)

        End While

        'While ((cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI50 TES TH4. Telescópica. 4 Hojas, 45 x 4 Kg máx. Incluye batería y selector de 4 posiciones") Or cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI50 TES TH2. Telescópica. 2 Hojas, 45 x 2 Kg máx. Incluye batería y selector de 4 posiciones")) And bFlag = False And cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("AP94"))

        '    bFlag = True

        '    MsgBox("Los operadores telescópicos no admiten carpintería antipánico", MsgBoxStyle.Exclamation)

        'End While

         While ((cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI50 TES TH4. Telescópica. 4 Hojas, 45 x 4 Kg máx. Incluye batería y selector de 4 posiciones") Or cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI50 TES TH2. Telescópica. 2 Hojas, 45 x 2 Kg máx. Incluye batería y selector de 4 posiciones")) And bFlag = False And cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto silicona. Sólo superior"))

            bFlag = True

            MsgBox("Los operadores telescópicos no admiten carpintería plintos solo superior", MsgBoxStyle.Exclamation)

        End While

        While ((cmbmodope.SelectedIndex = cmbmodope.FindStringExact("")) And (cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto silicona. Sólo superior") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto de pinza. Sólo superior")) And bFlag = False And (cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 4 Hojas Móviles") Or cmbconf.SelectedIndex = cmbconf.FindStringExact("4 Hojas Móviles") Or cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 2 Hojas Móviles") Or cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Móviles TES")))

            bFlag = True

            MsgBox("La configuración de hojas elegida no es compatible para la carpintería seleccionada", MsgBoxStyle.Exclamation)

        End While

        While (tb_pl.Text = Nothing And bFlag = False And (cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Móviles") Or cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Móvil")))

            bFlag = True

            MsgBox("El Paso Libre es un dato obligatorio cuando sólo existen hojas móviles", MsgBoxStyle.Exclamation)

        End While

        While ((cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI SW Push. Operador para puertas batientes con brazo articulado") Or cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI SW Pull. Operador para puertas batientes con brazo deslizante") Or cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI SWSP Push. Operador para puertas batientes con brazo articulado y cierre por muelle") Or cmbmodope.SelectedIndex = cmbmodope.FindStringExact("MI SWSP Pull. Operador para puertas batientes con brazo deslizante y cierre por muelle")) And bFlag = False And (cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("MI") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Full Glass") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plintos silicona superior e inferior") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto silicona. Sólo superior") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto superior de pinza e inferior de silicona") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto de pinza. Sólo superior") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("BW52")))

            bFlag = True

            MsgBox("Los operadores batientes se instalan sin las carpinterías de Master Ingenieros", MsgBoxStyle.Exclamation)

        End While

        While ((cmbconf.SelectedIndex = 2) Or (cmbconf.SelectedIndex = 4) Or (cmbconf.SelectedIndex = 6) Or (cmbconf.SelectedIndex = 8)) And (tb_pl.Text = "") And bFlag = False

            bFlag = True

            MsgBox("Introduzca un valor para el paso libre", MsgBoxStyle.Exclamation)

        End While


        ' Bucle While que se ejecuta siempre que el paso libre no sea un dato ****************************************

        While (tb_pl.Text = "" And bFlag = False And (cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 2 Hojas Móviles") Or cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 1 Hoja Móvil") Or cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 4 Hojas Móviles") Or cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 2 Hojas Móviles")))

            bFlag = True

            '******************** Inicio del condicional para calcular medidas y perfilerías ******************************

            ' Al modificar el valor del Text Box relativo al Paso Libre se inicializa el condicional 
            ' La función Val hace que las variables definidas adopten el contenido numérico del Text Box

            wh = Val(tb_wh.Text)
            hh = Val(tb_hh.Text)
            pl = Val(tb_pl.Text)

            '****************** MI - CASO I *****************************************************************************************************************************

            If cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("MI") And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 2 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías MI con 2 hojas fijas y 2 móviles

                AHF = (wh / 4) + (solMI / 2)
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías MI con 2 hojas fijas y 2 móviles

                HHF = hh - 10
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías MI con 2 hojas fijas y 2 móviles

                AHM = AHF
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías MI con 2 hojas fijas y 2 móviles

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías MI con 2 hojas fijas y 2 móviles

                AVF = AHF - deshMI
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías MI con 2 hojas fijas y 2 móviles

                HVF = HHF - desvMI
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías MI con 2 hojas fijas y 2 móviles

                AVM = AHM - deshMI
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías MI con 2 hojas fijas y 2 móviles

                HVM = HHM - desvMI - 16
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calculo las medidas de los componentes de la carpintería MI

                a1 = HHF
                a2 = AHF - 90
                a3 = HHF + 8
                a4 = HHM - 16
                a5 = AHM - 90
                a6 = HHM - 16
                a7 = AHM
                a8 = wh - (2 * AHF)
                a9 = wh
                a10 = a9

                a1text.Text = a1.ToString()
                a1text.Update()
                a2text.Text = a2.ToString()
                a2text.Update()
                a3text.Text = a3.ToString()
                a3text.Update()
                a4text.Text = a4.ToString()
                a4text.Update()
                a5text.Text = a5.ToString()
                a5text.Update()
                a6text.Text = a6.ToString()
                a6text.Update()
                a7text.Text = a7.ToString()
                a7text.Update()
                a8text.Text = a8.ToString()
                a8text.Update()
                a9text.Text = a9.ToString()
                a9text.Update()
                a10text.Text = a10.ToString()
                a10text.Update()

                au1 = 4
                au2 = 4
                au3 = 2
                au4 = 4
                au5 = 4
                au6 = 2
                au7 = 2
                au8 = 1
                au9 = 1
                au10 = 1

                au1text.Text = au1.ToString()
                au1text.Update()
                au2text.Text = au2.ToString()
                au2text.Update()
                au3text.Text = au3.ToString()
                au3text.Update()
                au4text.Text = au4.ToString()
                au4text.Update()
                au5text.Text = au5.ToString()
                au5text.Update()
                au6text.Text = au6.ToString()
                au6text.Update()
                au7text.Text = au7.ToString()
                au7text.Update()
                au8text.Text = au8.ToString()
                au8text.Update()
                au9text.Text = au9.ToString()
                au9text.Update()
                au10text.Text = au10.ToString()
                au10text.Update()

                Panel3.Visible = True
                Panel4.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                au17text.Visible = False
                a17text.Visible = False
                Label223.Visible = False
                TabControl2.SelectedIndex = 0

                '****************** MI - CASO III ****************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("MI") And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 1 Hoja Móvil") Then

                ' Calcula el ancho de la hoja fija para carpinterías MI con 1 hoja fija y 1 móvil

                AHF = (wh / 2) + (solMI / 2)
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías MI con 1 hoja fija y 1 móvil

                HHF = hh - 10
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías MI con 1 hoja fija y 1 móvil

                AHM = (wh / 2) + (solMI / 2)
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías MI con 1 hoja fija y 1 móvil

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías MI con 1 hoja fija y 1 móvil

                AVF = AHF - deshMI
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías MI con 1 hoja fija y 1 móvil

                HVF = HHF - desvMI
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías MI con 1 hoja fija y 1 móvil

                AVM = AHM - deshMI
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías MI con 1 hoja fija y 1 móvil

                HVM = HHM - desvMI - 16
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calculo las medidas de los componentes de la carpintería MI

                a1 = HHF
                a2 = AHF - 90
                a3 = HHF + 8
                a4 = HHM - 16
                a5 = AHM - 90
                a6 = HHM - 16
                a7 = AHM
                a8 = wh - AHF
                a9 = wh
                a10 = a9

                a1text.Text = a1.ToString()
                a1text.Update()
                a2text.Text = a2.ToString()
                a2text.Update()
                a3text.Text = a3.ToString()
                a3text.Update()
                a4text.Text = a4.ToString()
                a4text.Update()
                a5text.Text = a5.ToString()
                a5text.Update()
                a6text.Text = a6.ToString()
                a6text.Update()
                a7text.Text = a7.ToString()
                a7text.Update()
                a7text.Text = a7.ToString()
                a7text.Update()
                a8text.Text = a8.ToString()
                a8text.Update()
                a9text.Text = a9.ToString()
                a9text.Update()
                a10text.Text = a10.ToString()
                a10text.Update()

                au1 = 2
                au2 = 2
                au3 = 1
                au4 = 2
                au5 = 2
                au6 = 1
                au7 = 1
                au8 = 1
                au9 = 1
                au10 = 1

                au1text.Text = au1.ToString()
                au1text.Update()
                au2text.Text = au2.ToString()
                au2text.Update()
                au3text.Text = au3.ToString()
                au3text.Update()
                au4text.Text = au4.ToString()
                au4text.Update()
                au5text.Text = au5.ToString()
                au5text.Update()
                au6text.Text = au6.ToString()
                au6text.Update()
                au7text.Text = au7.ToString()
                au7text.Update()
                au8text.Text = au8.ToString()
                au8text.Update()
                au9text.Text = au9.ToString()
                au9text.Update()
                au10text.Text = au10.ToString()
                au10text.Update()

                Panel3.Visible = True
                Panel4.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                au17text.Visible = False
                a17text.Visible = False
                Label223.Visible = False
                TabControl2.SelectedIndex = 0

                '****************** MI - CASO V ****************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("MI") And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 4 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                AHF = ((wh + 200) / 6)
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                HHF = hh - 10
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                AHM = ((wh + 200) / 6)
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                HHM = hh - 5
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho de la hoja móvil lenta para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                AHML = ((wh + 200) / 6)
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                ' Calcula la altura de la hoja móvil lenta para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                HHML = hh
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                AVF = AHF - deshMI
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                HVF = HHF - desvMI
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                AVM = AHM - deshMI
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                HVM = HHM - desvMI - 16
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil lenta para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                AVML = AHML - deshMI - 10
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil lenta para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                HVML = HHML - desvMI - 16
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Calculo las medidas de los componentes de la carpintería MI

                a1 = HHF
                a2 = AHF - 90
                a3 = HHF + 8
                a4 = HHM - 16
                a5 = AHM - 90
                a6 = HHM - 16
                a7 = AHM
                a8 = wh - (2 * AHF)
                a9 = wh
                a10 = a9
                a11 = HHML - 16
                a12 = AHML - 100
                a13 = HHML - 16
                a14 = AHML
                a15 = HHML - 16

                a1text.Text = a1.ToString()
                a1text.Update()
                a2text.Text = a2.ToString()
                a2text.Update()
                a3text.Text = "No lleva portafelpudos lateral"
                a3text.Update()
                a4text.Text = a4.ToString()
                a4text.Update()
                a5text.Text = a5.ToString()
                a5text.Update()
                a6text.Text = a6.ToString()
                a6text.Update()
                a7text.Text = a7.ToString()
                a7text.Update()
                a8text.Text = a8.ToString()
                a8text.Update()
                a9text.Text = a9.ToString()
                a9text.Update()
                a10text.Text = a10.ToString()
                a10text.Update()
                a11text.Text = a11.ToString()
                a11text.Update()
                a12text.Text = a12.ToString()
                a12text.Update()
                a13text.Text = a13.ToString()
                a13text.Update()
                a14text.Text = a14.ToString()
                a14text.Update()
                a15text.Text = a15.ToString()
                a15text.Update()

                au1 = 4
                au2 = 4
                'au3 = 2
                au4 = 4
                au5 = 4
                au6 = 2
                au7 = 2
                au8 = 1
                au9 = 1
                au10 = 1
                au11 = 4
                au12 = 4
                au13 = 2
                au14 = 2
                au15 = 2

                au1text.Text = au1.ToString()
                au1text.Update()
                au2text.Text = au2.ToString()
                au2text.Update()
                au3text.Text = "No lleva portafelpudos lateral"
                au3text.Update()
                au4text.Text = au4.ToString()
                au4text.Update()
                au5text.Text = au5.ToString()
                au5text.Update()
                au6text.Text = au6.ToString()
                au6text.Update()
                au7text.Text = au7.ToString()
                au7text.Update()
                au8text.Text = au8.ToString()
                au8text.Update()
                au9text.Text = au9.ToString()
                au9text.Update()
                au10text.Text = au10.ToString()
                au10text.Update()
                au11text.Text = au11.ToString()
                au11text.Update()
                au12text.Text = au12.ToString()
                au12text.Update()
                au13text.Text = au13.ToString()
                au13text.Update()
                au14text.Text = au14.ToString()
                au14text.Update()
                au15text.Text = au15.ToString()
                au15text.Update()

                Panel3.Visible = True
                Panel4.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Panel11.Visible = True
                Label189.Visible = True
                TabControl2.SelectedIndex = 0

                '****************** MI - CASO VII *************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("MI") And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 2 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                AHF = ((wh + 100) / 3)
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                HHF = hh - 10
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                AHM = ((wh + 100) / 3)
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                HHM = hh - 5
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho de la hoja móvil lenta para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                AHML = ((wh + 100) / 3)
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                ' Calcula la altura de la hoja móvil lenta para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                HHML = hh
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                AVF = AHF - deshMI
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                HVF = HHF - desvMI
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                AVM = AHM - deshMI
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                HVM = HHM - desvMI - 21
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil lenta para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                AVML = AHML - deshMI - 10
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil lenta para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                HVML = HHML - desvMI - 16
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Calculo las medidas de los componentes de la carpintería MI

                a1 = HHF
                a2 = AHF - 90
                a3 = HHF + 8
                a4 = HHM - 16
                a5 = AHM - 90
                a6 = HHM - 16
                a7 = AHM
                a8 = wh - (AHF)
                a9 = wh
                a10 = a9
                a11 = HHML - 16
                a12 = AHML - 100
                a13 = HHML - 16
                a14 = AHML
                a15 = HHML - 16

                a1text.Text = a1.ToString()
                a1text.Update()
                a2text.Text = a2.ToString()
                a2text.Update()
                a3text.Text = "No lleva portafelpudos lateral"
                a3text.Update()
                a4text.Text = a4.ToString()
                a4text.Update()
                a5text.Text = a5.ToString()
                a5text.Update()
                a6text.Text = a6.ToString()
                a6text.Update()
                a7text.Text = a7.ToString()
                a7text.Update()
                a8text.Text = a8.ToString()
                a8text.Update()
                a9text.Text = a9.ToString()
                a9text.Update()
                a10text.Text = a10.ToString()
                a10text.Update()
                a11text.Text = a11.ToString()
                a11text.Update()
                a12text.Text = a12.ToString()
                a12text.Update()
                a13text.Text = a13.ToString()
                a13text.Update()
                a14text.Text = a14.ToString()
                a14text.Update()
                a15text.Text = a15.ToString()
                a15text.Update()

                au1 = 2
                au2 = 2
                au3 = 1
                au4 = 2
                au5 = 2
                au6 = 1
                au7 = 1
                au8 = 1
                au9 = 1
                au10 = 1
                au11 = 2
                au12 = 2
                au13 = 1
                au14 = 1
                au15 = 1

                au1text.Text = au1.ToString()
                au1text.Update()
                au2text.Text = au2.ToString()
                au2text.Update()
                au3text.Text = "No lleva portafelpudos lateral"
                au3text.Update()
                au4text.Text = au4.ToString()
                au4text.Update()
                au5text.Text = au5.ToString()
                au5text.Update()
                au6text.Text = au6.ToString()
                au6text.Update()
                au7text.Text = au7.ToString()
                au7text.Update()
                au8text.Text = au8.ToString()
                au8text.Update()
                au9text.Text = au9.ToString()
                au9text.Update()
                au10text.Text = au10.ToString()
                au10text.Update()
                au11text.Text = au11.ToString()
                au11text.Update()
                au12text.Text = au12.ToString()
                au12text.Update()
                au13text.Text = au13.ToString()
                au13text.Update()
                au14text.Text = au14.ToString()
                au14text.Update()
                au15text.Text = au15.ToString()
                au15text.Update()

                Panel3.Visible = True
                Panel4.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Panel11.Visible = True
                Label189.Visible = True
                TabControl2.SelectedIndex = 0

                '****************** FG - CASO I *****************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Full Glass") And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 2 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías FG con 2 hojas fijas y 2 móviles

                AHF = (wh / 4) + (solFG / 2)
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías FG con 2 hojas fijas y 2 móviles

                HHF = hh - 10
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías FG con 2 hojas fijas y 2 móviles

                AHM = (wh / 4) + (solFG / 2)
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías FG con 2 hojas fijas y 2 móviles

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías FG con 2 hojas fijas y 2 móviles

                AVF = AHF - deshFG
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías FG con 2 hojas fijas y 2 móviles

                HVF = HHF - desvFG
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías FG con 2 hojas fijas y 2 móviles

                AVM = AHM - deshFG
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías FG con 2 hojas fijas y 2 móviles

                HVM = HHM - desvFG
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calculo las medidas de los componentes de la carpintería FG

                b1 = HHF
                b2 = HHF
                b3 = AHF - 46
                b4 = HHM
                b5 = HHM
                b6 = AHM - 46
                b7 = AHM - 46
                b8 = wh - (2 * AHF)
                b9 = wh
                b10 = wh
                b11 = b10

                b1text.Text = b1.ToString()
                b1text.Update()
                b2text.Text = b2.ToString()
                b2text.Update()
                b3text.Text = b3.ToString()
                b3text.Update()
                b4text.Text = b4.ToString()
                b4text.Update()
                b5text.Text = b5.ToString()
                b5text.Update()
                b6text.Text = b6.ToString()
                b6text.Update()
                b7text.Text = b7.ToString()
                b7text.Update()
                b8text.Text = b8.ToString()
                b8text.Update()
                b9text.Text = b9.ToString()
                b9text.Update()
                b10text.Text = b10.ToString()
                b10text.Update()
                b11text.Text = b11.ToString()
                b11text.Update()

                bu1 = 2
                bu2 = 2
                bu3 = 4
                bu4 = 2
                bu5 = 2
                bu6 = 2
                bu7 = 2
                bu8 = 1
                bu9 = 1
                bu10 = 1
                bu11 = 1

                bu1text.Text = bu1.ToString()
                bu1text.Update()
                bu2text.Text = bu2.ToString()
                bu2text.Update()
                bu3text.Text = bu3.ToString()
                bu3text.Update()
                bu4text.Text = bu4.ToString()
                bu4text.Update()
                bu5text.Text = bu5.ToString()
                bu5text.Update()
                bu6text.Text = bu6.ToString()
                bu6text.Update()
                bu7text.Text = bu7.ToString()
                bu7text.Update()
                bu8text.Text = bu8.ToString()
                bu8text.Update()
                bu9text.Text = bu9.ToString()
                bu9text.Update()
                bu10text.Text = bu10.ToString()
                bu10text.Update()
                bu11text.Text = bu11.ToString()
                bu11text.Update()

                Panel4.Visible = True
                Panel3.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                TabControl2.SelectedIndex = 1

                '****************** FG - CASO III ****************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Full Glass") And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 1 Hoja Móvil") Then

                ' Calcula el ancho de la hoja fija para carpinterías FG con 1 hoja fija y 1 móvil

                AHF = (wh / 2) + (solFG / 2)
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías FG con 1 hoja fija y 1 móvil

                HHF = hh - 10
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías FG con 1 hoja fija y 1 móvil

                AHM = (wh / 2) + (solFG / 2)
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías FG con 1 hoja fija y 1 móvil

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías FG con 1 hoja fija y 1 móvil

                AVF = AHF - deshFG
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías FG con 1 hoja fija y 1 móvil

                HVF = HHF - desvFG
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías FG con 1 hoja fija y 1 móvil

                AVM = AHM - deshFG
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías FG con 1 hoja fija y 1 móvil

                HVM = HHM - desvFG
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calculo las medidas de los componentes de la carpintería FG

                b1 = HHF
                b2 = HHF
                b3 = AHF - 46
                b4 = HHM
                b5 = HHM
                b6 = AHM - 46
                b7 = AHM - 46
                b8 = wh - AHF
                b9 = wh
                b10 = wh
                b11 = b10

                b1text.Text = b1.ToString()
                b1text.Update()
                b2text.Text = b2.ToString()
                b2text.Update()
                b3text.Text = b3.ToString()
                b3text.Update()
                b4text.Text = b4.ToString()
                b4text.Update()
                b5text.Text = b5.ToString()
                b5text.Update()
                b6text.Text = b6.ToString()
                b6text.Update()
                b7text.Text = b7.ToString()
                b7text.Update()
                b8text.Text = b8.ToString()
                b8text.Update()
                b9text.Text = b9.ToString()
                b9text.Update()
                b10text.Text = b10.ToString()
                b10text.Update()
                b11text.Text = b11.ToString()
                b11text.Update()

                bu1 = 1
                bu2 = 1
                bu3 = 2
                bu4 = 1
                bu5 = 1
                bu6 = 1
                bu7 = 1
                bu8 = 1
                bu9 = 1
                bu10 = 1
                bu11 = 1

                bu1text.Text = bu1.ToString()
                bu1text.Update()
                bu2text.Text = bu2.ToString()
                bu2text.Update()
                bu3text.Text = bu3.ToString()
                bu3text.Update()
                bu4text.Text = bu4.ToString()
                bu4text.Update()
                bu5text.Text = bu5.ToString()
                bu5text.Update()
                bu6text.Text = bu6.ToString()
                bu6text.Update()
                bu7text.Text = bu7.ToString()
                bu7text.Update()
                bu8text.Text = bu8.ToString()
                bu8text.Update()
                bu9text.Text = bu9.ToString()
                bu9text.Update()
                bu10text.Text = bu10.ToString()
                bu10text.Update()
                bu11text.Text = bu11.ToString()
                bu11text.Update()

                Panel4.Visible = True
                Panel3.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                TabControl2.SelectedIndex = 1

                '****************** FG - CASO V ****************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Full Glass") And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 4 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                AHF = ((wh + 120) / 6)
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                HHF = hh - 10
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                AHM = ((wh + 120) / 6)
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                HHM = hh - 15
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho de la hoja móvil lenta para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                AHML = ((wh + 120) / 6)
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                ' Calcula la altura de la hoja móvil lenta para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                HHML = hh - 10
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                AVF = AHF - deshFG
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                HVF = HHF - desvFG
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                AVM = AHM - deshFG
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                HVM = HHM - desvFG - 5
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil lenta para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                AVML = AHML - ((3 * deshFG) / 2)
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil lenta para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                HVML = HHML - desvFG
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Calculo las medidas de los componentes de la carpintería FG

                b1 = HHF
                b2 = HHF
                b3 = AHF - 46
                b4 = HHM
                b5 = HHM
                b6 = AHM - 46
                b7 = AHM - 46
                b8 = wh - (2 * AHF)
                b9 = wh
                b10 = wh
                b11 = b10
                b12 = HHML
                b13 = HHML
                b14 = AHML - 46
                b15 = AHML - 46

                b1text.Text = b1.ToString()
                b1text.Update()
                b2text.Text = b2.ToString()
                b2text.Update()
                b3text.Text = b3.ToString()
                b3text.Update()
                b4text.Text = b4.ToString()
                b4text.Update()
                b5text.Text = b5.ToString()
                b5text.Update()
                b6text.Text = b6.ToString()
                b6text.Update()
                b7text.Text = b7.ToString()
                b7text.Update()
                b8text.Text = b8.ToString()
                b8text.Update()
                b9text.Text = b9.ToString()
                b9text.Update()
                b10text.Text = b10.ToString()
                b10text.Update()
                b11text.Text = b11.ToString()
                b11text.Update()
                b12text.Text = b12.ToString()
                b12text.Update()
                b13text.Text = "La hoja lenta no lleva este perfil"
                b13text.Update()
                b14text.Text = b14.ToString()
                b14text.Update()
                b15text.Text = b15.ToString()
                b15text.Update()

                bu1 = 2
                bu2 = 2
                bu3 = 4
                bu4 = 2
                bu5 = 2
                bu6 = 2
                bu7 = 2
                bu8 = 1
                bu9 = 1
                bu10 = 1
                bu11 = 1
                bu12 = 4
                bu13 = 1
                bu14 = 2
                bu15 = 2

                bu1text.Text = bu1.ToString()
                bu1text.Update()
                bu2text.Text = bu2.ToString()
                bu2text.Update()
                bu3text.Text = bu3.ToString()
                bu3text.Update()
                bu4text.Text = bu4.ToString()
                bu4text.Update()
                bu5text.Text = bu5.ToString()
                bu5text.Update()
                bu6text.Text = bu6.ToString()
                bu6text.Update()
                bu7text.Text = bu7.ToString()
                bu7text.Update()
                bu8text.Text = bu8.ToString()
                bu8text.Update()
                bu9text.Text = bu9.ToString()
                bu9text.Update()
                bu10text.Text = bu10.ToString()
                bu10text.Update()
                bu11text.Text = bu11.ToString()
                bu11text.Update()
                bu12text.Text = bu12.ToString()
                bu12text.Update()
                bu13text.Text = "La hoja lenta no lleva este perfil"
                bu13text.Update()
                bu14text.Text = bu14.ToString()
                bu14text.Update()
                bu15text.Text = bu15.ToString()
                bu15text.Update()

                Panel4.Visible = True
                Panel3.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Label198.Visible = True
                Panel12.Visible = True
                TabControl2.SelectedIndex = 1

                '****************** FG - CASO VII *************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Full Glass") And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 2 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                AHF = ((wh + 60) / 3)
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                HHF = hh - 10
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                AHM = ((wh + 60) / 3)
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                HHM = hh - 15
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho de la hoja móvil lenta para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                AHML = ((wh + 60) / 3)
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                ' Calcula la altura de la hoja móvil lenta para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                HHML = hh - 10
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                AVF = AHF - deshFG
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                HVF = HHF - desvFG
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                AVM = AHM - deshFG
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                HVM = HHM - desvFG - 5
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil lenta para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                AVML = AHML - ((3 * deshFG) / 2)
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil lenta para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                HVML = HHML - desvFG
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Calculo las medidas de los componentes de la carpintería FG

                b1 = HHF
                b2 = HHF
                b3 = AHF - 46
                b4 = HHM
                b5 = HHM
                b6 = AHM - 46
                b7 = AHM - 46
                b8 = wh - (AHF)
                b9 = wh
                b10 = wh
                b11 = b10
                b12 = HHML
                b13 = HHML
                b14 = AHML - 46
                b15 = AHML - 46

                b1text.Text = b1.ToString()
                b1text.Update()
                b2text.Text = b2.ToString()
                b2text.Update()
                b3text.Text = b3.ToString()
                b3text.Update()
                b4text.Text = b4.ToString()
                b4text.Update()
                b5text.Text = b5.ToString()
                b5text.Update()
                b6text.Text = b6.ToString()
                b6text.Update()
                b7text.Text = b7.ToString()
                b7text.Update()
                b8text.Text = b8.ToString()
                b8text.Update()
                b9text.Text = b9.ToString()
                b9text.Update()
                b10text.Text = b10.ToString()
                b10text.Update()
                b11text.Text = b11.ToString()
                b11text.Update()
                b12text.Text = b12.ToString()
                b12text.Update()
                b13text.Text = "La hoja lenta no lleva este perfil"
                b13text.Update()
                b14text.Text = b14.ToString()
                b14text.Update()
                b15text.Text = b15.ToString()
                b15text.Update()

                bu1 = 1
                bu2 = 1
                bu3 = 2
                bu4 = 1
                bu5 = 1
                bu6 = 1
                bu7 = 1
                bu8 = 1
                bu9 = 1
                bu10 = 1
                bu11 = 1
                bu12 = 2
                bu13 = 1
                bu14 = 1
                bu15 = 1

                bu1text.Text = bu1.ToString()
                bu1text.Update()
                bu2text.Text = bu2.ToString()
                bu2text.Update()
                bu3text.Text = bu3.ToString()
                bu3text.Update()
                bu4text.Text = bu4.ToString()
                bu4text.Update()
                bu5text.Text = bu5.ToString()
                bu5text.Update()
                bu6text.Text = bu6.ToString()
                bu6text.Update()
                bu7text.Text = bu7.ToString()
                bu7text.Update()
                bu8text.Text = bu8.ToString()
                bu8text.Update()
                bu9text.Text = bu9.ToString()
                bu9text.Update()
                bu10text.Text = bu10.ToString()
                bu10text.Update()
                bu11text.Text = bu11.ToString()
                bu11text.Update()
                bu12text.Text = bu12.ToString()
                bu12text.Update()
                bu13text.Text = "La hoja lenta no lleva este perfil"
                bu13text.Update()
                bu14text.Text = bu14.ToString()
                bu14text.Update()
                bu15text.Text = bu15.ToString()
                bu15text.Update()

                Panel4.Visible = True
                Panel3.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Label198.Visible = True
                Panel12.Visible = True
                TabControl2.SelectedIndex = 1


                '****************** PL SUPERIOR E INFERIOR - CASO I ************************************************************************************************************

            ElseIf (cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plintos silicona superior e inferior") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto superior de pinza e inferior de silicona")) And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 2 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías PL (Superior e inferior) con 2 hojas fijas y 2 móviles

                AHF = (wh / 4) + (solPL / 2)
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías PL (Superior e inferior) con 2 hojas fijas y 2 móviles

                HHF = hh
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías PL (Superior e inferior) con 2 hojas fijas y 2 móviles

                AHM = (wh / 4) + (solPL / 2)
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías PL (Superior e inferior) con 2 hojas fijas y 2 móviles

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías PL (Superior e inferior) con 2 hojas fijas y 2 móviles

                AVF = AHF
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías PL (Superior e inferior) con 2 hojas fijas y 2 móviles

                HVF = HHF - desv2PLF
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías PL (Superior e inferior) con 2 hojas fijas y 2 móviles

                AVM = AHM
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías PL (Superior e inferior) con 2 hojas fijas y 2 móviles

                HVM = HHM - desv2PLM
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calculo las medidas de los componentes de la carpintería PL (Superior e inferior)

                c1 = AHF
                c2 = AHM
                c3 = AHM
                c4 = wh - (2 * AHF)
                c5 = wh
                c6 = c5


                c1text.Text = c1.ToString()
                c1text.Update()
                c2text.Text = c2.ToString()
                c2text.Update()
                c3text.Text = c3.ToString()
                c3text.Update()
                c4text.Text = c4.ToString()
                c4text.Update()
                c5text.Text = c5.ToString()
                c5text.Update()
                c6text.Text = c6.ToString()
                c6text.Update()

                cu1 = 4
                cu2 = 2
                cu3 = 2
                cu4 = 1
                cu5 = 1
                cu6 = 1

                cu1text.Text = cu1.ToString()
                cu1text.Update()
                cu2text.Text = cu2.ToString()
                cu2text.Update()
                cu3text.Text = cu3.ToString()
                cu3text.Update()
                cu4text.Text = cu4.ToString()
                cu4text.Update()
                cu5text.Text = cu5.ToString()
                cu5text.Update()
                cu6text.Text = cu6.ToString()
                cu6text.Update()

                Panel5.Visible = True
                Panel3.Visible = False
                Panel4.Visible = False
                Panel6.Visible = False
                Label73.Visible = True
                c3text.Visible = True
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                TabControl2.SelectedIndex = 2

                '****************** PL SUPERIOR E INFERIOR - CASO III ***********************************************************************************************************

            ElseIf (cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plintos silicona superior e inferior") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto superior de pinza e inferior de silicona")) And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 1 Hoja Móvil") Then

                ' Calcula el ancho de la hoja fija para carpinterías PL (Superior e Inferior) con 1 hoja fija y 1 móvil

                AHF = (wh / 2) + (solPL / 2)
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías PL (Superior e Inferior) con 1 hoja fija y 1 móvil

                HHF = hh
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías PL (Superior e Inferior) con 1 hoja fija y 1 móvil

                AHM = (wh / 2) + (solPL / 2)
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías PL (Superior e Inferior) con 1 hoja fija y 1 móvil

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías PL (Superior e Inferior) con 1 hoja fija y 1 móvil

                AVF = AHF
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías PL (Superior e Inferior) con 1 hoja fija y 1 móvil

                HVF = HHF - desv2PLF
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías PL (Superior e Inferior) con 1 hoja fija y 1 móvil

                AVM = AHM
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías PL (Superior e Inferior) con 1 hoja fija y 1 móvil

                HVM = HHM - desv2PLM
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calculo las medidas de los componentes de la carpintería PL (Superior e inferior)

                c1 = AHF
                c2 = AHM
                c3 = AHM
                c4 = wh - AHF
                c5 = wh
                c6 = c5

                c1text.Text = c1.ToString()
                c1text.Update()
                c2text.Text = c2.ToString()
                c2text.Update()
                c3text.Text = c3.ToString()
                c3text.Update()
                c4text.Text = c4.ToString()
                c4text.Update()
                c5text.Text = c5.ToString()
                c5text.Update()
                c6text.Text = c6.ToString()
                c6text.Update()

                cu1 = 2
                cu2 = 1
                cu3 = 1
                cu4 = 1
                cu5 = 1
                cu6 = 1

                cu1text.Text = cu1.ToString()
                cu1text.Update()
                cu2text.Text = cu2.ToString()
                cu2text.Update()
                cu3text.Text = cu3.ToString()
                cu3text.Update()
                cu4text.Text = cu4.ToString()
                cu4text.Update()
                cu5text.Text = cu5.ToString()
                cu5text.Update()
                cu6text.Text = cu6.ToString()
                cu6text.Update()

                Panel5.Visible = True
                Panel3.Visible = False
                Panel4.Visible = False
                Panel6.Visible = False
                Label73.Visible = True
                c3text.Visible = True
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                TabControl2.SelectedIndex = 2

                '****************** PLINTOS SUPERIOR E INFERIOR - CASO V *************************************************************************************************************

            ElseIf (cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plintos silicona superior e inferior") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto superior de pinza e inferior de silicona")) And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 4 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                AHF = ((wh + 200) / 6)
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                HHF = hh
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                AHM = ((wh + 200) / 6)
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                HHM = hh - 10
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                AHML = ((wh + 200) / 6)
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                ' Calcula la altura de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                HHML = hh
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                AVF = ((wh + 200) / 6)
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                HVF = HHF - 95
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                AVM = ((wh + 200) / 6)
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                HVM = HHM - desv2PLM
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                AVML = ((wh + 200) / 6)
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                HVML = HHML - desv2PLM
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Calculo las medidas de los componentes de la carpintería PL (Superior e inferior)

                c1 = AHF
                c2 = AHM
                c3 = AHM
                c4 = wh - (2 * AHF)
                c5 = wh
                c6 = c5
                c7 = AHML
                c8 = AHML

                c1text.Text = c1.ToString()
                c1text.Update()
                c2text.Text = c2.ToString()
                c2text.Update()
                c3text.Text = c3.ToString()
                c3text.Update()
                c4text.Text = c4.ToString()
                c4text.Update()
                c5text.Text = c5.ToString()
                c5text.Update()
                c6text.Text = c6.ToString()
                c6text.Update()
                c7text.Text = c7.ToString()
                c7text.Update()
                c8text.Text = c8.ToString()
                c8text.Update()

                cu1 = 2
                cu2 = 2
                cu3 = 2
                cu4 = 1
                cu5 = 1
                cu6 = 1
                cu7 = 2
                cu8 = 2

                cu1text.Text = cu1.ToString()
                cu1text.Update()
                cu2text.Text = cu2.ToString()
                cu2text.Update()
                cu3text.Text = cu3.ToString()
                cu3text.Update()
                cu4text.Text = cu4.ToString()
                cu4text.Update()
                cu5text.Text = cu5.ToString()
                cu5text.Update()
                cu6text.Text = cu6.ToString()
                cu6text.Update()
                cu7text.Text = cu7.ToString()
                cu7text.Update()
                cu8text.Text = cu8.ToString()
                cu8text.Update()

                Panel5.Visible = True
                Panel3.Visible = False
                Panel4.Visible = False
                Panel6.Visible = False
                Label73.Visible = True
                c3text.Visible = True
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Panel13.Visible = True
                Label204.Visible = True
                TabControl2.SelectedIndex = 2

                '****************** PLINTOS SUPERIOR E INFERIOR - CASO VII ***********************************************************************************************************

            ElseIf (cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plintos silicona superior e inferior") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto superior de pinza e inferior de silicona")) And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 2 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                AHF = ((wh + 100) / 3)
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                HHF = hh
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                AHM = ((wh + 100) / 3)
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                HHM = hh - 10
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                AHML = ((wh + 100) / 3)
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                ' Calcula la altura de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                HHML = hh
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                AVF = ((wh + 100) / 3)
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                HVF = HHF - 95
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                AVM = ((wh + 100) / 3)
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                HVM = HHM - desv2PLM
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                AVML = ((wh + 100) / 3)
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                HVML = HHML - desv2PLM
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Calculo las medidas de los componentes de la carpintería PL (Superior e inferior)

                c1 = AHF
                c2 = AHM
                c3 = AHM
                c4 = wh - (AHF)
                c5 = wh
                c6 = c5
                c7 = AHML
                c8 = AHML

                c1text.Text = c1.ToString()
                c1text.Update()
                c2text.Text = c2.ToString()
                c2text.Update()
                c3text.Text = c3.ToString()
                c3text.Update()
                c4text.Text = c4.ToString()
                c4text.Update()
                c5text.Text = c5.ToString()
                c5text.Update()
                c6text.Text = c6.ToString()
                c6text.Update()
                c7text.Text = c7.ToString()
                c7text.Update()
                c8text.Text = c8.ToString()
                c8text.Update()

                cu1 = 1
                cu2 = 1
                cu3 = 1
                cu4 = 1
                cu5 = 1
                cu6 = 1
                cu7 = 1
                cu8 = 1

                cu1text.Text = cu1.ToString()
                cu1text.Update()
                cu2text.Text = cu2.ToString()
                cu2text.Update()
                cu3text.Text = cu3.ToString()
                cu3text.Update()
                cu4text.Text = cu4.ToString()
                cu4text.Update()
                cu5text.Text = cu5.ToString()
                cu5text.Update()
                cu6text.Text = cu6.ToString()
                cu6text.Update()
                cu7text.Text = cu7.ToString()
                cu7text.Update()
                cu8text.Text = cu8.ToString()
                cu8text.Update()

                Panel5.Visible = True
                Panel3.Visible = False
                Panel4.Visible = False
                Panel6.Visible = False
                Label73.Visible = True
                c3text.Visible = True
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Panel13.Visible = True
                Label204.Visible = True
                TabControl2.SelectedIndex = 2


                '****************** PL SUPERIOR - CASO I ************************************************************************************************************

                ElseIf (cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto silicona. Sólo superior") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto de pinza. Sólo Superior")) And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 2 Hojas Móviles") Then

                    ' Calcula el ancho de la hoja fija para carpinterías PL (Superior) con 2 hojas fijas y 2 móviles

                    AHF = (wh / 4) + (solPL / 2)
                    AHFtext.Text = AHF.ToString()
                    AHFtext.Update()

                    ' Calcula la altura de la hoja fija para carpinterías PL (Superior) con 2 hojas fijas y 2 móviles

                    HHF = hh
                    HHFtext.Text = HHF.ToString()
                    HHFtext.Update()

                    ' Calcula el ancho de la hoja móvil para carpinterías PL (Superior) con 2 hojas fijas y 2 móviles

                    AHM = (wh / 4) + (solPL / 2)
                    AHMtext.Text = AHM.ToString()
                    AHMtext.Update()

                    ' Calcula la altura de la hoja móvil para carpinterías PL (Superior) con 2 hojas fijas y 2 móviles

                    HHM = hh
                    HHMtext.Text = HHM.ToString()
                    HHMtext.Update()

                    ' Calcula el ancho del vidrio de la hoja fija para carpinterías PL (Superior) con 2 hojas fijas y 2 móviles

                    AVF = AHF
                    AVFtext.Text = AVF.ToString()
                    AVFtext.Update()

                    ' Calcula la altura del vidrio de la hoja fija para carpinterías PL (Superior) con 2 hojas fijas y 2 móviles

                    HVF = HHF - desv1PLF
                    HVFtext.Text = HVF.ToString()
                    HVFtext.Update()

                    ' Calcula el ancho del vidrio de la hoja móvil para carpinterías PL (Superior) con 2 hojas fijas y 2 móviles

                    AVM = AHM
                    AVMtext.Text = AVM.ToString()
                    AVMtext.Update()

                    ' Calcula la altura del vidrio de la hoja móvil para carpinterías PL (Superior) con 2 hojas fijas y 2 móviles

                    HVM = HHM - desv1PLM
                    HVMtext.Text = HVM.ToString()
                    HVMtext.Update()

                    ' Calculo las medidas de los componentes de la carpintería PL (Superior)

                    c1 = AHF
                    c2 = AHM
                    c4 = wh - (2 * AHF)
                    c5 = wh
                    c6 = c5

                    c1text.Text = c1.ToString()
                    c1text.Update()
                    c2text.Text = c2.ToString()
                    c2text.Update()
                    c3text.Text = c3.ToString()
                    c3text.Update()
                    c4text.Text = c4.ToString()
                    c4text.Update()
                    c5text.Text = c5.ToString()
                    c5text.Update()
                    c6text.Text = c6.ToString()
                    c6text.Update()

                    cu1 = 4
                    cu2 = 2
                    cu3 = 2
                    cu4 = 1
                    cu5 = 1
                    cu6 = 1

                    cu1text.Text = cu1.ToString()
                    cu1text.Update()
                    cu2text.Text = cu2.ToString()
                    cu2text.Update()
                    cu3text.Text = cu3.ToString()
                    cu3text.Update()
                    cu4text.Text = cu4.ToString()
                    cu4text.Update()
                    cu5text.Text = cu5.ToString()
                    cu5text.Update()
                    cu6text.Text = cu6.ToString()
                    cu6text.Update()

                    Panel5.Visible = True
                    Panel3.Visible = False
                    Panel4.Visible = False
                    Panel6.Visible = False
                    Label73.Visible = False
                    c3text.Visible = False
                    Label105.Visible = False
                    AHMLtext.Visible = False
                    HHMLtext.Visible = False
                    Label106.Visible = False
                    AVMLtext.Visible = False
                    HVMLtext.Visible = False
                    Label107.Visible = False
                    Label108.Visible = False
                    cu3text.Visible = False
                    TabControl2.SelectedIndex = 2

                    '****************** PL SUPERIOR - CASO III ***********************************************************************************************************

                ElseIf (cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto silicona. Sólo superior") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto de pinza. Sólo Superior")) And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 1 Hoja Móvil") Then

                    ' Calcula el ancho de la hoja fija para carpinterías PL (Superior) con 1 hoja fija y 1 móvil

                    AHF = (wh / 2) + (solPL / 2)
                    AHFtext.Text = AHF.ToString()
                    AHFtext.Update()

                    ' Calcula la altura de la hoja fija para carpinterías PL (Superior) con 1 hoja fija y 1 móvil

                    HHF = hh
                    HHFtext.Text = HHF.ToString()
                    HHFtext.Update()

                    ' Calcula el ancho de la hoja móvil para carpinterías PL (Superior) con 1 hoja fija y 1 móvil

                    AHM = (wh / 2) + (solPL / 2)
                    AHMtext.Text = AHM.ToString()
                    AHMtext.Update()

                    ' Calcula la altura de la hoja móvil para carpinterías PL (Superior) con 1 hoja fija y 1 móvil

                    HHM = hh
                    HHMtext.Text = HHM.ToString()
                    HHMtext.Update()

                    ' Calcula el ancho del vidrio de la hoja fija para carpinterías PL (Superior) con 1 hoja fija y 1 móvil

                    AVF = AHF
                    AVFtext.Text = AVF.ToString()
                    AVFtext.Update()

                    ' Calcula la altura del vidrio de la hoja fija para carpinterías PL (Superior) con 1 hoja fija y 1 móvil

                    HVF = HHF - desv1PLF
                    HVFtext.Text = HVF.ToString()
                    HVFtext.Update()

                    ' Calcula el ancho del vidrio de la hoja móvil para carpinterías PL (Superior) con 1 hoja fija y 1 móvil

                    AVM = AHM
                    AVMtext.Text = AVM.ToString()
                    AVMtext.Update()

                    ' Calcula la altura del vidrio de la hoja móvil para carpinterías PL (Superior) con 1 hoja fija y 1 móvil

                    HVM = HHM - desv1PLM
                    HVMtext.Text = HVM.ToString()
                    HVMtext.Update()

                    ' Calculo las medidas de los componentes de la carpintería PL (Superior)

                    c1 = AHF
                    c2 = AHM
                    c4 = wh - AHF
                    c5 = wh
                    c6 = c5

                    c1text.Text = c1.ToString()
                    c1text.Update()
                    c2text.Text = c2.ToString()
                    c2text.Update()
                    c3text.Text = c3.ToString()
                    c3text.Update()
                    c4text.Text = c4.ToString()
                    c4text.Update()
                    c5text.Text = c5.ToString()
                    c5text.Update()
                    c6text.Text = c6.ToString()
                    c6text.Update()

                    cu1 = 2
                    cu2 = 1
                    cu3 = 1
                    cu4 = 1
                    cu5 = 1
                    cu6 = 1

                    cu1text.Text = cu1.ToString()
                    cu1text.Update()
                    cu2text.Text = cu2.ToString()
                    cu2text.Update()
                    cu3text.Text = cu3.ToString()
                    cu3text.Update()
                    cu4text.Text = cu4.ToString()
                    cu4text.Update()
                    cu5text.Text = cu5.ToString()
                    cu5text.Update()
                    cu6text.Text = cu6.ToString()
                    cu6text.Update()

                    Panel5.Visible = True
                    Panel3.Visible = False
                    Panel4.Visible = False
                    Panel6.Visible = False
                    Label73.Visible = False
                    c3text.Visible = False
                    cu3text.Visible = False
                    Label105.Visible = False
                    AHMLtext.Visible = False
                    HHMLtext.Visible = False
                    Label106.Visible = False
                    AVMLtext.Visible = False
                    HVMLtext.Visible = False
                    Label107.Visible = False
                    Label108.Visible = False
                    TabControl2.SelectedIndex = 2

                    '****************** BW52 - CASO I *****************************************************************************************************************************

                ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("BW52") And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 2 Hojas Móviles") Then

                    ' Calcula el ancho de la hoja fija para carpinterías BW52 con 2 hojas fijas y 2 móviles

                    AHF = (wh / 4) + solBW52 / 2
                    AHFtext.Text = AHF.ToString()
                    AHFtext.Update()

                    ' Calcula la altura de la hoja fija para carpinterías BW52 con 2 hojas fijas y 2 móviles

                    HHF = hh - 10
                    HHFtext.Text = HHF.ToString()
                    HHFtext.Update()

                    ' Calcula el ancho de la hoja móvil para carpinterías BW52 con 2 hojas fijas y 2 móviles

                    AHM = (wh / 4) + solBW52 / 2
                    AHMtext.Text = AHM.ToString()
                    AHMtext.Update()

                    ' Calcula la altura de la hoja móvil para carpinterías BW52 con 2 hojas fijas y 2 móviles

                    HHM = hh
                    HHMtext.Text = HHM.ToString()
                    HHMtext.Update()

                    ' Calcula el ancho del vidrio de la hoja fija para carpinterías BW52 con 2 hojas fijas y 2 móviles

                    AVF = AHF - deshfBW52
                    AVFtext.Text = AVF.ToString()
                    AVFtext.Update()

                    ' Calcula la altura del vidrio de la hoja fija para carpinterías BW52 con 2 hojas fijas y 2 móviles

                    HVF = HHF - desvBW52
                    HVFtext.Text = HVF.ToString()
                    HVFtext.Update()

                    ' Calcula el ancho del vidrio de la hoja móvil para carpinterías BW52 con 2 hojas fijas y 2 móviles

                    AVM = AHM - deshmBW52
                    AVMtext.Text = AVM.ToString()
                    AVMtext.Update()

                    ' Calcula la altura del vidrio de la hoja móvil para carpinterías BW52 con 2 hojas fijas y 2 móviles

                    HVM = HHM - desvBW52
                    HVMtext.Text = HVM.ToString()
                    HVMtext.Update()

                    ' Calculo las medidas de los componentes de la carpintería BW52

                    d1 = HHF - 22
                    d2 = HHF + 8
                    d3 = HHF + 30
                    d4 = AHF - 10
                    d5 = AHF + 34
                    d6 = HHF - 113
                    d7 = AHF - 144
                    d8 = AHF - 10
                    d9 = HHM - 22
                    d10 = HHM
                    d11 = HHM
                    d12 = HHM
                    d13 = AHM - 20
                    d14 = AHM + 24
                    d15 = AHM - 13
                    d16 = HHM - 113
                    d17 = AHM - 154
                    d18 = AHM - 20
                    d19 = wh - (2 * AHF)
                    d20 = AHF
                    d21 = wh
                    d22 = d21

                    d1text.Text = d1.ToString()
                    d1text.Update()
                    d2text.Text = d2.ToString()
                    d2text.Update()
                    d3text.Text = d3.ToString()
                    d3text.Update()
                    d4text.Text = d4.ToString()
                    d4text.Update()
                    d5text.Text = d5.ToString()
                    d5text.Update()
                    d6text.Text = d6.ToString()
                    d6text.Update()
                    d7text.Text = d7.ToString()
                    d7text.Update()
                    d8text.Text = d8.ToString()
                    d8text.Update()
                    d9text.Text = d9.ToString()
                    d9text.Update()
                    d10text.Text = d10.ToString()
                    d10text.Update()
                    d11text.Text = d11.ToString()
                    d11text.Update()
                    d12text.Text = d12.ToString()
                    d12text.Update()
                    d13text.Text = d13.ToString()
                    d13text.Update()
                    d14text.Text = d14.ToString()
                    d14text.Update()
                    d15text.Text = d15.ToString()
                    d15text.Update()
                    d16text.Text = d16.ToString()
                    d16text.Update()
                    d17text.Text = d17.ToString()
                    d17text.Update()
                    d18text.Text = d18.ToString()
                    d18text.Update()
                    d19text.Text = d19.ToString()
                    d19text.Update()
                    d20text.Text = d20.ToString()
                    d20text.Update()
                    d21text.Text = d21.ToString()
                    d21text.Update()
                    d22text.Text = d22.ToString()
                    d22text.Update()

                    du1 = 4
                    du2 = 2
                    du3 = 2
                    du4 = 2
                    du5 = 2
                    du6 = 4
                    du7 = 4
                    du8 = 2
                    du9 = 4
                    du10 = 2
                    du11 = 2
                    du12 = 2
                    du13 = 2
                    du14 = 2
                    du15 = 2
                    du16 = 4
                    du17 = 4
                    du18 = 2
                    du19 = 1
                    du20 = 2
                    du21 = 1
                    du22 = 1

                    du1text.Text = du1.ToString()
                    du1text.Update()
                    du2text.Text = du2.ToString()
                    du2text.Update()
                    du3text.Text = du3.ToString()
                    du3text.Update()
                    du4text.Text = du4.ToString()
                    du4text.Update()
                    du5text.Text = du5.ToString()
                    du5text.Update()
                    du6text.Text = du6.ToString()
                    du6text.Update()
                    du7text.Text = du7.ToString()
                    du7text.Update()
                    du8text.Text = du8.ToString()
                    du8text.Update()
                    du9text.Text = du9.ToString()
                    du9text.Update()
                    du10text.Text = du10.ToString()
                    du10text.Update()
                    du11text.Text = du11.ToString()
                    du11text.Update()
                    du12text.Text = du12.ToString()
                    du12text.Update()
                    du13text.Text = du13.ToString()
                    du13text.Update()
                    du14text.Text = du14.ToString()
                    du14text.Update()
                    du15text.Text = du15.ToString()
                    du15text.Update()
                    du16text.Text = du16.ToString()
                    du16text.Update()
                    du17text.Text = du17.ToString()
                    du17text.Update()
                    du18text.Text = du18.ToString()
                    du18text.Update()
                    du19text.Text = du19.ToString()
                    du19text.Update()
                    du20text.Text = du20.ToString()
                    du20text.Update()
                    du21text.Text = du21.ToString()
                    du21text.Update()
                    du22text.Text = du22.ToString()
                    du22text.Update()

                    Panel3.Visible = False
                    Panel4.Visible = False
                    Panel5.Visible = False
                    Panel6.Visible = True
                    Label105.Visible = False
                    AHMLtext.Visible = False
                    HHMLtext.Visible = False
                    Label106.Visible = False
                    AVMLtext.Visible = False
                    HVMLtext.Visible = False
                    Label107.Visible = False
                    Label108.Visible = False
                    TabControl2.SelectedIndex = 3

                    '****************** BW52 - CASO III ****************************************************************************************************************************

                ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("BW52") And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 1 Hoja Móvil") Then

                    ' Calcula el ancho de la hoja fija para carpinterías BW52 con 1 hoja fija y 1 móvil

                    AHF = (wh / 2) + (solBW52 / 2)
                    AHFtext.Text = AHF.ToString()
                    AHFtext.Update()

                    ' Calcula la altura de la hoja fija para carpinterías BW52 con 1 hoja fija y 1 móvil

                    HHF = hh - 10
                    HHFtext.Text = HHF.ToString()
                    HHFtext.Update()

                    ' Calcula el ancho de la hoja móvil para carpinterías BW52 con 1 hoja fija y 1 móvil

                    AHM = (wh / 2) + (solBW52 / 2)
                    AHMtext.Text = AHM.ToString()
                    AHMtext.Update()

                    ' Calcula la altura de la hoja móvil para carpinterías BW52 con 1 hoja fija y 1 móvil

                    HHM = hh
                    HHMtext.Text = HHM.ToString()
                    HHMtext.Update()

                    ' Calcula el ancho del vidrio de la hoja fija para carpinterías BW52 con 1 hoja fija y 1 móvil

                    AVF = AHF - deshfBW52
                    AVFtext.Text = AVF.ToString()
                    AVFtext.Update()

                    ' Calcula la altura del vidrio de la hoja fija para carpinterías BW52 con 1 hoja fija y 1 móvil

                    HVF = HHF - desvBW52
                    HVFtext.Text = HVF.ToString()
                    HVFtext.Update()

                    ' Calcula el ancho del vidrio de la hoja móvil para carpinterías BW52 con 1 hoja fija y 1 móvil

                    AVM = AHM - deshmBW52
                    AVMtext.Text = AVM.ToString()
                    AVMtext.Update()

                    ' Calcula la altura del vidrio de la hoja móvil para carpinterías BW52 con 1 hoja fija y 1 móvil

                    HVM = HHM - desvBW52
                    HVMtext.Text = HVM.ToString()
                    HVMtext.Update()

                    ' Calculo las medidas de los componentes de la carpintería BW52

                    d1 = HHF - 22
                    d2 = HHF + 8
                    d3 = HHF + 30
                    d4 = AHF - 10
                    d5 = AHF + 34
                    d6 = HHF - 113
                    d7 = AHF - 144
                    d8 = AHF - 10
                    d9 = HHM - 22
                    d10 = HHM
                    d11 = HHM
                    d12 = HHM
                    d13 = AHM - 20
                    d14 = AHM + 24
                    d15 = AHM - 13
                    d16 = HHM - 113
                    d17 = AHM - 154
                    d18 = AHM - 20
                    d19 = wh - AHF
                    d20 = AHF
                    d21 = wh 
                    d22 = d21

                    d1text.Text = d1.ToString()
                    d1text.Update()
                    d2text.Text = d2.ToString()
                    d2text.Update()
                    d3text.Text = d3.ToString()
                    d3text.Update()
                    d4text.Text = d4.ToString()
                    d4text.Update()
                    d5text.Text = d5.ToString()
                    d5text.Update()
                    d6text.Text = d6.ToString()
                    d6text.Update()
                    d7text.Text = d7.ToString()
                    d7text.Update()
                    d8text.Text = d8.ToString()
                    d8text.Update()
                    d9text.Text = d9.ToString()
                    d9text.Update()
                    d10text.Text = d10.ToString()
                    d10text.Update()
                    d11text.Text = d11.ToString()
                    d11text.Update()
                    d12text.Text = d12.ToString()
                    d12text.Update()
                    d13text.Text = d13.ToString()
                    d13text.Update()
                    d14text.Text = d14.ToString()
                    d14text.Update()
                    d15text.Text = d15.ToString()
                    d15text.Update()
                    d16text.Text = d16.ToString()
                    d16text.Update()
                    d17text.Text = d17.ToString()
                    d17text.Update()
                    d18text.Text = d18.ToString()
                    d18text.Update()
                    d19text.Text = d19.ToString()
                    d19text.Update()
                    d20text.Text = d20.ToString()
                    d20text.Update()
                    d21text.Text = d21.ToString()
                    d21text.Update()
                    d22text.Text = d22.ToString()
                    d22text.Update()

                    du1 = 2
                    du2 = 1
                    du3 = 1
                    du4 = 1
                    du5 = 1
                    du6 = 2
                    du7 = 2
                    du8 = 1
                    du9 = 2
                    du10 = 1
                    du11 = 1
                    du12 = 1
                    du13 = 1
                    du14 = 1
                    du15 = 1
                    du16 = 2
                    du17 = 2
                    du18 = 1
                    du19 = 1
                    du20 = 1
                    du21 = 1
                    du22 = 1

                    du1text.Text = du1.ToString()
                    du1text.Update()
                    du2text.Text = du2.ToString()
                    du2text.Update()
                    du3text.Text = du3.ToString()
                    du3text.Update()
                    du4text.Text = du4.ToString()
                    du4text.Update()
                    du5text.Text = du5.ToString()
                    du5text.Update()
                    du6text.Text = du6.ToString()
                    du6text.Update()
                    du7text.Text = du7.ToString()
                    du7text.Update()
                    du8text.Text = du8.ToString()
                    du8text.Update()
                    du9text.Text = du9.ToString()
                    du9text.Update()
                    du10text.Text = du10.ToString()
                    du10text.Update()
                    du11text.Text = du11.ToString()
                    du11text.Update()
                    du12text.Text = du12.ToString()
                    du12text.Update()
                    du13text.Text = du13.ToString()
                    du13text.Update()
                    du14text.Text = du14.ToString()
                    du14text.Update()
                    du15text.Text = du15.ToString()
                    du15text.Update()
                    du16text.Text = du16.ToString()
                    du16text.Update()
                    du17text.Text = du17.ToString()
                    du17text.Update()
                    du18text.Text = du18.ToString()
                    du18text.Update()
                    du19text.Text = du19.ToString()
                    du19text.Update()
                    du20text.Text = du20.ToString()
                    du20text.Update()
                    du21text.Text = du21.ToString()
                    du21text.Update()
                    du22text.Text = du22.ToString()
                    du22text.Update()

                    Panel3.Visible = False
                    Panel4.Visible = False
                    Panel5.Visible = False
                    Panel6.Visible = True
                    Label105.Visible = False
                    AHMLtext.Visible = False
                    HHMLtext.Visible = False
                    Label106.Visible = False
                    AVMLtext.Visible = False
                    HVMLtext.Visible = False
                    Label107.Visible = False
                    Label108.Visible = False
                    TabControl2.SelectedIndex = 3

                    '****************** BW52 - CASO V *****************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("BW52") And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 4 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                AHF = ((wh + 296) / 6)
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                HHF = hh - 10
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                AHM = ((wh + 296) / 6)
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                HHM = hh - 5
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho de la hoja móvil lenta para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                AHML = ((wh + 296) / 6)
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                ' Calcula la altura de la hoja móvil lenta para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                HHML = hh
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                AVF = AHF - 119
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                HVF = HHF - desvBW52
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                AVM = AHM - 119
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                HVM = HHM - desvBW52
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil lenta para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                AVML = AHML - 129
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil lenta para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                HVML = HHML - desvBW52
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Calculo las medidas de los componentes de la carpintería BW52

                d1 = HHF - 22
                d2 = HHF + 8
                d3 = HHF + 30
                d4 = AHF - 10
                d5 = AHF + 34
                d6 = HHF - 113
                d7 = AHF - 144
                d8 = AHF - 10
                d9 = HHM - 22
                d10 = HHM
                d11 = HHM
                d12 = HHM
                d13 = AHM - 20
                d14 = AHM + 24
                d15 = AHM - 13
                d16 = HHM - 113
                d17 = AHM - 154
                d18 = AHM - 20
                d19 = wh - (2 * AHF)
                d20 = AHF
                d21 = wh
                d22 = d21
                d23 = HHM - 22
                d24 = HHM
                d25 = HHM
                d26 = HHM
                d27 = AHM - 20
                d28 = AHM + 24
                d29 = AHM - 13
                d30 = HHM - 113
                d31 = AHM - 154
                d32 = AHM - 20

                d1text.Text = d1.ToString()
                d1text.Update()
                d2text.Text = d2.ToString()
                d2text.Update()
                d3text.Text = d3.ToString()
                d3text.Update()
                d4text.Text = d4.ToString()
                d4text.Update()
                d5text.Text = d5.ToString()
                d5text.Update()
                d6text.Text = d6.ToString()
                d6text.Update()
                d7text.Text = d7.ToString()
                d7text.Update()
                d8text.Text = d8.ToString()
                d8text.Update()
                d9text.Text = d9.ToString()
                d9text.Update()
                d10text.Text = d10.ToString()
                d10text.Update()
                d11text.Text = d11.ToString()
                d11text.Update()
                d12text.Text = d12.ToString()
                d12text.Update()
                d13text.Text = d13.ToString()
                d13text.Update()
                d14text.Text = d14.ToString()
                d14text.Update()
                d15text.Text = d15.ToString()
                d15text.Update()
                d16text.Text = d16.ToString()
                d16text.Update()
                d17text.Text = d17.ToString()
                d17text.Update()
                d18text.Text = d18.ToString()
                d18text.Update()
                d19text.Text = d19.ToString()
                d19text.Update()
                d20text.Text = d20.ToString()
                d20text.Update()
                d21text.Text = d21.ToString()
                d21text.Update()
                d22text.Text = d22.ToString()
                d22text.Update()
                d23text.Text = d23.ToString()
                d23text.Update()
                d24text.Text = d24.ToString()
                d24text.Update()
                d25text.Text = "La hoja lenta no lleva este perfil"
                d25text.Update()
                d26text.Text = d26.ToString()
                d26text.Update()
                d27text.Text = d27.ToString()
                d27text.Update()
                d28text.Text = d28.ToString()
                d28text.Update()
                d29text.Text = d29.ToString()
                d29text.Update()
                d30text.Text = d30.ToString()
                d30text.Update()
                d31text.Text = d31.ToString()
                d31text.Update()
                d32text.Text = d32.ToString()
                d32text.Update()

                du1 = 4
                du2 = 2
                du3 = 2
                du4 = 2
                du5 = 2
                du6 = 4
                du7 = 4
                du8 = 2
                du9 = 4
                du10 = 2
                du11 = 2
                du12 = 2
                du13 = 2
                du14 = 2
                du15 = 2
                du16 = 4
                du17 = 4
                du18 = 2
                du19 = 1
                du20 = 2
                du21 = 1
                du22 = 1
                du23 = 4
                du24 = 4
                du25 = 1
                du26 = 4
                du27 = 2
                du28 = 2
                du29 = 2
                du30 = 4
                du31 = 4
                du32 = 2

                du1text.Text = du1.ToString()
                du1text.Update()
                du2text.Text = du2.ToString()
                du2text.Update()
                du3text.Text = du3.ToString()
                du3text.Update()
                du4text.Text = du4.ToString()
                du4text.Update()
                du5text.Text = du5.ToString()
                du5text.Update()
                du6text.Text = du6.ToString()
                du6text.Update()
                du7text.Text = du7.ToString()
                du7text.Update()
                du8text.Text = du8.ToString()
                du8text.Update()
                du9text.Text = du9.ToString()
                du9text.Update()
                du10text.Text = du10.ToString()
                du10text.Update()
                du11text.Text = du11.ToString()
                du11text.Update()
                du12text.Text = du12.ToString()
                du12text.Update()
                du13text.Text = du13.ToString()
                du13text.Update()
                du14text.Text = du14.ToString()
                du14text.Update()
                du15text.Text = du15.ToString()
                du15text.Update()
                du16text.Text = du16.ToString()
                du16text.Update()
                du17text.Text = du17.ToString()
                du17text.Update()
                du18text.Text = du18.ToString()
                du18text.Update()
                du19text.Text = du19.ToString()
                du19text.Update()
                du20text.Text = du20.ToString()
                du20text.Update()
                du21text.Text = du21.ToString()
                du21text.Update()
                du22text.Text = du22.ToString()
                du22text.Update()
                du23text.Text = du23.ToString()
                du23text.Update()
                du24text.Text = du24.ToString()
                du24text.Update()
                du25text.Text = "La hoja lenta no lleva este perfil"
                du25text.Update()
                du26text.Text = du26.ToString()
                du26text.Update()
                du27text.Text = du27.ToString()
                du27text.Update()
                du28text.Text = du28.ToString()
                du28text.Update()
                du29text.Text = du29.ToString()
                du29text.Update()
                du30text.Text = du30.ToString()
                du30text.Update()
                du31text.Text = du31.ToString()
                du31text.Update()
                du32text.Text = du32.ToString()
                du32text.Update()

                Panel3.Visible = False
                Panel4.Visible = False
                Panel5.Visible = False
                Panel6.Visible = True
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Panel14.Visible = True
                Label209.Visible = True
                TabControl2.SelectedIndex = 3

                '****************** BW52 - CASO VII ***************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("BW52") And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 2 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                AHF = ((wh + 148) / 3)
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                HHF = hh - 10
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                AHM = ((wh + 148) / 3)
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                HHM = hh - 5
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho de la hoja móvil lenta para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                AHML = ((wh + 148) / 3)
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                ' Calcula la altura de la hoja móvil lenta para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                HHML = hh
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                AVF = AHF - 119
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                HVF = HHF - desvBW52
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                AVM = AHM - 119
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                HVM = HHM - desvBW52
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil lenta para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                AVML = AHML - 129
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil lenta para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                HVML = HHML - desvBW52
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Calculo las medidas de los componentes de la carpintería BW52

                d1 = HHF - 22
                d2 = HHF + 8
                d3 = HHF + 30
                d4 = AHF - 10
                d5 = AHF + 34
                d6 = HHF - 113
                d7 = AHF - 144
                d8 = AHF - 10
                d9 = HHM - 22
                d10 = HHM
                d11 = HHM
                d12 = HHM
                d13 = AHM - 20
                d14 = AHM + 24
                d15 = AHM - 13
                d16 = HHM - 113
                d17 = AHM - 154
                d18 = AHM - 20
                d19 = wh - (AHF)
                d20 = AHF
                d21 = wh
                d22 = d21
                d23 = HHM - 22
                d24 = HHM
                d25 = HHM
                d26 = HHM
                d27 = AHM - 20
                d28 = AHM + 24
                d29 = AHM - 13
                d30 = HHM - 113
                d31 = AHM - 154
                d32 = AHM - 20


                d1text.Text = d1.ToString()
                d1text.Update()
                d2text.Text = d2.ToString()
                d2text.Update()
                d3text.Text = d3.ToString()
                d3text.Update()
                d4text.Text = d4.ToString()
                d4text.Update()
                d5text.Text = d5.ToString()
                d5text.Update()
                d6text.Text = d6.ToString()
                d6text.Update()
                d7text.Text = d7.ToString()
                d7text.Update()
                d8text.Text = d8.ToString()
                d8text.Update()
                d9text.Text = d9.ToString()
                d9text.Update()
                d10text.Text = d10.ToString()
                d10text.Update()
                d11text.Text = d11.ToString()
                d11text.Update()
                d12text.Text = d12.ToString()
                d12text.Update()
                d13text.Text = d13.ToString()
                d13text.Update()
                d14text.Text = d14.ToString()
                d14text.Update()
                d15text.Text = d15.ToString()
                d15text.Update()
                d16text.Text = d16.ToString()
                d16text.Update()
                d17text.Text = d17.ToString()
                d17text.Update()
                d18text.Text = d18.ToString()
                d18text.Update()
                d19text.Text = d19.ToString()
                d19text.Update()
                d20text.Text = d20.ToString()
                d20text.Update()
                d21text.Text = d21.ToString()
                d21text.Update()
                d22text.Text = d22.ToString()
                d22text.Update()
                d23text.Text = d23.ToString()
                d23text.Update()
                d24text.Text = d24.ToString()
                d24text.Update()
                d25text.Text = "La hoja lenta no lleva este perfil"
                d25text.Update()
                d26text.Text = d26.ToString()
                d26text.Update()
                d27text.Text = d27.ToString()
                d27text.Update()
                d28text.Text = d28.ToString()
                d28text.Update()
                d29text.Text = d29.ToString()
                d29text.Update()
                d30text.Text = d30.ToString()
                d30text.Update()
                d31text.Text = d31.ToString()
                d31text.Update()
                d32text.Text = d32.ToString()
                d32text.Update()

                du1 = 2
                du2 = 1
                du3 = 1
                du4 = 1
                du5 = 1
                du6 = 2
                du7 = 2
                du8 = 1
                du9 = 2
                du10 = 1
                du11 = 1
                du12 = 1
                du13 = 1
                du14 = 1
                du15 = 1
                du16 = 2
                du17 = 2
                du18 = 1
                du19 = 1
                du20 = 1
                du21 = 1
                du22 = 1
                du23 = 2
                du24 = 2
                du25 = 1
                du26 = 2
                du27 = 1
                du28 = 1
                du29 = 1
                du30 = 2
                du31 = 2
                du32 = 1

                du1text.Text = du1.ToString()
                du1text.Update()
                du2text.Text = du2.ToString()
                du2text.Update()
                du3text.Text = du3.ToString()
                du3text.Update()
                du4text.Text = du4.ToString()
                du4text.Update()
                du5text.Text = du5.ToString()
                du5text.Update()
                du6text.Text = du6.ToString()
                du6text.Update()
                du7text.Text = du7.ToString()
                du7text.Update()
                du8text.Text = du8.ToString()
                du8text.Update()
                du9text.Text = du9.ToString()
                du9text.Update()
                du10text.Text = du10.ToString()
                du10text.Update()
                du11text.Text = du11.ToString()
                du11text.Update()
                du12text.Text = du12.ToString()
                du12text.Update()
                du13text.Text = du13.ToString()
                du13text.Update()
                du14text.Text = du14.ToString()
                du14text.Update()
                du15text.Text = du15.ToString()
                du15text.Update()
                du16text.Text = du16.ToString()
                du16text.Update()
                du17text.Text = du17.ToString()
                du17text.Update()
                du18text.Text = du18.ToString()
                du18text.Update()
                du19text.Text = du19.ToString()
                du19text.Update()
                du20text.Text = du20.ToString()
                du20text.Update()
                du21text.Text = du21.ToString()
                du21text.Update()
                du22text.Text = du22.ToString()
                du22text.Update()
                du23text.Text = du23.ToString()
                du23text.Update()
                du24text.Text = du24.ToString()
                du24text.Update()
                du25text.Text = "La hoja lenta no lleva este perfil"
                du25text.Update()
                du26text.Text = du26.ToString()
                du26text.Update()
                du27text.Text = du27.ToString()
                du27text.Update()
                du28text.Text = du28.ToString()
                du28text.Update()
                du29text.Text = du29.ToString()
                du29text.Update()
                du30text.Text = du30.ToString()
                du30text.Update()
                du31text.Text = du31.ToString()
                du31text.Update()
                du32text.Text = du32.ToString()
                du32text.Update()

                Panel3.Visible = False
                Panel4.Visible = False
                Panel5.Visible = False
                Panel6.Visible = True
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Panel14.Visible = True
                Label209.Visible = True
                TabControl2.SelectedIndex = 3

                '****************** AP94 - CASO I *****************************************************************************************************************************

                'ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("AP94") And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 2 Hojas Móviles") Then

                '    If (wh > 4380) Then

                '        MsgBox("El ancho de hueco introducido está fuera de rango. Consulte con su comercial para soluciones personalizadas")

                '        Return

                '    End If

                '    ' Calcula el ancho de la hoja móvil para carpinterías AP94 con 2 hojas fijas y 2 móviles

                '    AHM = (wh / 4)
                '    AHMtext.Text = AHM.ToString()
                '    AHMtext.Update()

                '    ' Calcula el ancho de la hoja fija para carpinterías AP94 con 2 hojas fijas y 2 móviles

                '    AHF = AHM + 70
                '    AHFtext.Text = AHF.ToString()
                '    AHFtext.Update()

                '    ' Calcula la altura de la hoja fija para carpinterías AP94 con 2 hojas fijas y 2 móviles

                '    HHF = hh - 20
                '    HHFtext.Text = HHF.ToString()
                '    HHFtext.Update()

                '    ' Calcula la altura de la hoja móvil para carpinterías AP94 con 2 hojas fijas y 2 móviles

                '    HHM = HHF - 25
                '    HHMtext.Text = HHM.ToString()
                '    HHMtext.Update()

                '    ' Calcula el ancho del vidrio de la hoja fija para carpinterías AP94 con 2 hojas fijas y 2 móviles

                '    AVF = AHF - 209
                '    AVFtext.Text = AVF.ToString()
                '    AVFtext.Update()

                '    ' Calcula la altura del vidrio de la hoja fija para carpinterías AP94 con 2 hojas fijas y 2 móviles

                '    HVF = HHF - 164
                '    HVFtext.Text = HVF.ToString()
                '    HVFtext.Update()

                '    ' Calcula el ancho del vidrio de la hoja móvil para carpinterías AP94 con 2 hojas fijas y 2 móviles

                '    AVM = AHM - 127
                '    AVMtext.Text = AVM.ToString()
                '    AVMtext.Update()

                '    ' Calcula la altura del vidrio de la hoja móvil para carpinterías AP94 con 2 hojas fijas y 2 móviles

                '    HVM = HHM - 159
                '    HVMtext.Text = HVM.ToString()
                '    HVMtext.Update()

                '    ' Calculo las medidas de los componentes de la carpintería AP94

                '    e1 = HHF
                '    e2 = HHF - 46
                '    e3 = HHF - 41
                '    e4 = AHF - 46
                '    e5 = AHF - 59
                '    e6 = HHF - 146
                '    e7 = AHF - 53
                '    e8 = AHF - 190
                '    e9 = AHF - 195
                '    e10 = AHF - 190
                '    e11 = AHF - 235
                '    e12 = HHM - 41
                '    e13 = HHM - 54
                '    e14 = HHM - 54
                '    e15 = HHM - 54
                '    e16 = HHM - 141
                '    e17 = AHM - 17
                '    e18 = AHM - 108
                '    e19 = AHM - 113
                '    e20 = AHM - 108
                '    e21 = AHM - 153
                '    e22 = HHM - 41
                '    e23 = HHM - 41
                '    e24 = wh
                '    e25 = wh - AHF
                '    e26 = wh - AHF
                '    e27 = AHF
                '    e28 = wh
                '    e29 = wh

                '    e1text.Text = e1.ToString()
                '    e1text.Update()
                '    e2text.Text = e2.ToString()
                '    e2text.Update()
                '    e3text.Text = e3.ToString()
                '    e3text.Update()
                '    e4text.Text = e4.ToString()
                '    e4text.Update()
                '    e5text.Text = e5.ToString()
                '    e5text.Update()
                '    e6text.Text = e6.ToString()
                '    e6text.Update()
                '    e7text.Text = e7.ToString()
                '    e7text.Update()
                '    e8text.Text = e8.ToString()
                '    e8text.Update()
                '    e9text.Text = e9.ToString()
                '    e9text.Update()
                '    e10text.Text = e10.ToString()
                '    e10text.Update()
                '    e11text.Text = e11.ToString()
                '    e11text.Update()
                '    e12text.Text = e12.ToString()
                '    e12text.Update()
                '    e13text.Text = e13.ToString()
                '    e13text.Update()
                '    e14text.Text = e14.ToString()
                '    e14text.Update()
                '    e15text.Text = e15.ToString()
                '    e15text.Update()
                '    e16text.Text = e16.ToString()
                '    e16text.Update()
                '    e17text.Text = e17.ToString()
                '    e17text.Update()
                '    e18text.Text = e18.ToString()
                '    e18text.Update()
                '    e19text.Text = e19.ToString()
                '    e19text.Update()
                '    e20text.Text = e20.ToString()
                '    e20text.Update()
                '    e21text.Text = e21.ToString()
                '    e21text.Update()
                '    e22text.Text = e22.ToString()
                '    e22text.Update()
                '    e23text.Text = e23.ToString()
                '    e23text.Update()
                '    e24text.Text = e24.ToString()
                '    e24text.Update()
                '    e25text.Text = e25.ToString()
                '    e25text.Update()
                '    e26text.Text = e26.ToString()
                '    e26text.Update()
                '    e27text.Text = e27.ToString()
                '    e27text.Update()
                '    e28text.Text = e28.ToString()
                '    e28text.Update()
                '    e29text.Text = e29.ToString()
                '    e29text.Update()

                '    eu1 = 4
                '    eu2 = 2
                '    eu3 = 2
                '    eu4 = 2
                '    eu5 = 2
                '    eu6 = 4
                '    eu7 = 4
                '    eu8 = 2
                '    eu9 = 4
                '    eu10 = 2
                '    eu11 = 2
                '    eu12 = 2
                '    eu13 = 2
                '    eu14 = 2
                '    eu15 = 2
                '    eu16 = 4
                '    eu17 = 4
                '    eu18 = 2
                '    eu19 = 1
                '    eu20 = 2
                '    eu21 = 1
                '    eu22 = 1
                '    eu23 = 2
                '    eu24 = 1
                '    eu25 = 1
                '    eu26 = 2
                '    eu27 = 1
                '    eu28 = 1
                '    eu29 = 87

                '    eu1text.Text = e1.ToString()
                '    eu1text.Update()
                '    eu2text.Text = e2.ToString()
                '    eu2text.Update()
                '    eu3text.Text = e3.ToString()
                '    eu3text.Update()
                '    eu4text.Text = e4.ToString()
                '    eu4text.Update()
                '    eu5text.Text = e5.ToString()
                '    eu5text.Update()
                '    eu6text.Text = e6.ToString()
                '    eu6text.Update()
                '    eu7text.Text = e7.ToString()
                '    eu7text.Update()
                '    eu8text.Text = e8.ToString()
                '    eu8text.Update()
                '    eu9text.Text = e9.ToString()
                '    eu9text.Update()
                '    eu10text.Text = e10.ToString()
                '    eu10text.Update()
                '    eu11text.Text = e11.ToString()
                '    eu11text.Update()
                '    eu12text.Text = e12.ToString()
                '    eu12text.Update()
                '    eu13text.Text = e13.ToString()
                '    eu13text.Update()
                '    eu14text.Text = e14.ToString()
                '    eu14text.Update()
                '    eu15text.Text = e15.ToString()
                '    eu15text.Update()
                '    eu16text.Text = e16.ToString()
                '    eu16text.Update()
                '    eu17text.Text = e17.ToString()
                '    eu17text.Update()
                '    eu18text.Text = e18.ToString()
                '    eu18text.Update()
                '    eu19text.Text = e19.ToString()
                '    eu19text.Update()
                '    eu20text.Text = e20.ToString()
                '    eu20text.Update()
                '    eu21text.Text = e21.ToString()
                '    eu21text.Update()
                '    eu22text.Text = e22.ToString()
                '    eu22text.Update()
                '    eu23text.Text = e23.ToString()
                '    eu23text.Update()
                '    eu24text.Text = e24.ToString()
                '    eu24text.Update()
                '    eu25text.Text = e25.ToString()
                '    eu25text.Update()
                '    eu26text.Text = e26.ToString()
                '    eu26text.Update()
                '    eu27text.Text = e27.ToString()
                '    eu27text.Update()
                '    eu28text.Text = e28.ToString()
                '    eu28text.Update()
                '    eu29text.Text = e29.ToString()
                '    eu29text.Update()

                '    Panel3.Visible = False
                '    Panel4.Visible = False
                '    Panel5.Visible = False
                '    Panel6.Visible = False
                '    Label105.Visible = False
                '    AHMLtext.Visible = False
                '    HHMLtext.Visible = False
                '    Label106.Visible = False
                '    AVMLtext.Visible = False
                '    HVMLtext.Visible = False
                '    Label107.Visible = False
                '    Label108.Visible = False
                '    Label143.Visible = False
                '    Label146.Visible = False
                '    e22text.Visible = False
                '    e23text.Visible = False
                '    eu22text.Visible = False
                '    eu23text.Visible = False
                '    TabControl2.SelectedIndex = 4

                '    '****************** AP94 - CASO III ****************************************************************************************************************************

                '    ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("AP94") And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 1 Hoja Móvil") Then

                '        If (wh > 2190) Then

                '            MsgBox("El ancho de hueco introducido está fuera de rango. Consulte con su comercial para soluciones personalizadas")

                '            Return

                '        End If

                '        ' Calcula el ancho de la hoja móvil para carpinterías AP94 con 1 hoja fija y 1 móvil

                '        AHM = (wh / 2)
                '        AHMtext.Text = AHM.ToString()
                '        AHMtext.Update()

                '        ' Calcula el ancho de la hoja fija para carpinterías AP94 con 1 hoja fija y 1 móvil

                '        AHF = AHM + 70
                '        AHFtext.Text = AHF.ToString()
                '        AHFtext.Update()

                '        ' Calcula la altura de la hoja fija para carpinterías AP94 con 1 hoja fija y 1 móvil

                '        HHF = hh - 20
                '        HHFtext.Text = HHF.ToString()
                '        HHFtext.Update()

                '        ' Calcula la altura de la hoja móvil para carpinterías AP94 con 1 hoja fija y 1 móvil

                '        HHM = HHF - 25
                '        HHMtext.Text = HHM.ToString()
                '        HHMtext.Update()

                '        ' Calcula el ancho del vidrio de la hoja fija para carpinterías AP94 con 1 hoja fija y 1 móvil

                '        AVF = AHF - 209
                '        AVFtext.Text = AVF.ToString()
                '        AVFtext.Update()

                '        ' Calcula la altura del vidrio de la hoja fija para carpinterías AP94 con 1 hoja fija y 1 móvil

                '        HVF = HHF - 164
                '        HVFtext.Text = HVF.ToString()
                '        HVFtext.Update()

                '        ' Calcula el ancho del vidrio de la hoja móvil para carpinterías AP94 con 1 hoja fija y 1 móvil

                '        AVM = AHM - 127
                '        AVMtext.Text = AVM.ToString()
                '        AVMtext.Update()

                '        ' Calcula la altura del vidrio de la hoja móvil para carpinterías AP94 con 1 hoja fija y 1 móvil

                '        HVM = HHM - 159
                '        HVMtext.Text = HVM.ToString()
                '        HVMtext.Update()

                '        ' Calculo las medidas de los componentes de la carpintería AP94

                '        e1 = HHF
                '        e2 = HHF - 46
                '        e3 = HHF - 41
                '        e4 = AHF - 46
                '        e5 = AHF - 59
                '        e6 = HHF - 146
                '        e7 = AHF - 53
                '        e8 = AHF - 190
                '        e9 = AHF - 195
                '        e10 = AHF - 190
                '        e11 = AHF - 235
                '        e12 = HHM - 41
                '        e13 = HHM - 54
                '        e14 = HHM - 54
                '        e15 = HHM - 54
                '        e16 = HHM - 141
                '        e17 = AHM - 17
                '        e18 = AHM - 108
                '        e19 = AHM - 113
                '        e20 = AHM - 108
                '        e21 = AHM - 153
                '        e22 = HHM - 41
                '        e23 = HHM - 41
                '        e24 = wh
                '        e25 = wh - AHF
                '        e26 = wh - AHF
                '        e27 = AHF
                '        e28 = wh
                '        e29 = wh

                '        e1text.Text = e1.ToString()
                '        e1text.Update()
                '        e2text.Text = e2.ToString()
                '        e2text.Update()
                '        e3text.Text = e3.ToString()
                '        e3text.Update()
                '        e4text.Text = e4.ToString()
                '        e4text.Update()
                '        e5text.Text = e5.ToString()
                '        e5text.Update()
                '        e6text.Text = e6.ToString()
                '        e6text.Update()
                '        e7text.Text = e7.ToString()
                '        e7text.Update()
                '        e8text.Text = e8.ToString()
                '        e8text.Update()
                '        e9text.Text = e9.ToString()
                '        e9text.Update()
                '        e10text.Text = e10.ToString()
                '        e10text.Update()
                '        e11text.Text = e11.ToString()
                '        e11text.Update()
                '        e12text.Text = e12.ToString()
                '        e12text.Update()
                '        e13text.Text = e13.ToString()
                '        e13text.Update()
                '        e14text.Text = e14.ToString()
                '        e14text.Update()
                '        e15text.Text = e15.ToString()
                '        e15text.Update()
                '        e16text.Text = e16.ToString()
                '        e16text.Update()
                '        e17text.Text = e17.ToString()
                '        e17text.Update()
                '        e18text.Text = e18.ToString()
                '        e18text.Update()
                '        e19text.Text = e19.ToString()
                '        e19text.Update()
                '        e20text.Text = e20.ToString()
                '        e20text.Update()
                '        e21text.Text = e21.ToString()
                '        e21text.Update()
                '        e22text.Text = e22.ToString()
                '        e22text.Update()
                '        e23text.Text = e23.ToString()
                '        e23text.Update()
                '        e24text.Text = e24.ToString()
                '        e24text.Update()
                '        e25text.Text = e25.ToString()
                '        e25text.Update()
                '        e26text.Text = e26.ToString()
                '        e26text.Update()
                '        e27text.Text = e27.ToString()
                '        e27text.Update()
                '        e28text.Text = e28.ToString()
                '        e28text.Update()
                '        e29text.Text = e29.ToString()
                '        e29text.Update()

                '        eu1 = 4
                '        eu2 = 2
                '        eu3 = 2
                '        eu4 = 2
                '        eu5 = 2
                '        eu6 = 4
                '        eu7 = 4
                '        eu8 = 2
                '        eu9 = 4
                '        eu10 = 2
                '        eu11 = 2
                '        eu12 = 2
                '        eu13 = 2
                '        eu14 = 2
                '        eu15 = 2
                '        eu16 = 4
                '        eu17 = 4
                '        eu18 = 2
                '        eu19 = 1
                '        eu20 = 2
                '        eu21 = 1
                '        eu22 = 1
                '        eu23 = 2
                '        eu24 = 1
                '        eu25 = 1
                '        eu26 = 2
                '        eu27 = 1
                '        eu28 = 1
                '        eu29 = 87

                '        eu1text.Text = e1.ToString()
                '        eu1text.Update()
                '        eu2text.Text = e2.ToString()
                '        eu2text.Update()
                '        eu3text.Text = e3.ToString()
                '        eu3text.Update()
                '        eu4text.Text = e4.ToString()
                '        eu4text.Update()
                '        eu5text.Text = e5.ToString()
                '        eu5text.Update()
                '        eu6text.Text = e6.ToString()
                '        eu6text.Update()
                '        eu7text.Text = e7.ToString()
                '        eu7text.Update()
                '        eu8text.Text = e8.ToString()
                '        eu8text.Update()
                '        eu9text.Text = e9.ToString()
                '        eu9text.Update()
                '        eu10text.Text = e10.ToString()
                '        eu10text.Update()
                '        eu11text.Text = e11.ToString()
                '        eu11text.Update()
                '        eu12text.Text = e12.ToString()
                '        eu12text.Update()
                '        eu13text.Text = e13.ToString()
                '        eu13text.Update()
                '        eu14text.Text = e14.ToString()
                '        eu14text.Update()
                '        eu15text.Text = e15.ToString()
                '        eu15text.Update()
                '        eu16text.Text = e16.ToString()
                '        eu16text.Update()
                '        eu17text.Text = e17.ToString()
                '        eu17text.Update()
                '        eu18text.Text = e18.ToString()
                '        eu18text.Update()
                '        eu19text.Text = e19.ToString()
                '        eu19text.Update()
                '        eu20text.Text = e20.ToString()
                '        eu20text.Update()
                '        eu21text.Text = e21.ToString()
                '        eu21text.Update()
                '        eu22text.Text = e22.ToString()
                '        eu22text.Update()
                '        eu23text.Text = e23.ToString()
                '        eu23text.Update()
                '        eu24text.Text = e24.ToString()
                '        eu24text.Update()
                '        eu25text.Text = e25.ToString()
                '        eu25text.Update()
                '        eu26text.Text = e26.ToString()
                '        eu26text.Update()
                '        eu27text.Text = e27.ToString()
                '        eu27text.Update()
                '        eu28text.Text = e28.ToString()
                '        eu28text.Update()
                '        eu29text.Text = e29.ToString()
                '        eu29text.Update()

                '        Panel3.Visible = False
                '        Panel4.Visible = False
                '        Panel5.Visible = False
                '        Panel6.Visible = False
                '        Label105.Visible = False
                '        AHMLtext.Visible = False
                '        HHMLtext.Visible = False
                '        Label106.Visible = False
                '        AVMLtext.Visible = False
                '        HVMLtext.Visible = False
                '        Label107.Visible = False
                '        Label108.Visible = False
                '        Label143.Visible = False
                '        Label146.Visible = False
                '        e22text.Visible = False
                '        e23text.Visible = False
                '        eu22text.Visible = False
                '        eu23text.Visible = False
                '        TabControl2.SelectedIndex = 4

                End If

        End While

        While (tb_pl.Text <> "" And bFlag = False)

            ' Cambio en el valor de bFlag para salir del While 

            bFlag = True

            '******************** Inicio del condicional para calcular medidas y perfilerías ******************************

            ' Al modificar el valor del Text Box relativo al Paso Libre se inicializa el condicional 
            ' La función Val hace que las variables definidas adopten el contenido numérico del Text Box

            wh = Val(tb_wh.Text)
            hh = Val(tb_hh.Text)
            pl = Val(tb_pl.Text)


            '****************** MI - CASO I *****************************************************************************************************************************

            If cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("MI") And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 2 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías MI con 2 hojas fijas y 2 móviles

                AHF = ((wh - pl) / 2)
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías MI con 2 hojas fijas y 2 móviles

                HHF = hh - 10
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías MI con 2 hojas fijas y 2 móviles

                AHM = (pl / 2) + solMI
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías MI con 2 hojas fijas y 2 móviles

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías MI con 2 hojas fijas y 2 móviles

                AVF = AHF - deshMI
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías MI con 2 hojas fijas y 2 móviles

                HVF = HHF - desvMI
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías MI con 2 hojas fijas y 2 móviles

                AVM = AHM - deshMI
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías MI con 2 hojas fijas y 2 móviles

                HVM = HHM - desvMI - 16
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calculo las medidas de los componentes de la carpintería MI

                a1 = HHF
                a2 = AHF - 90
                a3 = HHF + 8
                a4 = HHM - 16
                a5 = AHM - 90
                a6 = HHM - 16
                a7 = AHM
                a8 = pl
                If (AHF > AHM) Then
                    a9 = (2 * AHF) + pl
                ElseIf (AHF < AHM) Then
                    a9 = (2 * AHM) + pl
                ElseIf (AHF = AHM) Then
                    a9 = (2 * AHF) + pl
                End If
                a10 = a9

                a1text.Text = a1.ToString()
                a1text.Update()
                a2text.Text = a2.ToString()
                a2text.Update()
                a3text.Text = a3.ToString()
                a3text.Update()
                a4text.Text = a4.ToString()
                a4text.Update()
                a5text.Text = a5.ToString()
                a5text.Update()
                a6text.Text = a6.ToString()
                a6text.Update()
                a7text.Text = a7.ToString()
                a7text.Update()
                a8text.Text = a8.ToString()
                a8text.Update()
                a9text.Text = a9.ToString()
                a9text.Update()
                a10text.Text = a10.ToString()
                a10text.Update()

                au1 = 4
                au2 = 4
                au3 = 2
                au4 = 4
                au5 = 4
                au6 = 2
                au7 = 2
                au8 = 1
                au9 = 1
                au10 = 1

                au1text.Text = au1.ToString()
                au1text.Update()
                au2text.Text = au2.ToString()
                au2text.Update()
                au3text.Text = au3.ToString()
                au3text.Update()
                au4text.Text = au4.ToString()
                au4text.Update()
                au5text.Text = au5.ToString()
                au5text.Update()
                au6text.Text = au6.ToString()
                au6text.Update()
                au7text.Text = au7.ToString()
                au7text.Update()
                au8text.Text = au8.ToString()
                au8text.Update()
                au9text.Text = au9.ToString()
                au9text.Update()
                au10text.Text = au10.ToString()
                au10text.Update()

                Panel3.Visible = True
                Panel4.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                au17text.Visible = False
                a17text.Visible = False
                Label223.Visible = False
                TabControl2.SelectedIndex = 0

                '****************** MI - CASO II *****************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("MI") And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Móviles") Then

                'Calculo de la anchura y altura de las hojas móviles con carpintería MI y 2 hojas móviles

                AHM = (pl / 2) + solMI
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                'Calculo de la anchura y altura de los vidrios de las hojas móviles con carpintería MI y 2 hojas móviles

                AVM = AHM - deshMI
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                HVM = HHM - desvMI - 16
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Sacar mensaje de no aplicación de hojas fijas en este caso 

                AHFtext.Text = "No hay h.fijas"
                HHFtext.Text = "No hay h.fijas"
                AVFtext.Text = "No hay h.fijas"
                HVFtext.Text = "No hay h.fijas"

                ' Calculo las medidas de los componentes de la carpintería MI

                a4 = HHM - 16
                a5 = AHM - 90
                a6 = HHM - 16
                a7 = AHM
                a8 = pl
                If (AHF > AHM) Then
                    a9 = (2 * AHF) + pl
                ElseIf (AHF < AHM) Then
                    a9 = (2 * AHM) + pl
                ElseIf (AHF = AHM) Then
                    a9 = (2 * AHM) + pl
                End If
                a10 = a9

                a1text.Text = "No hay h.fijas"
                a1text.Update()
                a2text.Text = "No hay h.fijas"
                a2text.Update()
                a3text.Text = "No hay h.fijas"
                a3text.Update()
                a4text.Text = a4.ToString()
                a4text.Update()
                a5text.Text = a5.ToString()
                a5text.Update()
                a6text.Text = a6.ToString()
                a6text.Update()
                a7text.Text = a7.ToString()
                a7text.Update()
                a8text.Text = a8.ToString()
                a8text.Update()
                a9text.Text = a9.ToString()
                a9text.Update()
                a10text.Text = a10.ToString()
                a10text.Update()

                au4 = 4
                au5 = 4
                au6 = 2
                au7 = 2
                au8 = 1
                au9 = 1
                au10 = 1

                au1text.Text = "No hay h.fijas"
                au1text.Update()
                au2text.Text = "No hay h.fijas"
                au2text.Update()
                au3text.Text = "No hay h.fijas"
                au3text.Update()
                au4text.Text = au4.ToString()
                au4text.Update()
                au5text.Text = au5.ToString()
                au5text.Update()
                au6text.Text = au6.ToString()
                au6text.Update()
                au7text.Text = au7.ToString()
                au7text.Update()
                au8text.Text = au8.ToString()
                au8text.Update()
                au9text.Text = au9.ToString()
                au9text.Update()
                au10text.Text = au10.ToString()
                au10text.Update()

                Panel3.Visible = True
                Panel4.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                au17text.Visible = False
                a17text.Visible = False
                Label223.Visible = False
                TabControl2.SelectedIndex = 0

                '****************** MI - CASO III ****************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("MI") And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 1 Hoja Móvil") Then

                ' Calcula el ancho de la hoja fija para carpinterías MI con 1 hoja fija y 1 móvil

                AHF = wh - pl
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías MI con 1 hoja fija y 1 móvil

                HHF = hh - 10
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías MI con 1 hoja fija y 1 móvil

                AHM = pl + solMI
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías MI con 1 hoja fija y 1 móvil

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías MI con 1 hoja fija y 1 móvil

                AVF = AHF - deshMI
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías MI con 1 hoja fija y 1 móvil

                HVF = HHF - desvMI
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías MI con 1 hoja fija y 1 móvil

                AVM = AHM - deshMI
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías MI con 1 hoja fija y 1 móvil

                HVM = HHM - desvMI - 16
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calculo las medidas de los componentes de la carpintería MI

                a1 = HHF
                a2 = AHF - 90
                a3 = HHF + 8
                a4 = HHM - 16
                a5 = AHM - 90
                a6 = HHM - 16
                a7 = AHM
                a8 = pl
                If (AHF > AHM) Then
                    a9 = AHF + pl
                ElseIf (AHF < AHM) Then
                    a9 = AHM + pl
                ElseIf (AHF = AHM) Then
                    a9 = AHF + pl
                End If
                a10 = a9

                a1text.Text = a1.ToString()
                a1text.Update()
                a2text.Text = a2.ToString()
                a2text.Update()
                a3text.Text = a3.ToString()
                a3text.Update()
                a4text.Text = a4.ToString()
                a4text.Update()
                a5text.Text = a5.ToString()
                a5text.Update()
                a6text.Text = a6.ToString()
                a6text.Update()
                a7text.Text = a7.ToString()
                a7text.Update()
                a8text.Text = a8.ToString()
                a8text.Update()
                a9text.Text = a9.ToString()
                a9text.Update()
                a10text.Text = a10.ToString()
                a10text.Update()

                au1 = 2
                au2 = 2
                au3 = 1
                au4 = 2
                au5 = 2
                au6 = 1
                au7 = 1
                au8 = 1
                au9 = 1
                au10 = 1

                au1text.Text = au1.ToString()
                au1text.Update()
                au2text.Text = au2.ToString()
                au2text.Update()
                au3text.Text = au3.ToString()
                au3text.Update()
                au4text.Text = au4.ToString()
                au4text.Update()
                au5text.Text = au5.ToString()
                au5text.Update()
                au6text.Text = au6.ToString()
                au6text.Update()
                au7text.Text = au7.ToString()
                au7text.Update()
                au8text.Text = au8.ToString()
                au8text.Update()
                au9text.Text = au9.ToString()
                au9text.Update()
                au10text.Text = au10.ToString()
                au10text.Update()

                Panel3.Visible = True
                Panel4.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                au17text.Visible = False
                a17text.Visible = False
                Label223.Visible = False
                TabControl2.SelectedIndex = 0

                '****************** MI - CASO IV ****************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("MI") And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Móvil") Then

                'Calculo de la anchura y altura de las hojas móviles con carpintería MI y 2 hojas móviles

                AHM = pl + solMI
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                'Calculo de la anchura y altura de los vidrios de las hojas móviles con carpintería MI y 2 hojas móviles

                AVM = AHM - deshMI
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                HVM = HHM - desvMI - 16
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Sacar mensaje de no aplicación de hojas fijas en este caso 

                AHFtext.Text = "No hay h.fijas"
                HHFtext.Text = "No hay h.fijas"
                AVFtext.Text = "No hay h.fijas"
                HVFtext.Text = "No hay h.fijas"

                ' Calculo las medidas de los componentes de la carpintería MI

                a4 = HHM - 16
                a5 = AHM - 90
                a6 = HHM - 16
                a7 = AHM
                a8 = pl
                If (AHF > AHM) Then
                    a9 = AHF + pl
                ElseIf (AHF < AHM) Then
                    a9 = AHM + pl
                ElseIf (AHF = AHM) Then
                    a9 = AHM + pl
                End If
                a10 = a9

                a1text.Text = "No hay h.fijas"
                a1text.Update()
                a2text.Text = "No hay h.fijas"
                a2text.Update()
                a3text.Text = "No hay h.fijas"
                a3text.Update()
                a4text.Text = a4.ToString()
                a4text.Update()
                a5text.Text = a5.ToString()
                a5text.Update()
                a6text.Text = a6.ToString()
                a6text.Update()
                a7text.Text = a7.ToString()
                a7text.Update()
                a8text.Text = a8.ToString()
                a8text.Update()
                a9text.Text = a9.ToString()
                a9text.Update()
                a10text.Text = a10.ToString()
                a10text.Update()

                au4 = 2
                au5 = 2
                au6 = 1
                au7 = 1
                au8 = 1
                au9 = 1
                au10 = 1

                au1text.Text = "No hay h.fijas"
                au1text.Update()
                au2text.Text = "No hay h.fijas"
                au2text.Update()
                au3text.Text = "No hay h.fijas"
                au3text.Update()
                au4text.Text = au4.ToString()
                au4text.Update()
                au5text.Text = au5.ToString()
                au5text.Update()
                au6text.Text = au6.ToString()
                au6text.Update()
                au7text.Text = au7.ToString()
                au7text.Update()
                au8text.Text = au8.ToString()
                au8text.Update()
                au9text.Text = au9.ToString()
                au9text.Update()
                au10text.Text = au10.ToString()
                au10text.Update()

                Panel3.Visible = True
                Panel4.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                au17text.Visible = False
                a17text.Visible = False
                Label223.Visible = False
                TabControl2.SelectedIndex = 0

                '****************** MI - CASO V ****************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("MI") And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 4 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                AHF = (wh - pl) / 2
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                HHF = hh - 10
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                AHM = (pl / 4) + solMI
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                HHM = hh - 5
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho de la hoja móvil lenta para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                AHML = AHM
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                ' Calcula la altura de la hoja móvil lenta para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                HHML = hh
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                AVF = AHF - deshMI
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                HVF = HHF - desvMI
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                AVM = AHM - deshMI
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                HVM = HHM - desvMI - 16
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil lenta para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                AVML = AHML - deshMI - 10
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil lenta para carpinterías MI con 2 Hojas Fijas + 4 Hojas Móviles

                HVML = HHML - desvMI - 16
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Calculo las medidas de los componentes de la carpintería MI

                a1 = HHF
                a2 = AHF - 90
                a3 = HHF + 8
                a4 = HHM - 16
                a5 = AHM - 90
                a6 = HHM - 16
                a7 = AHM
                a8 = pl
                If (AHF > AHM) Then
                    a9 = (2 * AHF) + pl
                ElseIf (AHF < AHM) Then
                    a9 = (2 * AHM) + pl
                ElseIf (AHF = AHM) Then
                    a9 = (2 * AHF) + pl
                End If
                a10 = a9
                a11 = HHML - 16
                a12 = AHML - 100
                a13 = HHML - 16
                a14 = AHML
                a15 = HHML - 16
                a17 = HHF + 8

                a1text.Text = a1.ToString()
                a1text.Update()
                a2text.Text = a2.ToString()
                a2text.Update()
                a3text.Text = "No lleva portafelpudos lateral"
                a3text.Update()
                a4text.Text = a4.ToString()
                a4text.Update()
                a5text.Text = a5.ToString()
                a5text.Update()
                a6text.Text = a6.ToString()
                a6text.Update()
                a7text.Text = a7.ToString()
                a7text.Update()
                a8text.Text = a8.ToString()
                a8text.Update()
                a9text.Text = a9.ToString()
                a9text.Update()
                a10text.Text = a10.ToString()
                a10text.Update()
                a11text.Text = a11.ToString()
                a11text.Update()
                a12text.Text = a12.ToString()
                a12text.Update()
                a13text.Text = a13.ToString()
                a13text.Update()
                a14text.Text = a14.ToString()
                a14text.Update()
                a15text.Text = a15.ToString()
                a15text.Update()
                a17text.Text = a17.ToString()
                a17text.Update()

                au1 = 4
                au2 = 4
                'au3 = 2
                au4 = 4
                au5 = 4
                au6 = 2
                au7 = 2
                au8 = 1
                au9 = 1
                au10 = 1
                au11 = 4
                au12 = 4
                au13 = 2
                au14 = 2
                au15 = 2
                au17 = 2

                au1text.Text = au1.ToString()
                au1text.Update()
                au2text.Text = au2.ToString()
                au2text.Update()
                au3text.Text = "No lleva portafelpudos lateral"
                au3text.Update()
                au4text.Text = au4.ToString()
                au4text.Update()
                au5text.Text = au5.ToString()
                au5text.Update()
                au6text.Text = au6.ToString()
                au6text.Update()
                au7text.Text = au7.ToString()
                au7text.Update()
                au8text.Text = au8.ToString()
                au8text.Update()
                au9text.Text = au9.ToString()
                au9text.Update()
                au10text.Text = au10.ToString()
                au10text.Update()
                au11text.Text = au11.ToString()
                au11text.Update()
                au12text.Text = au12.ToString()
                au12text.Update()
                au13text.Text = au13.ToString()
                au13text.Update()
                au14text.Text = au14.ToString()
                au14text.Update()
                au15text.Text = au15.ToString()
                au15text.Update()
                au17text.Text = au17.ToString()
                au17text.Update()

                Panel3.Visible = True
                Panel4.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Panel11.Visible = True
                Label189.Visible = True
                Label222.Visible = False
                au16text.Visible = False
                a16text.Visible = False
                TabControl2.SelectedIndex = 0

                '****************** MI - CASO VI **************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("MI") And cmbconf.SelectedIndex = cmbconf.FindStringExact("4 Hojas Móviles") Then

                'Calculo de la anchura y altura de las hojas móviles con carpintería MI y 4 hojas móviles

                AHM = (pl / 4) + solMI
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                HHM = hh - 5
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                'Calculo de la anchura y altura de las hojas móviles Lentas con carpintería MI y 4 hojas móviles

                AHML = AHM
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                HHML = hh
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                'Calculo de la anchura y altura de los vidrios de las hojas móviles con carpintería MI y 4 hojas móviles

                AVM = AHM - deshMI
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                HVM = HHM - desvMI - 21
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                'Calculo de la anchura y altura de los vidrios de las hojas móviles lentas con carpintería MI y 4 hojas móviles

                AVML = AHML - deshMI - 10
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                HVML = HHML - desvMI - 16
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Sacar mensaje de no aplicación de hojas fijas en este caso 

                AHFtext.Text = "No hay h.fijas"
                HHFtext.Text = "No hay h.fijas"
                AVFtext.Text = "No hay h.fijas"
                HVFtext.Text = "No hay h.fijas"

                ' Calculo las medidas de los componentes de la carpintería MI

                a4 = HHM - 16
                a5 = AHM - 90
                a6 = HHM - 16
                a7 = AHM
                a8 = pl
                If (AHF > AHM) Then
                    a9 = AHF + pl
                ElseIf (AHF < AHM) Then
                    a9 = (2 * AHM) + pl
                ElseIf (AHF = AHM) Then
                    a9 = (2 * AHM) + pl
                End If
                a10 = a9
                a11 = HHML - 16
                a12 = AHML - 100
                a13 = HHML - 16
                a14 = AHML
                a15 = HHML - 16
                a16 = hh


                a1text.Text = "No hay h.fijas"
                a1text.Update()
                a2text.Text = "No hay h.fijas"
                a2text.Update()
                a3text.Text = "No hay h.fijas"
                a3text.Update()
                a4text.Text = a4.ToString()
                a4text.Update()
                a5text.Text = a5.ToString()
                a5text.Update()
                a6text.Text = a6.ToString()
                a6text.Update()
                a7text.Text = a7.ToString()
                a7text.Update()
                a8text.Text = a8.ToString()
                a8text.Update()
                a9text.Text = a9.ToString()
                a9text.Update()
                a10text.Text = a10.ToString()
                a10text.Update()
                a11text.Text = a11.ToString()
                a11text.Update()
                a12text.Text = a12.ToString()
                a12text.Update()
                a13text.Text = a13.ToString()
                a13text.Update()
                a14text.Text = a14.ToString()
                a14text.Update()
                a15text.Text = a15.ToString()
                a15text.Update()
                a16text.Text = a16.ToString()
                a16text.Update()
                a17text.Text = "No hay h.fijas"
                a17text.Update()

                au4 = 4
                au5 = 4
                au6 = 2
                au7 = 2
                au8 = 1
                au9 = 1
                au10 = 1
                au11 = 4
                au12 = 4
                au13 = 2
                au14 = 1
                au15 = 1
                au16 = 2


                au1text.Text = "No hay h.fijas"
                au1text.Update()
                au2text.Text = "No hay h.fijas"
                au2text.Update()
                au3text.Text = "No hay h.fijas"
                au3text.Update()
                au4text.Text = au4.ToString()
                au4text.Update()
                au5text.Text = au5.ToString()
                au5text.Update()
                au6text.Text = au6.ToString()
                au6text.Update()
                au7text.Text = au7.ToString()
                au7text.Update()
                au8text.Text = au8.ToString()
                au8text.Update()
                au9text.Text = au9.ToString()
                au9text.Update()
                au10text.Text = au10.ToString()
                au10text.Update()
                au11text.Text = au11.ToString()
                au11text.Update()
                au12text.Text = au12.ToString()
                au12text.Update()
                au13text.Text = au13.ToString()
                au13text.Update()
                au14text.Text = au14.ToString()
                au14text.Update()
                au15text.Text = au15.ToString()
                au15text.Update()
                au16text.Text = au16.ToString()
                au16text.Update()
                au17text.Text = "No hay h.fijas"
                au17text.Update()

                Panel3.Visible = True
                Panel4.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Panel11.Visible = True
                Label189.Visible = True
                Label222.Visible = True
                a16text.Visible = True
                au16text.Visible = True
                TabControl2.SelectedIndex = 0

                '****************** MI - CASO VII *************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("MI") And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 2 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                AHF = wh - pl
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                HHF = hh - 10
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                AHM = (pl / 2) + solMI
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                HHM = hh - 5
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho de la hoja móvil lenta para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                AHML = (pl / 2) + solMI
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                ' Calcula la altura de la hoja móvil lenta para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                HHML = hh
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                AVF = AHF - deshMI
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                HVF = HHF - desvMI
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                AVM = AHM - deshMI
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                HVM = HHM - desvMI - 16
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil lenta para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                AVML = AHML - deshMI - 10
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil lenta para carpinterías MI con 1 Hoja Fija + 2 Hojas Móviles

                HVML = HHML - desvMI - 16
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Calculo las medidas de los componentes de la carpintería MI

                a1 = HHF
                a2 = AHF - 90
                a3 = HHF + 8
                a4 = HHM - 16
                a5 = AHM - 90
                a6 = HHM - 16
                a7 = AHM
                a8 = pl
                If (AHF > AHM) Then
                    a9 = (AHF) + pl
                ElseIf (AHF < AHM) Then
                    a9 = (AHM) + pl
                ElseIf (AHF = AHM) Then
                    a9 = (AHF) + pl
                End If
                a10 = a9
                a11 = HHML - 16
                a12 = AHML - 100
                a13 = HHML - 16
                a14 = AHML
                a15 = HHML - 16
                a17 = HHF + 8

                a1text.Text = a1.ToString()
                a1text.Update()
                a2text.Text = a2.ToString()
                a2text.Update()
                a3text.Text = "No lleva portafelpudos lateral"
                a3text.Update()
                a4text.Text = a4.ToString()
                a4text.Update()
                a5text.Text = a5.ToString()
                a5text.Update()
                a6text.Text = a6.ToString()
                a6text.Update()
                a7text.Text = a7.ToString()
                a7text.Update()
                a8text.Text = a8.ToString()
                a8text.Update()
                a9text.Text = a9.ToString()
                a9text.Update()
                a10text.Text = a10.ToString()
                a10text.Update()
                a11text.Text = a11.ToString()
                a11text.Update()
                a12text.Text = a12.ToString()
                a12text.Update()
                a13text.Text = a13.ToString()
                a13text.Update()
                a14text.Text = a14.ToString()
                a14text.Update()
                a15text.Text = a15.ToString()
                a15text.Update()
                a17text.Text = a17.ToString()
                a17text.Update()

                au1 = 2
                au2 = 2
                au3 = 1
                au4 = 2
                au5 = 2
                au6 = 1
                au7 = 1
                au8 = 1
                au9 = 1
                au10 = 1
                au11 = 2
                au12 = 2
                au13 = 1
                au14 = 1
                au15 = 1
                au17 = 1

                au1text.Text = au1.ToString()
                au1text.Update()
                au2text.Text = au2.ToString()
                au2text.Update()
                au3text.Text = "No lleva portafelpudos lateral"
                au3text.Update()
                au4text.Text = au4.ToString()
                au4text.Update()
                au5text.Text = au5.ToString()
                au5text.Update()
                au6text.Text = au6.ToString()
                au6text.Update()
                au7text.Text = au7.ToString()
                au7text.Update()
                au8text.Text = au8.ToString()
                au8text.Update()
                au9text.Text = au9.ToString()
                au9text.Update()
                au10text.Text = au10.ToString()
                au10text.Update()
                au11text.Text = au11.ToString()
                au11text.Update()
                au12text.Text = au12.ToString()
                au12text.Update()
                au13text.Text = au13.ToString()
                au13text.Update()
                au14text.Text = au14.ToString()
                au14text.Update()
                au15text.Text = au15.ToString()
                au15text.Update()
                au17text.Text = au17.ToString()
                au17text.Update()

                Panel3.Visible = True
                Panel4.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Panel11.Visible = True
                Label189.Visible = True
                Label222.Visible = False
                au16text.Visible = False
                a16text.Visible = False
                TabControl2.SelectedIndex = 0

                '****************** MI - CASO VIII **************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("MI") And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Móviles TES") Then

                'Calculo de la anchura y altura de las hojas móviles con carpintería MI y 2 hojas móviles

                AHM = (pl / 2) + solMI
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                HHM = hh - 5
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                'Calculo de la anchura y altura de las hojas móviles Lentas con carpintería MI y 2 hojas móviles

                AHML = AHM
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                HHML = hh
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                'Calculo de la anchura y altura de los vidrios de las hojas móviles con carpintería MI y 2 hojas móviles

                AVM = AHM - deshMI
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                HVM = HHM - desvMI - 16
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                'Calculo de la anchura y altura de los vidrios de las hojas móviles lentas con carpintería MI y 2 hojas móviles

                AVML = AHML - deshMI - 10
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                HVML = HHML - desvMI - 16
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Sacar mensaje de no aplicación de hojas fijas en este caso 

                AHFtext.Text = "No hay h.fijas"
                HHFtext.Text = "No hay h.fijas"
                AVFtext.Text = "No hay h.fijas"
                HVFtext.Text = "No hay h.fijas"

                ' Calculo las medidas de los componentes de la carpintería MI

                a4 = HHM - 16
                a5 = AHM - 90
                a6 = HHM - 16
                a7 = AHM
                a8 = pl
                If (AHF > AHM) Then
                    a9 = AHF + pl
                ElseIf (AHF < AHM) Then
                    a9 = AHM + pl
                ElseIf (AHF = AHM) Then
                    a9 = AHM + pl
                End If
                a10 = a9
                a11 = HHML - 16
                a12 = AHML - 100
                a13 = HHML - 16
                a14 = AHML
                a15 = HHML - 16
                a16 = hh

                a1text.Text = "No hay h.fijas"
                a1text.Update()
                a2text.Text = "No hay h.fijas"
                a2text.Update()
                a3text.Text = "No hay h.fijas"
                a3text.Update()
                a4text.Text = a4.ToString()
                a4text.Update()
                a5text.Text = a5.ToString()
                a5text.Update()
                a6text.Text = a6.ToString()
                a6text.Update()
                a7text.Text = a7.ToString()
                a7text.Update()
                a8text.Text = a8.ToString()
                a8text.Update()
                a9text.Text = a9.ToString()
                a9text.Update()
                a10text.Text = a10.ToString()
                a10text.Update()
                a11text.Text = a11.ToString()
                a11text.Update()
                a12text.Text = a12.ToString()
                a12text.Update()
                a13text.Text = a13.ToString()
                a13text.Update()
                a14text.Text = a14.ToString()
                a14text.Update()
                a15text.Text = a15.ToString()
                a15text.Update()
                a16text.Text = a16.ToString()
                a16text.Update()
                a17text.Text = "No hay h.fijas"
                a17text.Update()

                au4 = 2
                au5 = 2
                au6 = 1
                au7 = 1
                au8 = 1
                au9 = 1
                au10 = 1
                au11 = 2
                au12 = 2
                au13 = 1
                au14 = 1
                au15 = 1
                au16 = 1

                au1text.Text = "No hay h.fijas"
                au1text.Update()
                au2text.Text = "No hay h.fijas"
                au2text.Update()
                au3text.Text = "No hay h.fijas"
                au3text.Update()
                au4text.Text = au4.ToString()
                au4text.Update()
                au5text.Text = au5.ToString()
                au5text.Update()
                au6text.Text = au6.ToString()
                au6text.Update()
                au7text.Text = au7.ToString()
                au7text.Update()
                au8text.Text = au8.ToString()
                au8text.Update()
                au9text.Text = au9.ToString()
                au9text.Update()
                au10text.Text = au10.ToString()
                au10text.Update()
                au11text.Text = au11.ToString()
                au11text.Update()
                au12text.Text = au12.ToString()
                au12text.Update()
                au13text.Text = au13.ToString()
                au13text.Update()
                au14text.Text = au14.ToString()
                au14text.Update()
                au15text.Text = au15.ToString()
                au15text.Update()
                au16text.Text = au16.ToString()
                au16text.Update()
                au17text.Text = "No hay h.fijas"
                au17text.Update()

                Panel3.Visible = True
                Panel4.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Panel11.Visible = True
                Label189.Visible = True
                Label222.Visible = True
                a16text.Visible = True
                au16text.Visible = True
                TabControl2.SelectedIndex = 0

                '****************** FG - CASO I *****************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Full Glass") And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 2 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías FG con 2 hojas fijas y 2 móviles

                AHF = ((wh - pl) / 2)
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías FG con 2 hojas fijas y 2 móviles

                HHF = hh - 10
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías FG con 2 hojas fijas y 2 móviles

                AHM = (pl / 2) + solFG
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías FG con 2 hojas fijas y 2 móviles

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías FG con 2 hojas fijas y 2 móviles

                AVF = AHF - deshFG
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías FG con 2 hojas fijas y 2 móviles

                HVF = HHF - desvFG
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías FG con 2 hojas fijas y 2 móviles

                AVM = AHM - deshFG
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías FG con 2 hojas fijas y 2 móviles

                HVM = HHM - desvFG
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calculo las medidas de los componentes de la carpintería FG

                b1 = HHF
                b2 = HHF
                b3 = AHF - 46
                b4 = HHM
                b5 = HHM
                b6 = AHM - 46
                b7 = AHM - 46
                b8 = pl
                b9 = wh
                If (AHF > AHM) Then
                    b10 = (2 * AHF) + pl
                ElseIf (AHF < AHM) Then
                    b10 = (2 * AHM) + pl
                ElseIf (AHF = AHM) Then
                    b10 = (2 * AHF) + pl
                End If
                b11 = b10

                b1text.Text = b1.ToString()
                b1text.Update()
                b2text.Text = b2.ToString()
                b2text.Update()
                b3text.Text = b3.ToString()
                b3text.Update()
                b4text.Text = b4.ToString()
                b4text.Update()
                b5text.Text = b5.ToString()
                b5text.Update()
                b6text.Text = b6.ToString()
                b6text.Update()
                b7text.Text = b7.ToString()
                b7text.Update()
                b8text.Text = b8.ToString()
                b8text.Update()
                b9text.Text = b9.ToString()
                b9text.Update()
                b10text.Text = b10.ToString()
                b10text.Update()
                b11text.Text = b11.ToString()
                b11text.Update()

                bu1 = 2
                bu2 = 2
                bu3 = 4
                bu4 = 2
                bu5 = 2
                bu6 = 2
                bu7 = 2
                bu8 = 1
                bu9 = 1
                bu10 = 1
                bu11 = 1

                bu1text.Text = bu1.ToString()
                bu1text.Update()
                bu2text.Text = bu2.ToString()
                bu2text.Update()
                bu3text.Text = bu3.ToString()
                bu3text.Update()
                bu4text.Text = bu4.ToString()
                bu4text.Update()
                bu5text.Text = bu5.ToString()
                bu5text.Update()
                bu6text.Text = bu6.ToString()
                bu6text.Update()
                bu7text.Text = bu7.ToString()
                bu7text.Update()
                bu8text.Text = bu8.ToString()
                bu8text.Update()
                bu9text.Text = bu9.ToString()
                bu9text.Update()
                bu10text.Text = bu10.ToString()
                bu10text.Update()
                bu11text.Text = bu11.ToString()
                bu11text.Update()

                Panel4.Visible = True
                Panel3.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                TabControl2.SelectedIndex = 1

                '****************** FG - CASO II *****************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Full Glass") And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Móviles") Then

                'Calculo de la anchura y altura de las hojas móviles con carpintería FG y 2 hojas móviles

                AHM = (pl / 2) + solFG
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                'Calculo de la anchura y altura de los vidrios de las hojas móviles con carpintería FG y 2 hojas móviles

                AVM = AHM - deshFG
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                HVM = HHM - desvFG
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Sacar mensaje de no aplicación de hojas fijas en este caso 

                AHFtext.Text = "No hay h.fijas"
                HHFtext.Text = "No hay h.fijas"
                AVFtext.Text = "No hay h.fijas"
                HVFtext.Text = "No hay h.fijas"

                ' Calculo las medidas de los componentes de la carpintería FG

                b4 = HHM
                b5 = HHM
                b6 = AHM - 46
                b7 = AHM - 46
                b8 = pl
                b9 = wh
                If (AHF > AHM) Then
                    b10 = (2 * AHF) + pl
                ElseIf (AHF < AHM) Then
                    b10 = (2 * AHM) + pl
                ElseIf (AHF = AHM) Then
                    b10 = 2 * AHM + pl
                End If
                b11 = b10

                b1text.Text = "No hay h.fijas"
                b1text.Update()
                b2text.Text = "No hay h.fijas"
                b2text.Update()
                b3text.Text = "No hay h.fijas"
                b3text.Update()
                b4text.Text = b4.ToString()
                b4text.Update()
                b5text.Text = b5.ToString()
                b5text.Update()
                b6text.Text = b6.ToString()
                b6text.Update()
                b7text.Text = b7.ToString()
                b7text.Update()
                b8text.Text = b8.ToString()
                b8text.Update()
                b9text.Text = b9.ToString()
                b9text.Update()
                b10text.Text = b10.ToString()
                b10text.Update()
                b11text.Text = b11.ToString()
                b11text.Update()

                bu4 = 2
                bu5 = 2
                bu6 = 2
                bu7 = 2
                bu8 = 1
                bu9 = 1
                bu10 = 1
                bu11 = 1

                bu1text.Text = "No hay h.fijas"
                bu1text.Update()
                bu2text.Text = "No hay h.fijas"
                bu2text.Update()
                bu3text.Text = "No hay h.fijas"
                bu3text.Update()
                bu4text.Text = bu4.ToString()
                bu4text.Update()
                bu5text.Text = bu5.ToString()
                bu5text.Update()
                bu6text.Text = bu6.ToString()
                bu6text.Update()
                bu7text.Text = bu7.ToString()
                bu7text.Update()
                bu8text.Text = bu8.ToString()
                bu8text.Update()
                bu9text.Text = bu9.ToString()
                bu9text.Update()
                bu10text.Text = bu10.ToString()
                bu10text.Update()
                bu11text.Text = bu11.ToString()
                bu11text.Update()

                Panel4.Visible = True
                Panel3.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                TabControl2.SelectedIndex = 1

                '****************** FG - CASO III ****************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Full Glass") And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 1 Hoja Móvil") Then

                ' Calcula el ancho de la hoja fija para carpinterías FG con 1 hoja fija y 1 móvil

                AHF = wh - pl
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías FG con 1 hoja fija y 1 móvil

                HHF = hh - 10
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías FG con 1 hoja fija y 1 móvil

                AHM = pl + solFG
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías FG con 1 hoja fija y 1 móvil

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías FG con 1 hoja fija y 1 móvil

                AVF = AHF - deshFG
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías FG con 1 hoja fija y 1 móvil

                HVF = HHF - desvFG
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías FG con 1 hoja fija y 1 móvil

                AVM = AHM - deshFG
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías FG con 1 hoja fija y 1 móvil

                HVM = HHM - desvFG
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calculo las medidas de los componentes de la carpintería FG

                b1 = HHF
                b2 = HHF
                b3 = AHF - 46
                b4 = HHM
                b5 = HHM
                b6 = AHM - 46
                b7 = AHM - 46
                b8 = pl
                b9 = wh
                If (AHF > AHM) Then
                    b10 = (AHF) + pl
                ElseIf (AHF < AHM) Then
                    b10 = (AHM) + pl
                ElseIf (AHF = AHM) Then
                    b10 = AHF + pl
                End If
                b11 = b10

                b1text.Text = b1.ToString()
                b1text.Update()
                b2text.Text = b2.ToString()
                b2text.Update()
                b3text.Text = b3.ToString()
                b3text.Update()
                b4text.Text = b4.ToString()
                b4text.Update()
                b5text.Text = b5.ToString()
                b5text.Update()
                b6text.Text = b6.ToString()
                b6text.Update()
                b7text.Text = b7.ToString()
                b7text.Update()
                b8text.Text = b8.ToString()
                b8text.Update()
                b9text.Text = b9.ToString()
                b9text.Update()
                b10text.Text = b10.ToString()
                b10text.Update()
                b11text.Text = b11.ToString()
                b11text.Update()

                bu1 = 1
                bu2 = 1
                bu3 = 2
                bu4 = 1
                bu5 = 1
                bu6 = 1
                bu7 = 1
                bu8 = 1
                bu9 = 1
                bu10 = 1
                bu11 = 1

                bu1text.Text = bu1.ToString()
                bu1text.Update()
                bu2text.Text = bu2.ToString()
                bu2text.Update()
                bu3text.Text = bu3.ToString()
                bu3text.Update()
                bu4text.Text = bu4.ToString()
                bu4text.Update()
                bu5text.Text = bu5.ToString()
                bu5text.Update()
                bu6text.Text = bu6.ToString()
                bu6text.Update()
                bu7text.Text = bu7.ToString()
                bu7text.Update()
                bu8text.Text = bu8.ToString()
                bu8text.Update()
                bu9text.Text = bu9.ToString()
                bu9text.Update()
                bu10text.Text = bu10.ToString()
                bu10text.Update()
                bu11text.Text = bu11.ToString()
                bu11text.Update()

                Panel4.Visible = True
                Panel3.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                TabControl2.SelectedIndex = 1

                '****************** FG - CASO IV ****************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Full Glass") And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Móvil") Then

                'Calculo de la anchura y altura de las hojas móviles con carpintería FG y 1 hoja móvil

                AHM = pl + solFG
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                'Calculo de la anchura y altura de los vidrios de las hojas móviles con carpintería FG y 1 hoja móvil

                AVM = AHM - deshFG
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                HVM = HHM - desvFG
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Sacar mensaje de no aplicación de hojas fijas en este caso 

                AHFtext.Text = "No hay h.fijas"
                HHFtext.Text = "No hay h.fijas"
                AVFtext.Text = "No hay h.fijas"
                HVFtext.Text = "No hay h.fijas"

                ' Calculo las medidas de los componentes de la carpintería FG

                b4 = HHM
                b5 = HHM
                b6 = AHM - 46
                b7 = AHM - 46
                b8 = pl
                b9 = wh
                If (AHF > AHM) Then
                    b10 = AHF + pl
                ElseIf (AHF < AHM) Then
                    b10 = AHM + pl
                ElseIf (AHF = AHM) Then
                    b10 = AHM + pl
                End If
                b11 = b10

                b1text.Text = "No hay h.fijas"
                b1text.Update()
                b2text.Text = "No hay h.fijas"
                b2text.Update()
                b3text.Text = "No hay h.fijas"
                b3text.Update()
                b4text.Text = b4.ToString()
                b4text.Update()
                b5text.Text = b5.ToString()
                b5text.Update()
                b6text.Text = b6.ToString()
                b6text.Update()
                b7text.Text = b7.ToString()
                b7text.Update()
                b8text.Text = b8.ToString()
                b8text.Update()
                b9text.Text = b9.ToString()
                b9text.Update()
                b10text.Text = b10.ToString()
                b10text.Update()
                b11text.Text = b11.ToString()
                b11text.Update()

                bu4 = 1
                bu5 = 1
                bu6 = 1
                bu7 = 1
                bu8 = 1
                bu9 = 1
                bu10 = 1
                bu11 = 1

                bu1text.Text = "No hay h.fijas"
                bu1text.Update()
                bu2text.Text = "No hay h.fijas"
                bu2text.Update()
                bu3text.Text = "No hay h.fijas"
                bu3text.Update()
                bu4text.Text = bu4.ToString()
                bu4text.Update()
                bu5text.Text = bu5.ToString()
                bu5text.Update()
                bu6text.Text = bu6.ToString()
                bu6text.Update()
                bu7text.Text = bu7.ToString()
                bu7text.Update()
                bu8text.Text = bu8.ToString()
                bu8text.Update()
                bu9text.Text = bu9.ToString()
                bu9text.Update()
                bu10text.Text = bu10.ToString()
                bu10text.Update()
                bu11text.Text = bu11.ToString()
                bu11text.Update()

                Panel4.Visible = True
                Panel3.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                TabControl2.SelectedIndex = 1

                '****************** FG - CASO V ****************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Full Glass") And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 4 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                AHF = (wh - pl) / 2
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                HHF = hh - 10
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                AHM = (pl / 4) + solFG
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                HHM = hh - 15
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho de la hoja móvil lenta para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                AHML = AHM
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                ' Calcula la altura de la hoja móvil lenta para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                HHML = hh - 10
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                AVF = AHF - deshFG
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                HVF = HHF - desvFG
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                AVM = AHM - deshFG
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                HVM = HHM - desvFG - 5
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil lenta para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                AVML = AHML - ((3 * deshFG) / 2)
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil lenta para carpinterías FG con 2 Hojas Fijas + 4 Hojas Móviles

                HVML = HHML - desvFG
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Calculo las medidas de los componentes de la carpintería FG

                b1 = HHF
                b2 = HHF
                b3 = AHF - 46
                b4 = HHM
                b5 = HHM
                b6 = AHM - 46
                b7 = AHM - 46
                b8 = pl
                b9 = wh
                If (AHF > AHM) Then
                    b10 = (2 * AHF) + pl
                ElseIf (AHF < AHM) Then
                    b10 = (2 * AHM) + pl
                ElseIf (AHF = AHM) Then
                    b10 = (2 * AHF) + pl
                End If
                b11 = b10
                b12 = HHML
                b13 = HHML
                b14 = AHML - 46
                b15 = AHML - 46

                b1text.Text = b1.ToString()
                b1text.Update()
                b2text.Text = b2.ToString()
                b2text.Update()
                b3text.Text = b3.ToString()
                b3text.Update()
                b4text.Text = b4.ToString()
                b4text.Update()
                b5text.Text = b5.ToString()
                b5text.Update()
                b6text.Text = b6.ToString()
                b6text.Update()
                b7text.Text = b7.ToString()
                b7text.Update()
                b8text.Text = b8.ToString()
                b8text.Update()
                b9text.Text = b9.ToString()
                b9text.Update()
                b10text.Text = b10.ToString()
                b10text.Update()
                b11text.Text = b11.ToString()
                b11text.Update()
                b12text.Text = b12.ToString()
                b12text.Update()
                b13text.Text = "La hoja lenta no lleva este perfil"
                b13text.Update()
                b14text.Text = b14.ToString()
                b14text.Update()
                b15text.Text = b15.ToString()
                b15text.Update()

                bu1 = 2
                bu2 = 2
                bu3 = 4
                bu4 = 2
                bu5 = 2
                bu6 = 2
                bu7 = 2
                bu8 = 1
                bu9 = 1
                bu10 = 1
                bu11 = 1
                bu12 = 4
                bu13 = 1
                bu14 = 2
                bu15 = 2

                bu1text.Text = bu1.ToString()
                bu1text.Update()
                bu2text.Text = bu2.ToString()
                bu2text.Update()
                bu3text.Text = bu3.ToString()
                bu3text.Update()
                bu4text.Text = bu4.ToString()
                bu4text.Update()
                bu5text.Text = bu5.ToString()
                bu5text.Update()
                bu6text.Text = bu6.ToString()
                bu6text.Update()
                bu7text.Text = bu7.ToString()
                bu7text.Update()
                bu8text.Text = bu8.ToString()
                bu8text.Update()
                bu9text.Text = bu9.ToString()
                bu9text.Update()
                bu10text.Text = bu10.ToString()
                bu10text.Update()
                bu11text.Text = bu11.ToString()
                bu11text.Update()
                bu12text.Text = bu12.ToString()
                bu12text.Update()
                bu13text.Text = "La hoja lenta no lleva este perfil"
                bu13text.Update()
                bu14text.Text = bu14.ToString()
                bu14text.Update()
                bu15text.Text = bu15.ToString()
                bu15text.Update()

                Panel4.Visible = True
                Panel3.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Label198.Visible = True
                Panel12.Visible = True
                TabControl2.SelectedIndex = 1

                '****************** FG - CASO VI **************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Full Glass") And cmbconf.SelectedIndex = cmbconf.FindStringExact("4 Hojas Móviles") Then

                'Calculo de la anchura y altura de las hojas móviles con carpintería FG y 4 hojas móviles

                AHM = (pl / 4) + solFG
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                HHM = hh - 5
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                'Calculo de la anchura y altura de las hojas móviles Lentas con carpintería FG y 4 hojas móviles

                AHML = AHM
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                HHML = hh
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                'Calculo de la anchura y altura de los vidrios de las hojas móviles con carpintería FG y 4 hojas móviles

                AVM = AHM - deshFG
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                HVM = HHM - desvFG - 5
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                'Calculo de la anchura y altura de los vidrios de las hojas móviles lentas con carpintería FG y 4 hojas móviles

                AVML = AHML - (3 * deshFG / 2)
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                HVML = HHML - desvFG
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Sacar mensaje de no aplicación de hojas fijas en este caso 

                AHFtext.Text = "No hay h.fijas"
                HHFtext.Text = "No hay h.fijas"
                AVFtext.Text = "No hay h.fijas"
                HVFtext.Text = "No hay h.fijas"

                ' Calculo las medidas de los componentes de la carpintería FG

                b4 = HHM
                b5 = HHM
                b6 = AHM - 46
                b7 = AHM - 46
                b8 = pl
                b9 = wh
                If (AHF > AHM) Then
                    b10 = (2 * AHF) + pl
                ElseIf (AHF < AHM) Then
                    b10 = (2 * AHM) + pl
                ElseIf (AHF = AHM) Then
                    b10 = 2 * AHM + pl
                End If
                b11 = b10
                b12 = HHML
                b13 = HHML
                b14 = AHML - 46
                b15 = AHML - 46

                b1text.Text = "No hay h.fijas"
                b1text.Update()
                b2text.Text = "No hay h.fijas"
                b2text.Update()
                b3text.Text = "No hay h.fijas"
                b3text.Update()
                b4text.Text = b4.ToString()
                b4text.Update()
                b5text.Text = b5.ToString()
                b5text.Update()
                b6text.Text = b6.ToString()
                b6text.Update()
                b7text.Text = b7.ToString()
                b7text.Update()
                b8text.Text = b8.ToString()
                b8text.Update()
                b9text.Text = b9.ToString()
                b9text.Update()
                b10text.Text = b10.ToString()
                b10text.Update()
                b11text.Text = b11.ToString()
                b11text.Update()
                b12text.Text = b12.ToString()
                b12text.Update()
                b13text.Text = "La hoja lenta no lleva este perfil"
                b13text.Update()
                b14text.Text = b14.ToString()
                b14text.Update()
                b15text.Text = b15.ToString()
                b15text.Update()

                bu4 = 2
                bu5 = 2
                bu6 = 2
                bu7 = 2
                bu8 = 1
                bu9 = 1
                bu10 = 1
                bu11 = 1
                bu12 = 4
                bu13 = 1
                bu14 = 2
                bu15 = 2

                bu1text.Text = "No hay h.fijas"
                bu1text.Update()
                bu2text.Text = "No hay h.fijas"
                bu2text.Update()
                bu3text.Text = "No hay h.fijas"
                bu3text.Update()
                bu4text.Text = bu4.ToString()
                bu4text.Update()
                bu5text.Text = bu5.ToString()
                bu5text.Update()
                bu6text.Text = bu6.ToString()
                bu6text.Update()
                bu7text.Text = bu7.ToString()
                bu7text.Update()
                bu8text.Text = bu8.ToString()
                bu8text.Update()
                bu9text.Text = bu9.ToString()
                bu9text.Update()
                bu10text.Text = bu10.ToString()
                bu10text.Update()
                bu11text.Text = bu11.ToString()
                bu11text.Update()
                bu12text.Text = bu12.ToString()
                bu12text.Update()
                bu13text.Text = "La hoja lenta no lleva este perfil"
                bu13text.Update()
                bu14text.Text = bu14.ToString()
                bu14text.Update()
                bu15text.Text = bu15.ToString()
                bu15text.Update()

                Panel4.Visible = True
                Panel3.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Label198.Visible = True
                Panel12.Visible = True
                TabControl2.SelectedIndex = 1

                '****************** FG - CASO VII *************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Full Glass") And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 2 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                AHF = wh - pl
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                HHF = hh - 10
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                AHM = (pl / 2) + solFG
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                HHM = hh - 15
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho de la hoja móvil lenta para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                AHML = (pl / 2) + solFG
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                ' Calcula la altura de la hoja móvil lenta para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                HHML = hh - 10
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                AVF = AHF - deshFG
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                HVF = HHF - desvFG
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                AVM = AHM - deshFG
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                HVM = HHM - desvFG - 5
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil lenta para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                AVML = AHML - ((3 * deshFG) / 2)
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil lenta para carpinterías FG con 1 Hoja Fija + 2 Hojas Móviles

                HVML = HHML - desvFG
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Calculo las medidas de los componentes de la carpintería FG

                b1 = HHF
                b2 = HHF
                b3 = AHF - 46
                b4 = HHM
                b5 = HHM
                b6 = AHM - 46
                b7 = AHM - 46
                b8 = pl
                b9 = wh
                If (AHF > AHM) Then
                    b10 = (AHF) + pl
                ElseIf (AHF < AHM) Then
                    b10 = (AHM) + pl
                ElseIf (AHF = AHM) Then
                    b10 = (AHF) + pl
                End If
                b11 = b10
                b12 = HHML
                b13 = HHML
                b14 = AHML - 46
                b15 = AHML - 46

                b1text.Text = b1.ToString()
                b1text.Update()
                b2text.Text = b2.ToString()
                b2text.Update()
                b3text.Text = b3.ToString()
                b3text.Update()
                b4text.Text = b4.ToString()
                b4text.Update()
                b5text.Text = b5.ToString()
                b5text.Update()
                b6text.Text = b6.ToString()
                b6text.Update()
                b7text.Text = b7.ToString()
                b7text.Update()
                b8text.Text = b8.ToString()
                b8text.Update()
                b9text.Text = b9.ToString()
                b9text.Update()
                b10text.Text = b10.ToString()
                b10text.Update()
                b11text.Text = b11.ToString()
                b11text.Update()
                b12text.Text = b12.ToString()
                b12text.Update()
                b13text.Text = "La hoja lenta no lleva este perfil"
                b13text.Update()
                b14text.Text = b14.ToString()
                b14text.Update()
                b15text.Text = b15.ToString()
                b15text.Update()

                bu1 = 1
                bu2 = 1
                bu3 = 2
                bu4 = 1
                bu5 = 1
                bu6 = 1
                bu7 = 1
                bu8 = 1
                bu9 = 1
                bu10 = 1
                bu11 = 1
                bu12 = 2
                bu13 = 1
                bu14 = 1
                bu15 = 1

                bu1text.Text = bu1.ToString()
                bu1text.Update()
                bu2text.Text = bu2.ToString()
                bu2text.Update()
                bu3text.Text = bu3.ToString()
                bu3text.Update()
                bu4text.Text = bu4.ToString()
                bu4text.Update()
                bu5text.Text = bu5.ToString()
                bu5text.Update()
                bu6text.Text = bu6.ToString()
                bu6text.Update()
                bu7text.Text = bu7.ToString()
                bu7text.Update()
                bu8text.Text = bu8.ToString()
                bu8text.Update()
                bu9text.Text = bu9.ToString()
                bu9text.Update()
                bu10text.Text = bu10.ToString()
                bu10text.Update()
                bu11text.Text = bu11.ToString()
                bu11text.Update()
                bu12text.Text = bu12.ToString()
                bu12text.Update()
                bu13text.Text = "La hoja lenta no lleva este perfil"
                bu13text.Update()
                bu14text.Text = bu14.ToString()
                bu14text.Update()
                bu15text.Text = bu15.ToString()
                bu15text.Update()

                Panel4.Visible = True
                Panel3.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Label198.Visible = True
                Panel12.Visible = True
                TabControl2.SelectedIndex = 1

                '****************** FG - CASO VIII ************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Full Glass") And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Móviles TES") Then

                'Calculo de la anchura y altura de las hojas móviles con carpintería FG y 2 hojas móviles

                AHM = (pl / 2) + solFG
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                HHM = hh - 5
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                'Calculo de la anchura y altura de las hojas móviles Lentas con carpintería FG y 2 hojas móviles

                AHML = AHM
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                HHML = hh
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                'Calculo de la anchura y altura de los vidrios de las hojas móviles con carpintería FG y 2 hojas móviles

                AVM = AHM - deshFG
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                HVM = HHM - desvFG - 5
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                'Calculo de la anchura y altura de los vidrios de las hojas móviles lentas con carpintería FG y 2 hojas móviles

                AVML = AHML - (3 * deshFG / 2)
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                HVML = HHML - desvFG
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Sacar mensaje de no aplicación de hojas fijas en este caso 

                AHFtext.Text = "No hay h.fijas"
                HHFtext.Text = "No hay h.fijas"
                AVFtext.Text = "No hay h.fijas"
                HVFtext.Text = "No hay h.fijas"

                ' Calculo las medidas de los componentes de la carpintería FG

                b4 = HHM
                b5 = HHM
                b6 = AHM - 46
                b7 = AHM - 46
                b8 = pl
                b9 = wh
                If (AHF > AHM) Then
                    b10 = (AHF) + pl
                ElseIf (AHF < AHM) Then
                    b10 = (AHM) + pl
                ElseIf (AHF = AHM) Then
                    b10 = AHM + pl
                End If
                b11 = b10
                b12 = HHML
                b13 = HHML
                b14 = AHML - 46
                b15 = AHML - 46


                b1text.Text = "No hay h.fijas"
                b1text.Update()
                b2text.Text = "No hay h.fijas"
                b2text.Update()
                b3text.Text = "No hay h.fijas"
                b3text.Update()
                b4text.Text = b4.ToString()
                b4text.Update()
                b5text.Text = b5.ToString()
                b5text.Update()
                b6text.Text = b6.ToString()
                b6text.Update()
                b7text.Text = b7.ToString()
                b7text.Update()
                b8text.Text = b8.ToString()
                b8text.Update()
                b9text.Text = b9.ToString()
                b9text.Update()
                b10text.Text = b10.ToString()
                b10text.Update()
                b11text.Text = b11.ToString()
                b11text.Update()
                b12text.Text = b12.ToString()
                b12text.Update()
                b13text.Text = "La hoja lenta no lleva este perfil"
                b13text.Update()
                b14text.Text = b14.ToString()
                b14text.Update()
                b15text.Text = b15.ToString()
                b15text.Update()

                bu4 = 1
                bu5 = 1
                bu6 = 1
                bu7 = 1
                bu8 = 1
                bu9 = 1
                bu10 = 1
                bu11 = 1
                bu12 = 2
                bu13 = 1
                bu14 = 1
                bu15 = 1

                bu1text.Text = "No hay h.fijas"
                bu1text.Update()
                bu2text.Text = "No hay h.fijas"
                bu2text.Update()
                bu3text.Text = "No hay h.fijas"
                bu3text.Update()
                bu4text.Text = bu4.ToString()
                bu4text.Update()
                bu5text.Text = bu5.ToString()
                bu5text.Update()
                bu6text.Text = bu6.ToString()
                bu6text.Update()
                bu7text.Text = bu7.ToString()
                bu7text.Update()
                bu8text.Text = bu8.ToString()
                bu8text.Update()
                bu9text.Text = bu9.ToString()
                bu9text.Update()
                bu10text.Text = bu10.ToString()
                bu10text.Update()
                bu11text.Text = bu11.ToString()
                bu11text.Update()
                bu12text.Text = bu12.ToString()
                bu12text.Update()
                bu13text.Text = "La hoja lenta no lleva este perfil"
                bu13text.Update()
                bu14text.Text = bu14.ToString()
                bu14text.Update()
                bu15text.Text = bu15.ToString()
                bu15text.Update()

                Panel4.Visible = True
                Panel3.Visible = False
                Panel5.Visible = False
                Panel6.Visible = False
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Label198.Visible = True
                Panel12.Visible = True
                TabControl2.SelectedIndex = 1

                '****************** PL SUPERIOR E INFERIOR - CASO I ************************************************************************************************************

            ElseIf (cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plintos silicona superior e inferior") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto superior de pinza e inferior de silicona")) And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 2 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías PL (Superior e inferior) con 2 hojas fijas y 2 móviles

                AHF = ((wh - pl) / 2)
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías PL (Superior e inferior) con 2 hojas fijas y 2 móviles

                HHF = hh
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías PL (Superior e inferior) con 2 hojas fijas y 2 móviles

                AHM = (pl / 2) + solPL
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías PL (Superior e inferior) con 2 hojas fijas y 2 móviles

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías PL (Superior e inferior) con 2 hojas fijas y 2 móviles

                AVF = AHF
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías PL (Superior e inferior) con 2 hojas fijas y 2 móviles

                HVF = HHF - desv2PLF
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías PL (Superior e inferior) con 2 hojas fijas y 2 móviles

                AVM = AHM
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías PL (Superior e inferior) con 2 hojas fijas y 2 móviles

                HVM = HHM - desv2PLM
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calculo las medidas de los componentes de la carpintería PL (Superior e inferior)

                c1 = AHF
                c2 = AHM
                c3 = AHM
                c4 = pl
                If (AHF > AHM) Then
                    c5 = (2 * AHF) + pl
                ElseIf (AHF < AHM) Then
                    c5 = (2 * AHM) + pl
                ElseIf (AHF = AHM) Then
                    c5 = (2 * AHF) + pl
                End If
                c6 = c5

                c1text.Text = c1.ToString()
                c1text.Update()
                c2text.Text = c2.ToString()
                c2text.Update()
                c3text.Text = c3.ToString()
                c3text.Update()
                c4text.Text = c4.ToString()
                c4text.Update()
                c5text.Text = c5.ToString()
                c5text.Update()
                c6text.Text = c6.ToString()
                c6text.Update()

                cu1 = 4
                cu2 = 2
                cu3 = 2
                cu4 = 1
                cu5 = 1
                cu6 = 1

                cu1text.Text = cu1.ToString()
                cu1text.Update()
                cu2text.Text = cu2.ToString()
                cu2text.Update()
                cu3text.Text = cu3.ToString()
                cu3text.Update()
                cu4text.Text = cu4.ToString()
                cu4text.Update()
                cu5text.Text = cu5.ToString()
                cu5text.Update()
                cu6text.Text = cu6.ToString()
                cu6text.Update()

                Panel5.Visible = True
                Panel3.Visible = False
                Panel4.Visible = False
                Panel6.Visible = False
                Label73.Visible = True
                c3text.Visible = True
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                TabControl2.SelectedIndex = 2

                '****************** PL SUPERIOR E INFERIOR - CASO II ************************************************************************************************************

            ElseIf (cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plintos silicona superior e inferior") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto superior de pinza e inferior de silicona")) And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Móviles") Then

                'Calculo de la anchura y altura de las hojas móviles con carpintería PL (Superior e Inferior) y 2 hojas móviles

                AHM = (pl / 2) + solPL
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                'Calculo de la anchura y altura de los vidrios de las hojas móviles con carpintería PL (Superior e Inferior) y 2 hojas móviles

                AVM = AHM
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                HVM = HHM - desv2PLM
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Sacar mensaje de no aplicación de hojas fijas en este caso 

                AHFtext.Text = "No hay h.fijas"
                HHFtext.Text = "No hay h.fijas"
                AVFtext.Text = "No hay h.fijas"
                HVFtext.Text = "No hay h.fijas"

                ' Calculo las medidas de los componentes de la carpintería PL (Superior e inferior)

                c2 = AHM
                c3 = AHM
                c4 = pl
                If (AHF > AHM) Then
                    c5 = (2 * AHF) + pl
                ElseIf (AHF < AHM) Then
                    c5 = (2 * AHM) + pl
                ElseIf (AHF = AHM) Then
                    c5 = (2 * AHM) + pl
                End If
                c6 = c5

                c1text.Text = "No hay h.fijas"
                c1text.Update()
                c2text.Text = c2.ToString()
                c2text.Update()
                c3text.Text = c3.ToString()
                c3text.Update()
                c4text.Text = c4.ToString()
                c4text.Update()
                c5text.Text = c5.ToString()
                c5text.Update()
                c6text.Text = c6.ToString()
                c6text.Update()

                cu2 = 2
                cu3 = 2
                cu4 = 1
                cu5 = 1
                cu6 = 1

                cu1text.Text = "No hay h.fijas"
                cu1text.Update()
                cu2text.Text = cu2.ToString()
                cu2text.Update()
                cu3text.Text = cu3.ToString()
                cu3text.Update()
                cu4text.Text = cu4.ToString()
                cu4text.Update()
                cu5text.Text = cu5.ToString()
                cu5text.Update()
                cu6text.Text = cu6.ToString()
                cu6text.Update()

                Panel5.Visible = True
                Panel3.Visible = False
                Panel4.Visible = False
                Panel6.Visible = False
                Label73.Visible = True
                c3text.Visible = True
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                TabControl2.SelectedIndex = 2

                '****************** PL SUPERIOR E INFERIOR - CASO III ***********************************************************************************************************

            ElseIf (cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plintos silicona superior e inferior") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto superior de pinza e inferior de silicona")) And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 1 Hoja Móvil") Then

                ' Calcula el ancho de la hoja fija para carpinterías PL (Superior e Inferior) con 1 hoja fija y 1 móvil

                AHF = wh - pl
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías PL (Superior e Inferior) con 1 hoja fija y 1 móvil

                HHF = hh
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías PL (Superior e Inferior) con 1 hoja fija y 1 móvil

                AHM = pl + solPL
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías PL (Superior e Inferior) con 1 hoja fija y 1 móvil

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías PL (Superior e Inferior) con 1 hoja fija y 1 móvil

                AVF = AHF
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías PL (Superior e Inferior) con 1 hoja fija y 1 móvil

                HVF = HHF - desv2PLF
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías PL (Superior e Inferior) con 1 hoja fija y 1 móvil

                AVM = AHM
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías PL (Superior e Inferior) con 1 hoja fija y 1 móvil

                HVM = HHM - desv2PLM
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calculo las medidas de los componentes de la carpintería PL (Superior e inferior)

                c1 = AHF
                c2 = AHM
                c3 = AHM
                c4 = pl
                If (AHF > AHM) Then
                    c5 = (AHF) + pl
                ElseIf (AHF < AHM) Then
                    c5 = (AHM) + pl
                ElseIf (AHF = AHM) Then
                    c5 = AHF + pl
                End If
                c6 = c5

                c1text.Text = c1.ToString()
                c1text.Update()
                c2text.Text = c2.ToString()
                c2text.Update()
                c3text.Text = c3.ToString()
                c3text.Update()
                c4text.Text = c4.ToString()
                c4text.Update()
                c5text.Text = c5.ToString()
                c5text.Update()
                c6text.Text = c6.ToString()
                c6text.Update()

                cu1 = 2
                cu2 = 1
                cu3 = 1
                cu4 = 1
                cu5 = 1
                cu6 = 1

                cu1text.Text = cu1.ToString()
                cu1text.Update()
                cu2text.Text = cu2.ToString()
                cu2text.Update()
                cu3text.Text = cu3.ToString()
                cu3text.Update()
                cu4text.Text = cu4.ToString()
                cu4text.Update()
                cu5text.Text = cu5.ToString()
                cu5text.Update()
                cu6text.Text = cu6.ToString()
                cu6text.Update()

                Panel5.Visible = True
                Panel3.Visible = False
                Panel4.Visible = False
                Panel6.Visible = False
                Label73.Visible = True
                c3text.Visible = True
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                TabControl2.SelectedIndex = 2

                '****************** PL SUPERIOR E INFERIOR - CASO IV ****************************************************************************************************************

            ElseIf (cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plintos silicona superior e inferior") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto superior de pinza e inferior de silicona")) And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Móvil") Then

                'Calculo de la anchura y altura de las hojas móviles con carpintería PL (Superior e Inferior) y 1 hoja móvil

                AHM = pl + solPL
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                'Calculo de la anchura y altura de los vidrios de las hojas móviles con carpintería PL (Superior e Inferior) y 1 hoja móvil

                AVM = AHM
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                HVM = HHM - desv2PLM
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Sacar mensaje de no aplicación de hojas fijas en este caso 

                AHFtext.Text = "No hay h.fijas"
                HHFtext.Text = "No hay h.fijas"
                AVFtext.Text = "No hay h.fijas"
                HVFtext.Text = "No hay h.fijas"

                ' Calculo las medidas de los componentes de la carpintería PL (Superior e inferior)

                c2 = AHM
                c3 = AHM
                c4 = pl
                If (AHF > AHM) Then
                    c5 = (AHF) + pl
                ElseIf (AHF < AHM) Then
                    c5 = (AHM) + pl
                ElseIf (AHF = AHM) Then
                    c5 = AHM + pl
                End If
                c6 = c5

                c1text.Text = "No hay h.fijas"
                c1text.Update()
                c2text.Text = c2.ToString()
                c2text.Update()
                c3text.Text = c3.ToString()
                c3text.Update()
                c4text.Text = c4.ToString()
                c4text.Update()
                c5text.Text = c5.ToString()
                c5text.Update()
                c6text.Text = c6.ToString()
                c6text.Update()

                cu2 = 1
                cu3 = 1
                cu4 = 1
                cu5 = 1
                cu6 = 1

                cu1text.Text = "No hay h.fijas"
                cu1text.Update()
                cu2text.Text = cu2.ToString()
                cu2text.Update()
                cu3text.Text = cu3.ToString()
                cu3text.Update()
                cu4text.Text = cu4.ToString()
                cu4text.Update()
                cu5text.Text = cu5.ToString()
                cu5text.Update()
                cu6text.Text = cu6.ToString()
                cu6text.Update()

                Panel5.Visible = True
                Panel3.Visible = False
                Panel4.Visible = False
                Panel6.Visible = False
                Label73.Visible = True
                c3text.Visible = True
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                TabControl2.SelectedIndex = 2

                '****************** PLINTOS SUPERIOR E INFERIOR - CASO V *************************************************************************************************************

            ElseIf (cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plintos silicona superior e inferior") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto superior de pinza e inferior de silicona")) And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 4 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                AHF = (wh - pl) / 2
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                HHF = hh
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                AHM = (pl / 4) + solPL
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                HHM = hh - 10
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                AHML = (pl / 4) + solPL
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                ' Calcula la altura de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                HHML = hh
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                AVF = (wh - pl) / 2
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                HVF = HHF - 15
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                AVM = (pl / 4) + solPL
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                HVM = HHM - desv2PLM
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                AVML = (pl / 4) + solPL
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 2 Hojas Fijas + 4 Hojas Móviles

                HVML = HHML - desv2PLM
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Calculo las medidas de los componentes de la carpintería PL (Superior e inferior)

                c1 = AHF
                c2 = AHM
                c3 = AHM
                c4 = pl
                If (AHF > AHM) Then
                    c5 = (2 * AHF) + pl
                ElseIf (AHF < AHM) Then
                    c5 = (2 * AHM) + pl
                ElseIf (AHF = AHM) Then
                    c5 = (2 * AHF) + pl
                End If
                c6 = c5
                c7 = AHML
                c8 = AHML

                c1text.Text = c1.ToString()
                c1text.Update()
                c2text.Text = c2.ToString()
                c2text.Update()
                c3text.Text = c3.ToString()
                c3text.Update()
                c4text.Text = c4.ToString()
                c4text.Update()
                c5text.Text = c5.ToString()
                c5text.Update()
                c6text.Text = c6.ToString()
                c6text.Update()
                c7text.Text = c7.ToString()
                c7text.Update()
                c8text.Text = c8.ToString()
                c8text.Update()

                cu1 = 2
                cu2 = 2
                cu3 = 2
                cu4 = 1
                cu5 = 1
                cu6 = 1
                cu7 = 2
                cu8 = 2

                cu1text.Text = cu1.ToString()
                cu1text.Update()
                cu2text.Text = cu2.ToString()
                cu2text.Update()
                cu3text.Text = cu3.ToString()
                cu3text.Update()
                cu4text.Text = cu4.ToString()
                cu4text.Update()
                cu5text.Text = cu5.ToString()
                cu5text.Update()
                cu6text.Text = cu6.ToString()
                cu6text.Update()
                cu7text.Text = cu7.ToString()
                cu7text.Update()
                cu8text.Text = cu8.ToString()
                cu8text.Update()

                Panel5.Visible = True
                Panel3.Visible = False
                Panel4.Visible = False
                Panel6.Visible = False
                Label73.Visible = True
                c3text.Visible = True
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Panel13.Visible = True
                Label204.Visible = True
                TabControl2.SelectedIndex = 2

                '****************** PLINTOS SUPERIOR E INFERIOR - CASO VI ************************************************************************************************************

            ElseIf (cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plintos silicona superior e inferior") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto superior de pinza e inferior de silicona")) And cmbconf.SelectedIndex = cmbconf.FindStringExact("4 Hojas Móviles") Then

                ' Calcula el ancho de la hoja móvil para carpinterías PL (Superior e Inferior) con 4 Hojas Móviles

                AHM = (pl / 4) + solPL
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías PL (Superior e Inferior) con 4 Hojas Móviles

                HHM = hh - 10
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 4 Hojas Móviles

                AHML = (pl / 4) + solPL
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                ' Calcula la altura de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 4 Hojas Móviles

                HHML = hh
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías PL (Superior e Inferior) con 4 Hojas Móviles

                AVM = (pl / 4) + solPL
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías PL (Superior e Inferior) con 4 Hojas Móviles

                HVM = HHM - desv2PLM
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 4 Hojas Móviles

                AVML = (pl / 4) + solPL
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 4 Hojas Móviles

                HVML = HHML - desv2PLM
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Sacar mensaje de no aplicación de hojas fijas en este caso 

                AHFtext.Text = "No hay h.fijas"
                HHFtext.Text = "No hay h.fijas"
                AVFtext.Text = "No hay h.fijas"
                HVFtext.Text = "No hay h.fijas"

                ' Calculo las medidas de los componentes de la carpintería PL (Superior e inferior)

                c2 = AHM
                c3 = AHM
                c4 = pl
                If (AHF > AHM) Then
                    c5 = (2 * AHF) + pl
                ElseIf (AHF < AHM) Then
                    c5 = (2 * AHM) + pl
                ElseIf (AHF = AHM) Then
                    c5 = (2 * AHM) + pl
                End If
                c6 = c5
                c7 = AHML
                c8 = AHML

                c1text.Text = "No hay h.fijas"
                c1text.Update()
                c2text.Text = c2.ToString()
                c2text.Update()
                c3text.Text = c3.ToString()
                c3text.Update()
                c4text.Text = c4.ToString()
                c4text.Update()
                c5text.Text = c5.ToString()
                c5text.Update()
                c6text.Text = c6.ToString()
                c6text.Update()
                c7text.Text = c7.ToString()
                c7text.Update()
                c8text.Text = c8.ToString()
                c8text.Update()

                cu2 = 2
                cu3 = 2
                cu4 = 1
                cu5 = 1
                cu6 = 1
                cu7 = 2
                cu8 = 2

                cu1text.Text = "No hay h.fijas"
                cu1text.Update()
                cu2text.Text = cu2.ToString()
                cu2text.Update()
                cu3text.Text = cu3.ToString()
                cu3text.Update()
                cu4text.Text = cu4.ToString()
                cu4text.Update()
                cu5text.Text = cu5.ToString()
                cu5text.Update()
                cu6text.Text = cu6.ToString()
                cu6text.Update()
                cu7text.Text = cu7.ToString()
                cu7text.Update()
                cu8text.Text = cu8.ToString()
                cu8text.Update()

                Panel5.Visible = True
                Panel3.Visible = False
                Panel4.Visible = False
                Panel6.Visible = False
                Label73.Visible = True
                c3text.Visible = True
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Panel13.Visible = True
                Label204.Visible = True
                TabControl2.SelectedIndex = 2

                '****************** PLINTOS SUPERIOR E INFERIOR - CASO VII ***********************************************************************************************************

            ElseIf (cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plintos silicona superior e inferior") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto superior de pinza e inferior de silicona")) And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 2 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                AHF = wh - pl
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                HHF = hh
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                AHM = (pl / 2) + solPL
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                HHM = hh - 10
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                AHML = (pl / 2) + solPL
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                ' Calcula la altura de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                HHML = hh
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                AVF = wh - pl
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                HVF = HHF - 15
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                AVM = (pl / 2) + solPL
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                HVM = HHM - desv2PLM
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                AVML = (pl / 2) + solPL
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 1 Hoja Fija + 2 Hojas Móviles

                HVML = HHML - desv2PLM
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Calculo las medidas de los componentes de la carpintería PL (Superior e inferior)

                c1 = AHF
                c2 = AHM
                c3 = AHM
                c4 = pl
                If (AHF > AHM) Then
                    c5 = AHF + pl
                ElseIf (AHF < AHM) Then
                    c5 = AHM + pl
                ElseIf (AHF = AHM) Then
                    c5 = AHF + pl
                End If
                c6 = c5
                c7 = AHML
                c8 = AHML

                c1text.Text = c1.ToString()
                c1text.Update()
                c2text.Text = c2.ToString()
                c2text.Update()
                c3text.Text = c3.ToString()
                c3text.Update()
                c4text.Text = c4.ToString()
                c4text.Update()
                c5text.Text = c5.ToString()
                c5text.Update()
                c6text.Text = c6.ToString()
                c6text.Update()
                c7text.Text = c7.ToString()
                c7text.Update()
                c8text.Text = c8.ToString()
                c8text.Update()

                cu1 = 1
                cu2 = 1
                cu3 = 1
                cu4 = 1
                cu5 = 1
                cu6 = 1
                cu7 = 1
                cu8 = 1

                cu1text.Text = cu1.ToString()
                cu1text.Update()
                cu2text.Text = cu2.ToString()
                cu2text.Update()
                cu3text.Text = cu3.ToString()
                cu3text.Update()
                cu4text.Text = cu4.ToString()
                cu4text.Update()
                cu5text.Text = cu5.ToString()
                cu5text.Update()
                cu6text.Text = cu6.ToString()
                cu6text.Update()
                cu7text.Text = cu7.ToString()
                cu7text.Update()
                cu8text.Text = cu8.ToString()
                cu8text.Update()

                Panel5.Visible = True
                Panel3.Visible = False
                Panel4.Visible = False
                Panel6.Visible = False
                Label73.Visible = True
                c3text.Visible = True
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Panel13.Visible = True
                Label204.Visible = True
                TabControl2.SelectedIndex = 2

                '****************** PLINTOS SUPERIOR E INFERIOR - CASO VIII **********************************************************************************************************

            ElseIf (cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plintos silicona superior e inferior") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto superior de pinza e inferior de silicona")) And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Móviles TES") Then

                ' Calcula el ancho de la hoja móvil para carpinterías PL (Superior e Inferior) con 2 Hojas Móviles

                AHM = (pl / 2) + solPL
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías PL (Superior e Inferior) con 2 Hojas Móviles

                HHM = hh - 10
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 2 Hojas Móviles

                AHML = (pl / 2) + solPL
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                ' Calcula la altura de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 2 Hojas Móviles

                HHML = hh
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías PL (Superior e Inferior) con 2 Hojas Móviles

                AVM = (pl / 2) + solPL
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías PL (Superior e Inferior) con 2 Hojas Móviles

                HVM = HHM - desv2PLM
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 2 Hojas Móviles

                AVML = (pl / 2) + solPL
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil lenta para carpinterías PL (Superior e Inferior) con 2 Hojas Móviles

                HVML = HHML - desv2PLM
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Sacar mensaje de no aplicación de hojas fijas en este caso 

                AHFtext.Text = "No hay h.fijas"
                HHFtext.Text = "No hay h.fijas"
                AVFtext.Text = "No hay h.fijas"
                HVFtext.Text = "No hay h.fijas"

                ' Calculo las medidas de los componentes de la carpintería PL (Superior e inferior)

                c2 = AHM
                c3 = AHM
                c4 = pl
                If (AHF > AHM) Then
                    c5 = (AHF) + pl
                ElseIf (AHF < AHM) Then
                    c5 = (AHM) + pl
                ElseIf (AHF = AHM) Then
                    c5 = (AHM) + pl
                End If
                c6 = c5
                c7 = AHML
                c8 = AHML

                c1text.Text = "No hay h.fijas"
                c1text.Update()
                c2text.Text = c2.ToString()
                c2text.Update()
                c3text.Text = c3.ToString()
                c3text.Update()
                c4text.Text = c4.ToString()
                c4text.Update()
                c5text.Text = c5.ToString()
                c5text.Update()
                c6text.Text = c6.ToString()
                c6text.Update()
                c7text.Text = c7.ToString()
                c7text.Update()
                c8text.Text = c8.ToString()
                c8text.Update()

                cu2 = 1
                cu3 = 1
                cu4 = 1
                cu5 = 1
                cu6 = 1
                cu7 = 1
                cu8 = 1

                cu1text.Text = "No hay h.fijas"
                cu1text.Update()
                cu2text.Text = cu2.ToString()
                cu2text.Update()
                cu3text.Text = cu3.ToString()
                cu3text.Update()
                cu4text.Text = cu4.ToString()
                cu4text.Update()
                cu5text.Text = cu5.ToString()
                cu5text.Update()
                cu6text.Text = cu6.ToString()
                cu6text.Update()
                cu7text.Text = cu7.ToString()
                cu7text.Update()
                cu8text.Text = cu8.ToString()
                cu8text.Update()

                Panel5.Visible = True
                Panel3.Visible = False
                Panel4.Visible = False
                Panel6.Visible = False
                Label73.Visible = True
                c3text.Visible = True
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Panel13.Visible = True
                Label204.Visible = True
                TabControl2.SelectedIndex = 2

                '****************** PL SUPERIOR - CASO I ************************************************************************************************************

            ElseIf (cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto silicona. Sólo superior") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto de pinza. Sólo Superior")) And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 2 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías PL (Superior) con 2 hojas fijas y 2 móviles

                AHF = ((wh - pl) / 2)
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías PL (Superior) con 2 hojas fijas y 2 móviles

                HHF = hh
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías PL (Superior) con 2 hojas fijas y 2 móviles

                AHM = (pl / 2) + solPL
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías PL (Superior) con 2 hojas fijas y 2 móviles

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías PL (Superior) con 2 hojas fijas y 2 móviles

                AVF = AHF
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías PL (Superior) con 2 hojas fijas y 2 móviles

                HVF = HHF - desv1PLF
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías PL (Superior) con 2 hojas fijas y 2 móviles

                AVM = AHM
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías PL (Superior) con 2 hojas fijas y 2 móviles

                HVM = HHM - desv1PLM
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calculo las medidas de los componentes de la carpintería PL (Superior)

                c1 = AHF
                c2 = AHM
                c4 = pl
                If (AHF > AHM) Then
                    c5 = (2 * AHF) + pl
                ElseIf (AHF < AHM) Then
                    c5 = (2 * AHM) + pl
                ElseIf (AHF = AHM) Then
                    c5 = (2 * AHF) + pl
                End If
                c6 = c5

                c1text.Text = c1.ToString()
                c1text.Update()
                c2text.Text = c2.ToString()
                c2text.Update()
                c3text.Text = c3.ToString()
                c3text.Update()
                c4text.Text = c4.ToString()
                c4text.Update()
                c5text.Text = c5.ToString()
                c5text.Update()
                c6text.Text = c6.ToString()
                c6text.Update()

                cu1 = 4
                cu2 = 2
                cu3 = 2
                cu4 = 1
                cu5 = 1
                cu6 = 1

                cu1text.Text = cu1.ToString()
                cu1text.Update()
                cu2text.Text = cu2.ToString()
                cu2text.Update()
                cu3text.Text = cu3.ToString()
                cu3text.Update()
                cu4text.Text = cu4.ToString()
                cu4text.Update()
                cu5text.Text = cu5.ToString()
                cu5text.Update()
                cu6text.Text = cu6.ToString()
                cu6text.Update()

                Panel5.Visible = True
                Panel3.Visible = False
                Panel4.Visible = False
                Panel6.Visible = False
                Label73.Visible = False
                c3text.Visible = False
                cu3text.Visible = False
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False

                TabControl2.SelectedIndex = 2


                '****************** PL SUPERIOR - CASO II ************************************************************************************************************

            ElseIf (cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto silicona. Sólo superior") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto de pinza. Sólo Superior")) And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Móviles") Then

                'Calculo de la anchura y altura de las hojas móviles con carpintería PL (Superior) y 2 hojas móviles

                AHM = (pl / 2) + solPL
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                'Calculo de la anchura y altura de los vidrios de las hojas móviles con carpintería PL (Superior) y 2 hojas móviles

                AVM = AHM
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                HVM = HHM - desv1PLM
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Sacar mensaje de no aplicación de hojas fijas en este caso 

                AHFtext.Text = "No hay h.fijas"
                HHFtext.Text = "No hay h.fijas"
                AVFtext.Text = "No hay h.fijas"
                HVFtext.Text = "No hay h.fijas"

                ' Calculo las medidas de los componentes de la carpintería PL (Superior)

                c2 = AHM
                c4 = pl
                If (AHF > AHM) Then
                    c5 = (2 * AHF) + pl
                ElseIf (AHF < AHM) Then
                    c5 = (2 * AHM) + pl
                ElseIf (AHF = AHM) Then
                    c5 = (2 * AHM) + pl
                End If
                c6 = c5

                c1text.Text = "No hay h.fijas"
                c1text.Update()
                c2text.Text = c2.ToString()
                c2text.Update()
                c3text.Text = c3.ToString()
                c3text.Update()
                c4text.Text = c4.ToString()
                c4text.Update()
                c5text.Text = c5.ToString()
                c5text.Update()
                c6text.Text = c6.ToString()
                c6text.Update()

                cu2 = 2
                cu3 = 2
                cu4 = 1
                cu5 = 1
                cu6 = 1

                cu1text.Text = "No hay h.fijas"
                cu1text.Update()
                cu2text.Text = cu2.ToString()
                cu2text.Update()
                cu3text.Text = cu3.ToString()
                cu3text.Update()
                cu4text.Text = cu4.ToString()
                cu4text.Update()
                cu5text.Text = cu5.ToString()
                cu5text.Update()
                cu6text.Text = cu6.ToString()
                cu6text.Update()

                Panel5.Visible = True
                Panel3.Visible = False
                Panel4.Visible = False
                Panel6.Visible = False
                Label73.Visible = False
                c3text.Visible = False
                cu3text.Visible = False
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False

                TabControl2.SelectedIndex = 2

                '****************** PL SUPERIOR - CASO III ***********************************************************************************************************

            ElseIf (cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto silicona. Sólo superior") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto de pinza. Sólo Superior")) And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 1 Hoja Móvil") Then

                ' Calcula el ancho de la hoja fija para carpinterías PL (Superior) con 1 hoja fija y 1 móvil

                AHF = wh - pl
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías PL (Superior) con 1 hoja fija y 1 móvil

                HHF = hh
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías PL (Superior) con 1 hoja fija y 1 móvil

                AHM = pl + solPL
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías PL (Superior) con 1 hoja fija y 1 móvil

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías PL (Superior) con 1 hoja fija y 1 móvil

                AVF = AHF
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías PL (Superior) con 1 hoja fija y 1 móvil

                HVF = HHF - desv1PLF
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías PL (Superior) con 1 hoja fija y 1 móvil

                AVM = AHM
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías PL (Superior) con 1 hoja fija y 1 móvil

                HVM = HHM - desv1PLM
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calculo las medidas de los componentes de la carpintería PL (Superior)

                c1 = AHF
                c2 = AHM
                c4 = pl
                If (AHF > AHM) Then
                    c5 = (AHF) + pl
                ElseIf (AHF < AHM) Then
                    c5 = (AHM) + pl
                ElseIf (AHF = AHM) Then
                    c5 = AHF + pl
                End If
                c6 = c5

                c1text.Text = c1.ToString()
                c1text.Update()
                c2text.Text = c2.ToString()
                c2text.Update()
                c3text.Text = c3.ToString()
                c3text.Update()
                c4text.Text = c4.ToString()
                c4text.Update()
                c5text.Text = c5.ToString()
                c5text.Update()
                c6text.Text = c6.ToString()
                c6text.Update()

                cu1 = 2
                cu2 = 1
                cu3 = 1
                cu4 = 1
                cu5 = 1
                cu6 = 1

                cu1text.Text = cu1.ToString()
                cu1text.Update()
                cu2text.Text = cu2.ToString()
                cu2text.Update()
                cu3text.Text = cu3.ToString()
                cu3text.Update()
                cu4text.Text = cu4.ToString()
                cu4text.Update()
                cu5text.Text = cu5.ToString()
                cu5text.Update()
                cu6text.Text = cu6.ToString()
                cu6text.Update()

                Panel5.Visible = True
                Panel3.Visible = False
                Panel4.Visible = False
                Panel6.Visible = False
                Label73.Visible = False
                c3text.Visible = False
                cu3text.Visible = False
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                TabControl2.SelectedIndex = 2

                '****************** PL SUPERIOR - CASO IV ****************************************************************************************************************

            ElseIf (cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto silicona. Sólo superior") Or cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("Plinto de pinza. Sólo Superior")) And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Móvil") Then

                'Calculo de la anchura y altura de las hojas móviles con carpintería PL (Superior) y 1 hoja móvil

                AHM = pl + solPL
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                'Calculo de la anchura y altura de los vidrios de las hojas móviles con carpintería PL (Superior) y 1 hoja móvil

                AVM = AHM
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                HVM = HHM - desv1PLM
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Sacar mensaje de no aplicación de hojas fijas en este caso 

                AHFtext.Text = "No hay h.fijas"
                HHFtext.Text = "No hay h.fijas"
                AVFtext.Text = "No hay h.fijas"
                HVFtext.Text = "No hay h.fijas"

                ' Calculo las medidas de los componentes de la carpintería PL (Superior)

                c2 = AHM
                c4 = pl
                If (AHF > AHM) Then
                    c5 = (AHF) + pl
                ElseIf (AHF < AHM) Then
                    c5 = (AHM) + pl
                ElseIf (AHF = AHM) Then
                    c5 = AHM + pl
                End If
                c6 = c5

                c1text.Text = "No hay h.fijas"
                c1text.Update()
                c2text.Text = c2.ToString()
                c2text.Update()
                c3text.Text = c3.ToString()
                c3text.Update()
                c4text.Text = c4.ToString()
                c4text.Update()
                c5text.Text = c5.ToString()
                c5text.Update()
                c6text.Text = c6.ToString()
                c6text.Update()

                cu2 = 1
                cu3 = 1
                cu4 = 1
                cu5 = 1
                cu6 = 1

                cu1text.Text = "No hay h.fijas"
                cu1text.Update()
                cu2text.Text = cu2.ToString()
                cu2text.Update()
                cu3text.Text = cu3.ToString()
                cu3text.Update()
                cu4text.Text = cu4.ToString()
                cu4text.Update()
                cu5text.Text = cu5.ToString()
                cu5text.Update()
                cu6text.Text = cu6.ToString()
                cu6text.Update()

                Panel5.Visible = True
                Panel3.Visible = False
                Panel4.Visible = False
                Panel6.Visible = False
                Label73.Visible = False
                c3text.Visible = False
                cu3text.Visible = False
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                TabControl2.SelectedIndex = 2

                '****************** BW52 - CASO I *****************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("BW52") And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 2 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías BW52 con 2 hojas fijas y 2 móviles

                AHF = ((wh - pl) / 2)
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías BW52 con 2 hojas fijas y 2 móviles

                HHF = hh - 10
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías BW52 con 2 hojas fijas y 2 móviles

                AHM = (pl / 2) + solBW52
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías BW52 con 2 hojas fijas y 2 móviles

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías BW52 con 2 hojas fijas y 2 móviles

                AVF = AHF - deshfBW52
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías BW52 con 2 hojas fijas y 2 móviles

                HVF = HHF - desvBW52
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías BW52 con 2 hojas fijas y 2 móviles

                AVM = AHM - deshmBW52
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías BW52 con 2 hojas fijas y 2 móviles

                HVM = HHM - desvBW52
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calculo las medidas de los componentes de la carpintería BW52

                d1 = HHF - 22
                d2 = HHF + 8
                d3 = HHF + 30
                d4 = AHF - 10
                d5 = AHF + 34
                d6 = HHF - 113
                d7 = AHF - 144
                d8 = AHF - 10
                d9 = HHM - 22
                d10 = HHM
                d11 = HHM
                d12 = HHM
                d13 = AHM - 20
                d14 = AHM + 24
                d15 = AHM - 13
                d16 = HHM - 113
                d17 = AHM - 154
                d18 = AHM - 20
                d19 = pl
                d20 = AHF

                If (AHF > AHM) Then
                    d21 = (2 * AHF) + pl
                ElseIf (AHF < AHM) Then
                    d21 = (2 * AHM) + pl
                ElseIf (AHF = AHM) Then
                    d21 = (2 * AHF) + pl
                End If
                d22 = d21

                d1text.Text = d1.ToString()
                d1text.Update()
                d2text.Text = d2.ToString()
                d2text.Update()
                d3text.Text = d3.ToString()
                d3text.Update()
                d4text.Text = d4.ToString()
                d4text.Update()
                d5text.Text = d5.ToString()
                d5text.Update()
                d6text.Text = d6.ToString()
                d6text.Update()
                d7text.Text = d7.ToString()
                d7text.Update()
                d8text.Text = d8.ToString()
                d8text.Update()
                d9text.Text = d9.ToString()
                d9text.Update()
                d10text.Text = d10.ToString()
                d10text.Update()
                d11text.Text = d11.ToString()
                d11text.Update()
                d12text.Text = d12.ToString()
                d12text.Update()
                d13text.Text = d13.ToString()
                d13text.Update()
                d14text.Text = d14.ToString()
                d14text.Update()
                d15text.Text = d15.ToString()
                d15text.Update()
                d16text.Text = d16.ToString()
                d16text.Update()
                d17text.Text = d17.ToString()
                d17text.Update()
                d18text.Text = d18.ToString()
                d18text.Update()
                d19text.Text = d19.ToString()
                d19text.Update()
                d20text.Text = d20.ToString()
                d20text.Update()
                d21text.Text = d21.ToString()
                d21text.Update()
                d22text.Text = d22.ToString()
                d22text.Update()

                du1 = 4
                du2 = 2
                du3 = 2
                du4 = 2
                du5 = 2
                du6 = 4
                du7 = 4
                du8 = 2
                du9 = 4
                du10 = 2
                du11 = 2
                du12 = 2
                du13 = 2
                du14 = 2
                du15 = 2
                du16 = 4
                du17 = 4
                du18 = 2
                du19 = 1
                du20 = 2
                du21 = 1
                du22 = 1

                du1text.Text = du1.ToString()
                du1text.Update()
                du2text.Text = du2.ToString()
                du2text.Update()
                du3text.Text = du3.ToString()
                du3text.Update()
                du4text.Text = du4.ToString()
                du4text.Update()
                du5text.Text = du5.ToString()
                du5text.Update()
                du6text.Text = du6.ToString()
                du6text.Update()
                du7text.Text = du7.ToString()
                du7text.Update()
                du8text.Text = du8.ToString()
                du8text.Update()
                du9text.Text = du9.ToString()
                du9text.Update()
                du10text.Text = du10.ToString()
                du10text.Update()
                du11text.Text = du11.ToString()
                du11text.Update()
                du12text.Text = du12.ToString()
                du12text.Update()
                du13text.Text = du13.ToString()
                du13text.Update()
                du14text.Text = du14.ToString()
                du14text.Update()
                du15text.Text = du15.ToString()
                du15text.Update()
                du16text.Text = du16.ToString()
                du16text.Update()
                du17text.Text = du17.ToString()
                du17text.Update()
                du18text.Text = du18.ToString()
                du18text.Update()
                du19text.Text = du19.ToString()
                du19text.Update()
                du20text.Text = du20.ToString()
                du20text.Update()
                du21text.Text = du21.ToString()
                du21text.Update()
                du22text.Text = du22.ToString()
                du22text.Update()

                Panel3.Visible = False
                Panel4.Visible = False
                Panel5.Visible = False
                Panel6.Visible = True
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                TabControl2.SelectedIndex = 3

                '****************** BW52 - CASO II *****************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("BW52") And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Móviles") Then

                'Calculo de la anchura y altura de las hojas móviles con carpintería BW52 y 2 hojas móviles

                AHM = (pl / 2) + solBW52
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                'Calculo de la anchura y altura de los vidrios de las hojas móviles con carpintería BW52 y 2 hojas móviles

                AVM = AHM - deshmBW52
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                HVM = HHM - desvBW52
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Sacar mensaje de no aplicación de hojas fijas en este caso 

                AHFtext.Text = "No hay h.fijas"
                HHFtext.Text = "No hay h.fijas"
                AVFtext.Text = "No hay h.fijas"
                HVFtext.Text = "No hay h.fijas"

                ' Calculo las medidas de los componentes de la carpintería BW52

                d9 = HHM - 22
                d10 = HHM
                d11 = HHM
                d12 = HHM
                d13 = AHM - 20
                d14 = AHM + 24
                d15 = AHM - 13
                d16 = HHM - 113
                d17 = AHM - 154
                d18 = AHM - 20
                d19 = pl
                d20 = AHF

                If (AHF > AHM) Then
                    d21 = (2 * AHF) + pl
                ElseIf (AHF < AHM) Then
                    d21 = (2 * AHM) + pl
                ElseIf (AHF = AHM) Then
                    d21 = (2 * AHM) + pl
                End If
                d22 = d21

                d1text.Text = "No hay h.fijas"
                d1text.Update()
                d2text.Text = "No hay h.fijas"
                d2text.Update()
                d3text.Text = "No hay h.fijas"
                d3text.Update()
                d4text.Text = "No hay h.fijas"
                d4text.Update()
                d5text.Text = "No hay h.fijas"
                d5text.Update()
                d6text.Text = "No hay h.fijas"
                d6text.Update()
                d7text.Text = "No hay h.fijas"
                d7text.Update()
                d8text.Text = "No hay h.fijas"
                d8text.Update()
                d9text.Text = d9.ToString()
                d9text.Update()
                d10text.Text = d10.ToString()
                d10text.Update()
                d11text.Text = d11.ToString()
                d11text.Update()
                d12text.Text = d12.ToString()
                d12text.Update()
                d13text.Text = d13.ToString()
                d13text.Update()
                d14text.Text = d14.ToString()
                d14text.Update()
                d15text.Text = d15.ToString()
                d15text.Update()
                d16text.Text = d16.ToString()
                d16text.Update()
                d17text.Text = d17.ToString()
                d17text.Update()
                d18text.Text = d18.ToString()
                d18text.Update()
                d19text.Text = d19.ToString()
                d19text.Update()
                d20text.Text = d20.ToString()
                d20text.Update()
                d21text.Text = d21.ToString()
                d21text.Update()
                d22text.Text = d22.ToString()
                d22text.Update()

                du9 = 4
                du10 = 2
                du11 = 2
                du12 = 2
                du13 = 2
                du14 = 2
                du15 = 2
                du16 = 4
                du17 = 4
                du18 = 2
                du19 = 1
                du20 = 2
                du21 = 1
                du22 = 1

                du1text.Text = "No hay h.fijas"
                du1text.Update()
                du2text.Text = "No hay h.fijas"
                du2text.Update()
                du3text.Text = "No hay h.fijas"
                du3text.Update()
                du4text.Text = "No hay h.fijas"
                du4text.Update()
                du5text.Text = "No hay h.fijas"
                du5text.Update()
                du6text.Text = "No hay h.fijas"
                du6text.Update()
                du7text.Text = "No hay h.fijas"
                du7text.Update()
                du8text.Text = "No hay h.fijas"
                du8text.Update()
                du9text.Text = du9.ToString()
                du9text.Update()
                du10text.Text = du10.ToString()
                du10text.Update()
                du11text.Text = du11.ToString()
                du11text.Update()
                du12text.Text = du12.ToString()
                du12text.Update()
                du13text.Text = du13.ToString()
                du13text.Update()
                du14text.Text = du14.ToString()
                du14text.Update()
                du15text.Text = du15.ToString()
                du15text.Update()
                du16text.Text = du16.ToString()
                du16text.Update()
                du17text.Text = du17.ToString()
                du17text.Update()
                du18text.Text = du18.ToString()
                du18text.Update()
                du19text.Text = du19.ToString()
                du19text.Update()
                du20text.Text = du20.ToString()
                du20text.Update()
                du21text.Text = du21.ToString()
                du21text.Update()
                du22text.Text = du22.ToString()
                du22text.Update()

                Panel3.Visible = False
                Panel4.Visible = False
                Panel5.Visible = False
                Panel6.Visible = True
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                TabControl2.SelectedIndex = 3

                '****************** BW52 - CASO III ****************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("BW52") And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 1 Hoja Móvil") Then

                ' Calcula el ancho de la hoja fija para carpinterías BW52 con 1 hoja fija y 1 móvil

                AHF = wh - pl
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías BW52 con 1 hoja fija y 1 móvil

                HHF = hh - 10
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías BW52 con 1 hoja fija y 1 móvil

                AHM = pl + solBW52
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías BW52 con 1 hoja fija y 1 móvil

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías BW52 con 1 hoja fija y 1 móvil

                AVF = AHF - deshfBW52
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías BW52 con 1 hoja fija y 1 móvil

                HVF = HHF - desvBW52
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías BW52 con 1 hoja fija y 1 móvil

                AVM = AHM - deshmBW52
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías BW52 con 1 hoja fija y 1 móvil

                HVM = HHM - desvBW52
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calculo las medidas de los componentes de la carpintería BW52

                d1 = HHF - 22
                d2 = HHF + 8
                d3 = HHF + 30
                d4 = AHF - 10
                d5 = AHF + 34
                d6 = HHF - 113
                d7 = AHF - 144
                d8 = AHF - 10
                d9 = HHM - 22
                d10 = HHM
                d11 = HHM
                d12 = HHM
                d13 = AHM - 20
                d14 = AHM + 24
                d15 = AHM - 13
                d16 = HHM - 113
                d17 = AHM - 154
                d18 = AHM - 20
                d19 = pl
                d20 = AHF

                If (AHF > AHM) Then
                    d21 = (AHF) + pl
                ElseIf (AHF < AHM) Then
                    d21 = (AHM) + pl
                ElseIf (AHF = AHM) Then
                    d21 = AHF + pl
                End If
                d22 = d21

                d1text.Text = d1.ToString()
                d1text.Update()
                d2text.Text = d2.ToString()
                d2text.Update()
                d3text.Text = d3.ToString()
                d3text.Update()
                d4text.Text = d4.ToString()
                d4text.Update()
                d5text.Text = d5.ToString()
                d5text.Update()
                d6text.Text = d6.ToString()
                d6text.Update()
                d7text.Text = d7.ToString()
                d7text.Update()
                d8text.Text = d8.ToString()
                d8text.Update()
                d9text.Text = d9.ToString()
                d9text.Update()
                d10text.Text = d10.ToString()
                d10text.Update()
                d11text.Text = d11.ToString()
                d11text.Update()
                d12text.Text = d12.ToString()
                d12text.Update()
                d13text.Text = d13.ToString()
                d13text.Update()
                d14text.Text = d14.ToString()
                d14text.Update()
                d15text.Text = d15.ToString()
                d15text.Update()
                d16text.Text = d16.ToString()
                d16text.Update()
                d17text.Text = d17.ToString()
                d17text.Update()
                d18text.Text = d18.ToString()
                d18text.Update()
                d19text.Text = d19.ToString()
                d19text.Update()
                d20text.Text = d20.ToString()
                d20text.Update()
                d21text.Text = d21.ToString()
                d21text.Update()
                d22text.Text = d22.ToString()
                d22text.Update()

                du1 = 2
                du2 = 1
                du3 = 1
                du4 = 1
                du5 = 1
                du6 = 2
                du7 = 2
                du8 = 1
                du9 = 2
                du10 = 1
                du11 = 1
                du12 = 1
                du13 = 1
                du14 = 1
                du15 = 1
                du16 = 2
                du17 = 2
                du18 = 1
                du19 = 1
                du20 = 1
                du21 = 1
                du22 = 1

                du1text.Text = du1.ToString()
                du1text.Update()
                du2text.Text = du2.ToString()
                du2text.Update()
                du3text.Text = du3.ToString()
                du3text.Update()
                du4text.Text = du4.ToString()
                du4text.Update()
                du5text.Text = du5.ToString()
                du5text.Update()
                du6text.Text = du6.ToString()
                du6text.Update()
                du7text.Text = du7.ToString()
                du7text.Update()
                du8text.Text = du8.ToString()
                du8text.Update()
                du9text.Text = du9.ToString()
                du9text.Update()
                du10text.Text = du10.ToString()
                du10text.Update()
                du11text.Text = du11.ToString()
                du11text.Update()
                du12text.Text = du12.ToString()
                du12text.Update()
                du13text.Text = du13.ToString()
                du13text.Update()
                du14text.Text = du14.ToString()
                du14text.Update()
                du15text.Text = du15.ToString()
                du15text.Update()
                du16text.Text = du16.ToString()
                du16text.Update()
                du17text.Text = du17.ToString()
                du17text.Update()
                du18text.Text = du18.ToString()
                du18text.Update()
                du19text.Text = du19.ToString()
                du19text.Update()
                du20text.Text = du20.ToString()
                du20text.Update()
                du21text.Text = du21.ToString()
                du21text.Update()
                du22text.Text = du22.ToString()
                du22text.Update()

                Panel3.Visible = False
                Panel4.Visible = False
                Panel5.Visible = False
                Panel6.Visible = True
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                TabControl2.SelectedIndex = 3

                '****************** BW52 - CASO IV ****************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("BW52") And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Móvil") Then

                'Calculo de la anchura y altura de las hojas móviles con carpintería BW52 y 1 hoja móvil

                AHM = pl + solBW52
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                HHM = hh
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                'Calculo de la anchura y altura de los vidrios de las hojas móviles con carpintería BW52 y 1 hoja móvil

                AVM = AHM - deshmBW52
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                HVM = HHM - desvBW52
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Sacar mensaje de no aplicación de hojas fijas en este caso 

                AHFtext.Text = "No hay h.fijas"
                HHFtext.Text = "No hay h.fijas"
                AVFtext.Text = "No hay h.fijas"
                HVFtext.Text = "No hay h.fijas"

                d9 = HHM - 22
                d10 = HHM
                d11 = HHM
                d12 = HHM
                d13 = AHM - 20
                d14 = AHM + 24
                d15 = AHM - 13
                d16 = HHM - 113
                d17 = AHM - 154
                d18 = AHM - 20
                d19 = pl
                d20 = AHF

                If (AHF > AHM) Then
                    d21 = (AHF) + pl
                ElseIf (AHF < AHM) Then
                    d21 = (AHM) + pl
                ElseIf (AHF = AHM) Then
                    d21 = AHM + pl
                End If
                d22 = d21

                d1text.Text = "No hay h.fijas"
                d1text.Update()
                d2text.Text = "No hay h.fijas"
                d2text.Update()
                d3text.Text = "No hay h.fijas"
                d3text.Update()
                d4text.Text = "No hay h.fijas"
                d4text.Update()
                d5text.Text = "No hay h.fijas"
                d5text.Update()
                d6text.Text = "No hay h.fijas"
                d6text.Update()
                d7text.Text = "No hay h.fijas"
                d7text.Update()
                d8text.Text = "No hay h.fijas"
                d8text.Update()
                d9text.Text = d9.ToString()
                d9text.Update()
                d10text.Text = d10.ToString()
                d10text.Update()
                d11text.Text = d11.ToString()
                d11text.Update()
                d12text.Text = d12.ToString()
                d12text.Update()
                d13text.Text = d13.ToString()
                d13text.Update()
                d14text.Text = d14.ToString()
                d14text.Update()
                d15text.Text = d15.ToString()
                d15text.Update()
                d16text.Text = d16.ToString()
                d16text.Update()
                d17text.Text = d17.ToString()
                d17text.Update()
                d18text.Text = d18.ToString()
                d18text.Update()
                d19text.Text = d19.ToString()
                d19text.Update()
                d20text.Text = d20.ToString()
                d20text.Update()
                d21text.Text = d21.ToString()
                d21text.Update()
                d22text.Text = d22.ToString()
                d22text.Update()

                du9 = 2
                du10 = 1
                du11 = 1
                du12 = 1
                du13 = 1
                du14 = 1
                du15 = 1
                du16 = 2
                du17 = 2
                du18 = 1
                du19 = 1
                du20 = 1
                du21 = 1
                du22 = 1

                du1text.Text = "No hay h.fijas"
                du1text.Update()
                du2text.Text = "No hay h.fijas"
                du2text.Update()
                du3text.Text = "No hay h.fijas"
                du3text.Update()
                du4text.Text = "No hay h.fijas"
                du4text.Update()
                du5text.Text = "No hay h.fijas"
                du5text.Update()
                du6text.Text = "No hay h.fijas"
                du6text.Update()
                du7text.Text = "No hay h.fijas"
                du7text.Update()
                du8text.Text = "No hay h.fijas"
                du8text.Update()
                du9text.Text = du9.ToString()
                du9text.Update()
                du10text.Text = du10.ToString()
                du10text.Update()
                du11text.Text = du11.ToString()
                du11text.Update()
                du12text.Text = du12.ToString()
                du12text.Update()
                du13text.Text = du13.ToString()
                du13text.Update()
                du14text.Text = du14.ToString()
                du14text.Update()
                du15text.Text = du15.ToString()
                du15text.Update()
                du16text.Text = du16.ToString()
                du16text.Update()
                du17text.Text = du17.ToString()
                du17text.Update()
                du18text.Text = du18.ToString()
                du18text.Update()
                du19text.Text = du19.ToString()
                du19text.Update()
                du20text.Text = du20.ToString()
                du20text.Update()
                du21text.Text = du21.ToString()
                du21text.Update()
                du22text.Text = du22.ToString()
                du22text.Update()

                Panel3.Visible = False
                Panel4.Visible = False
                Panel5.Visible = False
                Panel6.Visible = True
                Label105.Visible = False
                AHMLtext.Visible = False
                HHMLtext.Visible = False
                Label106.Visible = False
                AVMLtext.Visible = False
                HVMLtext.Visible = False
                Label107.Visible = False
                Label108.Visible = False
                TabControl2.SelectedIndex = 3

                '****************** BW52 - CASO V *****************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("BW52") And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 4 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                AHF = wh - pl
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                HHF = hh - 10
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                AHM = (pl / 4) + solBW52
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                HHM = hh - 5
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho de la hoja móvil lenta para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                AHML = (pl / 4) + solBW52
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                ' Calcula la altura de la hoja móvil lenta para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                HHML = hh
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                AVF = AHF - 119
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                HVF = HHF - desvBW52
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                AVM = AHM - 119
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                HVM = HHM - desvBW52
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil lenta para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                AVML = AHML - 129
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil lenta para carpinterías BW52 con 2 Hojas Fijas + 4 Hojas Móviles

                HVML = HHML - desvBW52
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Calculo las medidas de los componentes de la carpintería BW52

                d1 = HHF - 22
                d2 = HHF + 8
                d3 = HHF + 30
                d4 = AHF - 10
                d5 = AHF + 34
                d6 = HHF - 113
                d7 = AHF - 144
                d8 = AHF - 10
                d9 = HHM - 22
                d10 = HHM
                d11 = HHM
                d12 = HHM
                d13 = AHM - 20
                d14 = AHM + 24
                d15 = AHM - 13
                d16 = HHM - 113
                d17 = AHM - 154
                d18 = AHM - 20
                d19 = pl
                d20 = AHF

                If (AHF > AHM) Then
                    d21 = (2 * AHF) + pl
                ElseIf (AHF < AHM) Then
                    d21 = (2 * AHM) + pl
                ElseIf (AHF = AHM) Then
                    d21 = (2 * AHF) + pl
                End If
                d22 = d21
                d23 = HHM - 22
                d24 = HHM
                d25 = HHM
                d26 = HHM
                d27 = AHM - 20
                d28 = AHM + 24
                d29 = AHM - 13
                d30 = HHM - 113
                d31 = AHM - 154
                d32 = AHM - 20

                d1text.Text = d1.ToString()
                d1text.Update()
                d2text.Text = d2.ToString()
                d2text.Update()
                d3text.Text = d3.ToString()
                d3text.Update()
                d4text.Text = d4.ToString()
                d4text.Update()
                d5text.Text = d5.ToString()
                d5text.Update()
                d6text.Text = d6.ToString()
                d6text.Update()
                d7text.Text = d7.ToString()
                d7text.Update()
                d8text.Text = d8.ToString()
                d8text.Update()
                d9text.Text = d9.ToString()
                d9text.Update()
                d10text.Text = d10.ToString()
                d10text.Update()
                d11text.Text = d11.ToString()
                d11text.Update()
                d12text.Text = d12.ToString()
                d12text.Update()
                d13text.Text = d13.ToString()
                d13text.Update()
                d14text.Text = d14.ToString()
                d14text.Update()
                d15text.Text = d15.ToString()
                d15text.Update()
                d16text.Text = d16.ToString()
                d16text.Update()
                d17text.Text = d17.ToString()
                d17text.Update()
                d18text.Text = d18.ToString()
                d18text.Update()
                d19text.Text = d19.ToString()
                d19text.Update()
                d20text.Text = d20.ToString()
                d20text.Update()
                d21text.Text = d21.ToString()
                d21text.Update()
                d22text.Text = d22.ToString()
                d22text.Update()
                d23text.Text = d23.ToString()
                d23text.Update()
                d24text.Text = d24.ToString()
                d24text.Update()
                d25text.Text = "La hoja lenta no lleva este perfil"
                d25text.Update()
                d26text.Text = d26.ToString()
                d26text.Update()
                d27text.Text = d27.ToString()
                d27text.Update()
                d28text.Text = d28.ToString()
                d28text.Update()
                d29text.Text = d29.ToString()
                d29text.Update()
                d30text.Text = d30.ToString()
                d30text.Update()
                d31text.Text = d31.ToString()
                d31text.Update()
                d32text.Text = d32.ToString()
                d32text.Update()

                du1 = 4
                du2 = 2
                du3 = 2
                du4 = 2
                du5 = 2
                du6 = 4
                du7 = 4
                du8 = 2
                du9 = 4
                du10 = 2
                du11 = 2
                du12 = 2
                du13 = 2
                du14 = 2
                du15 = 2
                du16 = 4
                du17 = 4
                du18 = 2
                du19 = 1
                du20 = 2
                du21 = 1
                du22 = 1
                du23 = 4
                du24 = 4
                du25 = 1
                du26 = 4
                du27 = 2
                du28 = 2
                du29 = 2
                du30 = 4
                du31 = 4
                du32 = 2

                du1text.Text = du1.ToString()
                du1text.Update()
                du2text.Text = du2.ToString()
                du2text.Update()
                du3text.Text = du3.ToString()
                du3text.Update()
                du4text.Text = du4.ToString()
                du4text.Update()
                du5text.Text = du5.ToString()
                du5text.Update()
                du6text.Text = du6.ToString()
                du6text.Update()
                du7text.Text = du7.ToString()
                du7text.Update()
                du8text.Text = du8.ToString()
                du8text.Update()
                du9text.Text = du9.ToString()
                du9text.Update()
                du10text.Text = du10.ToString()
                du10text.Update()
                du11text.Text = du11.ToString()
                du11text.Update()
                du12text.Text = du12.ToString()
                du12text.Update()
                du13text.Text = du13.ToString()
                du13text.Update()
                du14text.Text = du14.ToString()
                du14text.Update()
                du15text.Text = du15.ToString()
                du15text.Update()
                du16text.Text = du16.ToString()
                du16text.Update()
                du17text.Text = du17.ToString()
                du17text.Update()
                du18text.Text = du18.ToString()
                du18text.Update()
                du19text.Text = du19.ToString()
                du19text.Update()
                du20text.Text = du20.ToString()
                du20text.Update()
                du21text.Text = du21.ToString()
                du21text.Update()
                du22text.Text = du22.ToString()
                du22text.Update()
                du23text.Text = du23.ToString()
                du23text.Update()
                du24text.Text = du24.ToString()
                du24text.Update()
                du25text.Text = "La hoja lenta no lleva este perfil"
                du25text.Update()
                du26text.Text = du26.ToString()
                du26text.Update()
                du27text.Text = du27.ToString()
                du27text.Update()
                du28text.Text = du28.ToString()
                du28text.Update()
                du29text.Text = du29.ToString()
                du29text.Update()
                du30text.Text = du30.ToString()
                du30text.Update()
                du31text.Text = du31.ToString()
                du31text.Update()
                du32text.Text = du32.ToString()
                du32text.Update()

                Panel3.Visible = False
                Panel4.Visible = False
                Panel5.Visible = False
                Panel6.Visible = True
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Panel14.Visible = True
                Label209.Visible = True
                TabControl2.SelectedIndex = 3

                '****************** BW52 - CASO VI ****************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("BW52") And cmbconf.SelectedIndex = cmbconf.FindStringExact("4 Hojas Móviles") Then

                ' Calcula el ancho de la hoja móvil para carpinterías BW52 con 4 Hojas Móviles

                AHM = (pl / 4) + solBW52
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías BW52 con 4 Hojas Móviles

                HHM = hh - 5
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho de la hoja móvil lenta para carpinterías BW52 con 4 Hojas Móviles

                AHML = (pl / 4) + solBW52
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                ' Calcula la altura de la hoja móvil lenta para carpinterías BW52 con 4 Hojas Móviles

                HHML = hh
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías BW52 con 4 Hojas Móviles

                AVM = AHM - 119
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías BW52 con 4 Hojas Móviles

                HVM = HHM - desvBW52
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil lenta para carpinterías BW52 con 4 Hojas Móviles

                AVML = AHML - 129
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil lenta para carpinterías BW52 con 4 Hojas Móviles

                HVML = HHML - desvBW52
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Sacar mensaje de no aplicación de hojas fijas en este caso 

                AHFtext.Text = "No hay h.fijas"
                HHFtext.Text = "No hay h.fijas"
                AVFtext.Text = "No hay h.fijas"
                HVFtext.Text = "No hay h.fijas"

                d9 = HHM - 22
                d10 = HHM
                d11 = HHM
                d12 = HHM
                d13 = AHM - 20
                d14 = AHM + 24
                d15 = AHM - 13
                d16 = HHM - 113
                d17 = AHM - 154
                d18 = AHM - 20
                d19 = pl
                d20 = AHF

                If (AHF > AHM) Then
                    d21 = (2 * AHF) + pl
                ElseIf (AHF < AHM) Then
                    d21 = (2 * AHM) + pl
                ElseIf (AHF = AHM) Then
                    d21 = (2 * AHM) + pl
                End If
                d22 = d21
                d23 = HHM - 22
                d24 = HHM
                d25 = HHM
                d26 = HHM
                d27 = AHM - 20
                d28 = AHM + 24
                d29 = AHM - 13
                d30 = HHM - 113
                d31 = AHM - 154
                d32 = AHM - 20

                d1text.Text = "No hay h.fijas"
                d1text.Update()
                d2text.Text = "No hay h.fijas"
                d2text.Update()
                d3text.Text = "No hay h.fijas"
                d3text.Update()
                d4text.Text = "No hay h.fijas"
                d4text.Update()
                d5text.Text = "No hay h.fijas"
                d5text.Update()
                d6text.Text = "No hay h.fijas"
                d6text.Update()
                d7text.Text = "No hay h.fijas"
                d7text.Update()
                d8text.Text = "No hay h.fijas"
                d8text.Update()
                d9text.Text = d9.ToString()
                d9text.Update()
                d10text.Text = d10.ToString()
                d10text.Update()
                d11text.Text = d11.ToString()
                d11text.Update()
                d12text.Text = d12.ToString()
                d12text.Update()
                d13text.Text = d13.ToString()
                d13text.Update()
                d14text.Text = d14.ToString()
                d14text.Update()
                d15text.Text = d15.ToString()
                d15text.Update()
                d16text.Text = d16.ToString()
                d16text.Update()
                d17text.Text = d17.ToString()
                d17text.Update()
                d18text.Text = d18.ToString()
                d18text.Update()
                d19text.Text = d19.ToString()
                d19text.Update()
                d20text.Text = "Solo cuando hay fijos"
                d20text.Update()
                d21text.Text = d21.ToString()
                d21text.Update()
                d22text.Text = d22.ToString()
                d22text.Update()
                d23text.Text = d23.ToString()
                d23text.Update()
                d24text.Text = d24.ToString()
                d24text.Update()
                d25text.Text = "La hoja lenta no lleva este perfil"
                d25text.Update()
                d26text.Text = d26.ToString()
                d26text.Update()
                d27text.Text = d27.ToString()
                d27text.Update()
                d28text.Text = d28.ToString()
                d28text.Update()
                d29text.Text = d29.ToString()
                d29text.Update()
                d30text.Text = d30.ToString()
                d30text.Update()
                d31text.Text = d31.ToString()
                d31text.Update()
                d32text.Text = d32.ToString()
                d32text.Update()

                du9 = 4
                du10 = 2
                du11 = 2
                du12 = 2
                du13 = 2
                du14 = 2
                du15 = 2
                du16 = 4
                du17 = 4
                du18 = 2
                du19 = 1
                du20 = 2
                du21 = 1
                du22 = 1
                du23 = 4
                du24 = 4
                du25 = 1
                du26 = 4
                du27 = 2
                du28 = 2
                du29 = 2
                du30 = 4
                du31 = 4
                du32 = 2

                du1text.Text = "No hay h.fijas"
                du1text.Update()
                du2text.Text = "No hay h.fijas"
                du2text.Update()
                du3text.Text = "No hay h.fijas"
                du3text.Update()
                du4text.Text = "No hay h.fijas"
                du4text.Update()
                du5text.Text = "No hay h.fijas"
                du5text.Update()
                du6text.Text = "No hay h.fijas"
                du6text.Update()
                du7text.Text = "No hay h.fijas"
                du7text.Update()
                du8text.Text = "No hay h.fijas"
                du8text.Update()
                du9text.Text = du9.ToString()
                du9text.Update()
                du10text.Text = du10.ToString()
                du10text.Update()
                du11text.Text = du11.ToString()
                du11text.Update()
                du12text.Text = du12.ToString()
                du12text.Update()
                du13text.Text = du13.ToString()
                du13text.Update()
                du14text.Text = du14.ToString()
                du14text.Update()
                du15text.Text = du15.ToString()
                du15text.Update()
                du16text.Text = du16.ToString()
                du16text.Update()
                du17text.Text = du17.ToString()
                du17text.Update()
                du18text.Text = du18.ToString()
                du18text.Update()
                du19text.Text = du19.ToString()
                du19text.Update()
                du20text.Text = "Solo cuando hay fijos"
                du20text.Update()
                du21text.Text = du21.ToString()
                du21text.Update()
                du22text.Text = du22.ToString()
                du22text.Update()
                du23text.Text = du23.ToString()
                du23text.Update()
                du24text.Text = du24.ToString()
                du24text.Update()
                du25text.Text = "La hoja lenta no lleva este perfil"
                du25text.Update()
                du26text.Text = du26.ToString()
                du26text.Update()
                du27text.Text = du27.ToString()
                du27text.Update()
                du28text.Text = du28.ToString()
                du28text.Update()
                du29text.Text = du29.ToString()
                du29text.Update()
                du30text.Text = du30.ToString()
                du30text.Update()
                du31text.Text = du31.ToString()
                du31text.Update()
                du32text.Text = du32.ToString()
                du32text.Update()

                Panel3.Visible = False
                Panel4.Visible = False
                Panel5.Visible = False
                Panel6.Visible = True
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Panel14.Visible = True
                Label209.Visible = True
                TabControl2.SelectedIndex = 3

                '****************** BW52 - CASO VII ***************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("BW52") And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 2 Hojas Móviles") Then

                ' Calcula el ancho de la hoja fija para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                AHF = wh - pl
                AHFtext.Text = AHF.ToString()
                AHFtext.Update()

                ' Calcula la altura de la hoja fija para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                HHF = hh - 10
                HHFtext.Text = HHF.ToString()
                HHFtext.Update()

                ' Calcula el ancho de la hoja móvil para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                AHM = (pl / 2) + solBW52
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                HHM = hh - 5
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho de la hoja móvil lenta para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                AHML = (pl / 2) + solBW52
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                ' Calcula la altura de la hoja móvil lenta para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                HHML = hh
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                ' Calcula el ancho del vidrio de la hoja fija para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                AVF = AHF - 119
                AVFtext.Text = AVF.ToString()
                AVFtext.Update()

                ' Calcula la altura del vidrio de la hoja fija para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                HVF = HHF - desvBW52
                HVFtext.Text = HVF.ToString()
                HVFtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                AVM = AHM - 119
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                HVM = HHM - desvBW52
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil lenta para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                AVML = AHML - 129
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil lenta para carpinterías BW52 con 1 Hoja Fija + 2 Hojas Móviles

                HVML = HHML - desvBW52
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Calculo las medidas de los componentes de la carpintería BW52

                d1 = HHF - 22
                d2 = HHF + 8
                d3 = HHF + 30
                d4 = AHF - 10
                d5 = AHF + 34
                d6 = HHF - 113
                d7 = AHF - 144
                d8 = AHF - 10
                d9 = HHM - 22
                d10 = HHM
                d11 = HHM
                d12 = HHM
                d13 = AHM - 20
                d14 = AHM + 24
                d15 = AHM - 13
                d16 = HHM - 113
                d17 = AHM - 154
                d18 = AHM - 20
                d19 = pl
                d20 = AHF

                If (AHF > AHM) Then
                    d21 = AHF + pl
                ElseIf (AHF < AHM) Then
                    d21 = AHM + pl
                ElseIf (AHF = AHM) Then
                    d21 = AHF + pl
                End If
                d22 = d21
                d23 = HHM - 22
                d24 = HHM
                d25 = HHM
                d26 = HHM
                d27 = AHM - 20
                d28 = AHM + 24
                d29 = AHM - 13
                d30 = HHM - 113
                d31 = AHM - 154
                d32 = AHM - 20

                d1text.Text = d1.ToString()
                d1text.Update()
                d2text.Text = d2.ToString()
                d2text.Update()
                d3text.Text = d3.ToString()
                d3text.Update()
                d4text.Text = d4.ToString()
                d4text.Update()
                d5text.Text = d5.ToString()
                d5text.Update()
                d6text.Text = d6.ToString()
                d6text.Update()
                d7text.Text = d7.ToString()
                d7text.Update()
                d8text.Text = d8.ToString()
                d8text.Update()
                d9text.Text = d9.ToString()
                d9text.Update()
                d10text.Text = d10.ToString()
                d10text.Update()
                d11text.Text = d11.ToString()
                d11text.Update()
                d12text.Text = d12.ToString()
                d12text.Update()
                d13text.Text = d13.ToString()
                d13text.Update()
                d14text.Text = d14.ToString()
                d14text.Update()
                d15text.Text = d15.ToString()
                d15text.Update()
                d16text.Text = d16.ToString()
                d16text.Update()
                d17text.Text = d17.ToString()
                d17text.Update()
                d18text.Text = d18.ToString()
                d18text.Update()
                d19text.Text = d19.ToString()
                d19text.Update()
                d20text.Text = d20.ToString()
                d20text.Update()
                d21text.Text = d21.ToString()
                d21text.Update()
                d22text.Text = d22.ToString()
                d22text.Update()
                d23text.Text = d23.ToString()
                d23text.Update()
                d24text.Text = d24.ToString()
                d24text.Update()
                d25text.Text = "La hoja lenta no lleva este perfil"
                d25text.Update()
                d26text.Text = d26.ToString()
                d26text.Update()
                d27text.Text = d27.ToString()
                d27text.Update()
                d28text.Text = d28.ToString()
                d28text.Update()
                d29text.Text = d29.ToString()
                d29text.Update()
                d30text.Text = d30.ToString()
                d30text.Update()
                d31text.Text = d31.ToString()
                d31text.Update()
                d32text.Text = d32.ToString()
                d32text.Update()

                du1 = 2
                du2 = 1
                du3 = 1
                du4 = 1
                du5 = 1
                du6 = 2
                du7 = 2
                du8 = 1
                du9 = 2
                du10 = 1
                du11 = 1
                du12 = 1
                du13 = 1
                du14 = 1
                du15 = 1
                du16 = 2
                du17 = 2
                du18 = 1
                du19 = 1
                du20 = 1
                du21 = 1
                du22 = 1
                du23 = 2
                du24 = 2
                du25 = 1
                du26 = 2
                du27 = 1
                du28 = 1
                du29 = 1
                du30 = 2
                du31 = 2
                du32 = 1

                du1text.Text = du1.ToString()
                du1text.Update()
                du2text.Text = du2.ToString()
                du2text.Update()
                du3text.Text = du3.ToString()
                du3text.Update()
                du4text.Text = du4.ToString()
                du4text.Update()
                du5text.Text = du5.ToString()
                du5text.Update()
                du6text.Text = du6.ToString()
                du6text.Update()
                du7text.Text = du7.ToString()
                du7text.Update()
                du8text.Text = du8.ToString()
                du8text.Update()
                du9text.Text = du9.ToString()
                du9text.Update()
                du10text.Text = du10.ToString()
                du10text.Update()
                du11text.Text = du11.ToString()
                du11text.Update()
                du12text.Text = du12.ToString()
                du12text.Update()
                du13text.Text = du13.ToString()
                du13text.Update()
                du14text.Text = du14.ToString()
                du14text.Update()
                du15text.Text = du15.ToString()
                du15text.Update()
                du16text.Text = du16.ToString()
                du16text.Update()
                du17text.Text = du17.ToString()
                du17text.Update()
                du18text.Text = du18.ToString()
                du18text.Update()
                du19text.Text = du19.ToString()
                du19text.Update()
                du20text.Text = du20.ToString()
                du20text.Update()
                du21text.Text = du21.ToString()
                du21text.Update()
                du22text.Text = du22.ToString()
                du22text.Update()
                du23text.Text = du23.ToString()
                du23text.Update()
                du24text.Text = du24.ToString()
                du24text.Update()
                du25text.Text = "La hoja lenta no lleva este perfil"
                du25text.Update()
                du26text.Text = du26.ToString()
                du26text.Update()
                du27text.Text = du27.ToString()
                du27text.Update()
                du28text.Text = du28.ToString()
                du28text.Update()
                du29text.Text = du29.ToString()
                du29text.Update()
                du30text.Text = du30.ToString()
                du30text.Update()
                du31text.Text = du31.ToString()
                du31text.Update()
                du32text.Text = du32.ToString()
                du32text.Update()

                Panel3.Visible = False
                Panel4.Visible = False
                Panel5.Visible = False
                Panel6.Visible = True
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Panel14.Visible = True
                Label209.Visible = True
                TabControl2.SelectedIndex = 3

                '****************** BW52 - CASO VIII **************************************************************************************************************************

            ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("BW52") And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Móviles TES") Then

                ' Calcula el ancho de la hoja móvil para carpinterías BW52 con 2 Hojas Móviles

                AHM = (pl / 2) + solBW52
                AHMtext.Text = AHM.ToString()
                AHMtext.Update()

                ' Calcula la altura de la hoja móvil para carpinterías BW52 con 2 Hojas Móviles

                HHM = hh - 5
                HHMtext.Text = HHM.ToString()
                HHMtext.Update()

                ' Calcula el ancho de la hoja móvil lenta para carpinterías BW52 con 2 Hojas Móviles

                AHML = (pl / 2) + solBW52
                AHMLtext.Text = AHML.ToString()
                AHMLtext.Update()

                ' Calcula la altura de la hoja móvil lenta para carpinterías BW52 con 2 Hojas Móviles

                HHML = hh
                HHMLtext.Text = HHML.ToString()
                HHMLtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil para carpinterías BW52 con 2 Hojas Móviles

                AVM = AHM - 119
                AVMtext.Text = AVM.ToString()
                AVMtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil para carpinterías BW52 con 2 Hojas Móviles

                HVM = HHM - desvBW52
                HVMtext.Text = HVM.ToString()
                HVMtext.Update()

                ' Calcula el ancho del vidrio de la hoja móvil lenta para carpinterías BW52 con 2 Hojas Móviles

                AVML = AHML - 129
                AVMLtext.Text = AVML.ToString()
                AVMLtext.Update()

                ' Calcula la altura del vidrio de la hoja móvil lenta para carpinterías BW52 con 2 Hojas Móviles

                HVML = HHML - desvBW52
                HVMLtext.Text = HVML.ToString()
                HVMLtext.Update()

                ' Sacar mensaje de no aplicación de hojas fijas en este caso 

                AHFtext.Text = "No hay h.fijas"
                HHFtext.Text = "No hay h.fijas"
                AVFtext.Text = "No hay h.fijas"
                HVFtext.Text = "No hay h.fijas"

                d9 = HHM - 22
                d10 = HHM
                d11 = HHM
                d12 = HHM
                d13 = AHM - 20
                d14 = AHM + 24
                d15 = AHM - 13
                d16 = HHM - 113
                d17 = AHM - 154
                d18 = AHM - 20
                d19 = pl
                d20 = AHF

                If (AHF > AHM) Then
                    d21 = (AHF) + pl
                ElseIf (AHF < AHM) Then
                    d21 = (AHM) + pl
                ElseIf (AHF = AHM) Then
                    d21 = AHM + pl
                End If
                d22 = d21
                d23 = HHM - 22
                d24 = HHM
                d25 = HHM
                d26 = HHM
                d27 = AHM - 20
                d28 = AHM + 24
                d29 = AHM - 13
                d30 = HHM - 113
                d31 = AHM - 154
                d32 = AHM - 20

                d1text.Text = "No hay h.fijas"
                d1text.Update()
                d2text.Text = "No hay h.fijas"
                d2text.Update()
                d3text.Text = "No hay h.fijas"
                d3text.Update()
                d4text.Text = "No hay h.fijas"
                d4text.Update()
                d5text.Text = "No hay h.fijas"
                d5text.Update()
                d6text.Text = "No hay h.fijas"
                d6text.Update()
                d7text.Text = "No hay h.fijas"
                d7text.Update()
                d8text.Text = "No hay h.fijas"
                d8text.Update()
                d9text.Text = d9.ToString()
                d9text.Update()
                d10text.Text = d10.ToString()
                d10text.Update()
                d11text.Text = d11.ToString()
                d11text.Update()
                d12text.Text = d12.ToString()
                d12text.Update()
                d13text.Text = d13.ToString()
                d13text.Update()
                d14text.Text = d14.ToString()
                d14text.Update()
                d15text.Text = d15.ToString()
                d15text.Update()
                d16text.Text = d16.ToString()
                d16text.Update()
                d17text.Text = d17.ToString()
                d17text.Update()
                d18text.Text = d18.ToString()
                d18text.Update()
                d19text.Text = d19.ToString()
                d19text.Update()
                d20text.Text = "Solo cuando hay fijos"
                d20text.Update()
                d21text.Text = d21.ToString()
                d21text.Update()
                d22text.Text = d22.ToString()
                d22text.Update()
                d23text.Text = d23.ToString()
                d23text.Update()
                d24text.Text = d24.ToString()
                d24text.Update()
                d25text.Text = "La hoja lenta no lleva este perfil"
                d25text.Update()
                d26text.Text = d26.ToString()
                d26text.Update()
                d27text.Text = d27.ToString()
                d27text.Update()
                d28text.Text = d28.ToString()
                d28text.Update()
                d29text.Text = d29.ToString()
                d29text.Update()
                d30text.Text = d30.ToString()
                d30text.Update()
                d31text.Text = d31.ToString()
                d31text.Update()
                d32text.Text = d32.ToString()
                d32text.Update()

                du9 = 2
                du10 = 1
                du11 = 1
                du12 = 1
                du13 = 1
                du14 = 1
                du15 = 1
                du16 = 2
                du17 = 2
                du18 = 1
                du19 = 1
                du20 = 1
                du21 = 1
                du22 = 1
                du23 = 2
                du24 = 2
                du25 = 1
                du26 = 2
                du27 = 1
                du28 = 1
                du29 = 1
                du30 = 2
                du31 = 2
                du32 = 1

                du1text.Text = "No hay h.fijas"
                du1text.Update()
                du2text.Text = "No hay h.fijas"
                du2text.Update()
                du3text.Text = "No hay h.fijas"
                du3text.Update()
                du4text.Text = "No hay h.fijas"
                du4text.Update()
                du5text.Text = "No hay h.fijas"
                du5text.Update()
                du6text.Text = "No hay h.fijas"
                du6text.Update()
                du7text.Text = "No hay h.fijas"
                du7text.Update()
                du8text.Text = "No hay h.fijas"
                du8text.Update()
                du9text.Text = du9.ToString()
                du9text.Update()
                du10text.Text = du10.ToString()
                du10text.Update()
                du11text.Text = du11.ToString()
                du11text.Update()
                du12text.Text = du12.ToString()
                du12text.Update()
                du13text.Text = du13.ToString()
                du13text.Update()
                du14text.Text = du14.ToString()
                du14text.Update()
                du15text.Text = du15.ToString()
                du15text.Update()
                du16text.Text = du16.ToString()
                du16text.Update()
                du17text.Text = du17.ToString()
                du17text.Update()
                du18text.Text = du18.ToString()
                du18text.Update()
                du19text.Text = du19.ToString()
                du19text.Update()
                du20text.Text = "Solo cuando hay fijos"
                du20text.Update()
                du21text.Text = du21.ToString()
                du21text.Update()
                du22text.Text = du22.ToString()
                du22text.Update()
                du23text.Text = du23.ToString()
                du23text.Update()
                du24text.Text = du24.ToString()
                du24text.Update()
                du25text.Text = "La hoja lenta no lleva este perfil"
                du25text.Update()
                du26text.Text = du26.ToString()
                du26text.Update()
                du27text.Text = du27.ToString()
                du27text.Update()
                du28text.Text = du28.ToString()
                du28text.Update()
                du29text.Text = du29.ToString()
                du29text.Update()
                du30text.Text = du30.ToString()
                du30text.Update()
                du31text.Text = du31.ToString()
                du31text.Update()
                du32text.Text = du32.ToString()
                du32text.Update()

                Panel3.Visible = False
                Panel4.Visible = False
                Panel5.Visible = False
                Panel6.Visible = True
                Label105.Visible = True
                AHMLtext.Visible = True
                HHMLtext.Visible = True
                Label106.Visible = True
                AVMLtext.Visible = True
                HVMLtext.Visible = True
                Label107.Visible = True
                Label108.Visible = True
                Panel14.Visible = True
                Label209.Visible = True
                TabControl2.SelectedIndex = 3

                '****************** AP94 - CASO I *****************************************************************************************************************************

                'ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("AP94") And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Fijas + 2 Hojas Móviles") Then

                '    If (wh > 4380 Or pl > 2050) Then

                '        MsgBox("El ancho de hueco y/o el paso libre introducidos están fuera de rango. Consulte con su comercial para soluciones personalizadas")

                '        Return

                '    End If

                '    ' Calcula el ancho de la hoja móvil para carpinterías AP94 con 2 hojas fijas y 2 móviles

                '    AHM = (pl / 2)
                '    AHMtext.Text = AHM.ToString()
                '    AHMtext.Update()

                '    ' Calcula el ancho de la hoja fija para carpinterías AP94 con 2 hojas fijas y 2 móviles

                '    AHF = AHM + 70
                '    AHFtext.Text = AHF.ToString()
                '    AHFtext.Update()

                '    ' Calcula la altura de la hoja fija para carpinterías AP94 con 2 hojas fijas y 2 móviles

                '    HHF = hh - 20
                '    HHFtext.Text = HHF.ToString()
                '    HHFtext.Update()

                '    ' Calcula la altura de la hoja móvil para carpinterías AP94 con 2 hojas fijas y 2 móviles

                '    HHM = HHF - 25
                '    HHMtext.Text = HHM.ToString()
                '    HHMtext.Update()

                '    ' Calcula el ancho del vidrio de la hoja fija para carpinterías AP94 con 2 hojas fijas y 2 móviles

                '    AVF = AHF - 209
                '    AVFtext.Text = AVF.ToString()
                '    AVFtext.Update()

                '    ' Calcula la altura del vidrio de la hoja fija para carpinterías AP94 con 2 hojas fijas y 2 móviles

                '    HVF = HHF - 164
                '    HVFtext.Text = HVF.ToString()
                '    HVFtext.Update()

                '    ' Calcula el ancho del vidrio de la hoja móvil para carpinterías AP94 con 2 hojas fijas y 2 móviles

                '    AVM = AHM - 127
                '    AVMtext.Text = AVM.ToString()
                '    AVMtext.Update()

                '    ' Calcula la altura del vidrio de la hoja móvil para carpinterías AP94 con 2 hojas fijas y 2 móviles

                '    HVM = HHM - 159
                '    HVMtext.Text = HVM.ToString()
                '    HVMtext.Update()

                '    ' Calculo las medidas de los componentes de la carpintería AP94

                '    e1 = HHF
                '    e2 = HHF - 46
                '    e3 = HHF - 41
                '    e4 = AHF - 46
                '    e5 = AHF - 59
                '    e6 = HHF - 146
                '    e7 = AHF - 53
                '    e8 = AHF - 190
                '    e9 = AHF - 195
                '    e10 = AHF - 190
                '    e11 = AHF - 235
                '    e12 = HHM - 41
                '    e13 = HHM - 54
                '    e14 = HHM - 54
                '    e15 = HHM - 54
                '    e16 = HHM - 141
                '    e17 = AHM - 17
                '    e18 = AHM - 108
                '    e19 = AHM - 113
                '    e20 = AHM - 108
                '    e21 = AHM - 153
                '    e22 = HHM - 41
                '    e23 = HHM - 41
                '    e24 = wh
                '    e25 = pl
                '    e26 = pl
                '    e27 = AHF
                '    e28 = wh
                '    e29 = wh

                '    e1text.Text = e1.ToString()
                '    e1text.Update()
                '    e2text.Text = e2.ToString()
                '    e2text.Update()
                '    e3text.Text = e3.ToString()
                '    e3text.Update()
                '    e4text.Text = e4.ToString()
                '    e4text.Update()
                '    e5text.Text = e5.ToString()
                '    e5text.Update()
                '    e6text.Text = e6.ToString()
                '    e6text.Update()
                '    e7text.Text = e7.ToString()
                '    e7text.Update()
                '    e8text.Text = e8.ToString()
                '    e8text.Update()
                '    e9text.Text = e9.ToString()
                '    e9text.Update()
                '    e10text.Text = e10.ToString()
                '    e10text.Update()
                '    e11text.Text = e11.ToString()
                '    e11text.Update()
                '    e12text.Text = e12.ToString()
                '    e12text.Update()
                '    e13text.Text = e13.ToString()
                '    e13text.Update()
                '    e14text.Text = e14.ToString()
                '    e14text.Update()
                '    e15text.Text = e15.ToString()
                '    e15text.Update()
                '    e16text.Text = e16.ToString()
                '    e16text.Update()
                '    e17text.Text = e17.ToString()
                '    e17text.Update()
                '    e18text.Text = e18.ToString()
                '    e18text.Update()
                '    e19text.Text = e19.ToString()
                '    e19text.Update()
                '    e20text.Text = e20.ToString()
                '    e20text.Update()
                '    e21text.Text = e21.ToString()
                '    e21text.Update()
                '    e22text.Text = e22.ToString()
                '    e22text.Update()
                '    e23text.Text = e23.ToString()
                '    e23text.Update()
                '    e24text.Text = e24.ToString()
                '    e24text.Update()
                '    e25text.Text = e25.ToString()
                '    e25text.Update()
                '    e26text.Text = e26.ToString()
                '    e26text.Update()
                '    e27text.Text = e27.ToString()
                '    e27text.Update()
                '    e28text.Text = e28.ToString()
                '    e28text.Update()
                '    e29text.Text = e29.ToString()
                '    e29text.Update()

                '    eu1 = 4
                '    eu2 = 2
                '    eu3 = 2
                '    eu4 = 2
                '    eu5 = 2
                '    eu6 = 4
                '    eu7 = 4
                '    eu8 = 2
                '    eu9 = 4
                '    eu10 = 2
                '    eu11 = 2
                '    eu12 = 2
                '    eu13 = 2
                '    eu14 = 2
                '    eu15 = 2
                '    eu16 = 4
                '    eu17 = 4
                '    eu18 = 2
                '    eu19 = 1
                '    eu20 = 2
                '    eu21 = 1
                '    eu22 = 1
                '    eu23 = 2
                '    eu24 = 1
                '    eu25 = 1
                '    eu26 = 2
                '    eu27 = 1
                '    eu28 = 1
                '    eu29 = 87

                '    eu1text.Text = eu1.ToString()
                '    eu1text.Update()
                '    eu2text.Text = eu2.ToString()
                '    eu2text.Update()
                '    eu3text.Text = eu3.ToString()
                '    eu3text.Update()
                '    eu4text.Text = eu4.ToString()
                '    eu4text.Update()
                '    eu5text.Text = eu5.ToString()
                '    eu5text.Update()
                '    eu6text.Text = eu6.ToString()
                '    eu6text.Update()
                '    eu7text.Text = eu7.ToString()
                '    eu7text.Update()
                '    eu8text.Text = eu8.ToString()
                '    eu8text.Update()
                '    eu9text.Text = eu9.ToString()
                '    eu9text.Update()
                '    eu10text.Text = eu10.ToString()
                '    eu10text.Update()
                '    eu11text.Text = eu11.ToString()
                '    eu11text.Update()
                '    eu12text.Text = eu12.ToString()
                '    eu12text.Update()
                '    eu13text.Text = eu13.ToString()
                '    eu13text.Update()
                '    eu14text.Text = eu14.ToString()
                '    eu14text.Update()
                '    eu15text.Text = eu15.ToString()
                '    eu15text.Update()
                '    eu16text.Text = eu16.ToString()
                '    eu16text.Update()
                '    eu17text.Text = eu17.ToString()
                '    eu17text.Update()
                '    eu18text.Text = eu18.ToString()
                '    eu18text.Update()
                '    eu19text.Text = eu19.ToString()
                '    eu19text.Update()
                '    eu20text.Text = eu20.ToString()
                '    eu20text.Update()
                '    eu21text.Text = eu21.ToString()
                '    eu21text.Update()
                '    eu22text.Text = eu22.ToString()
                '    eu22text.Update()
                '    eu23text.Text = eu23.ToString()
                '    eu23text.Update()
                '    eu24text.Text = eu24.ToString()
                '    eu24text.Update()
                '    eu25text.Text = eu25.ToString()
                '    eu25text.Update()
                '    eu26text.Text = eu26.ToString()
                '    eu26text.Update()
                '    eu27text.Text = eu27.ToString()
                '    eu27text.Update()
                '    eu28text.Text = eu28.ToString()
                '    eu28text.Update()
                '    eu29text.Text = eu29.ToString()
                '    eu29text.Update()

                '    Panel3.Visible = False
                '    Panel4.Visible = False
                '    Panel5.Visible = False
                '    Panel6.Visible = False
                '    Label105.Visible = False
                '    AHMLtext.Visible = False
                '    HHMLtext.Visible = False
                '    Label106.Visible = False
                '    AVMLtext.Visible = False
                '    HVMLtext.Visible = False
                '    Label107.Visible = False
                '    Label108.Visible = False
                '    Label143.Visible = False
                '    Label146.Visible = False
                '    e22text.Visible = False
                '    e23text.Visible = False
                '    eu22text.Visible = False
                '    eu23text.Visible = False
                '    TabControl2.SelectedIndex = 4

                '    '****************** AP94 - CASO II *****************************************************************************************************************************


                'ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("AP94") And cmbconf.SelectedIndex = cmbconf.FindStringExact("2 Hojas Móviles") Then

                '    If (pl > 2050) Then

                '        MsgBox("El ancho de hueco y/o el paso libre introducidos están fuera de rango. Consulte con su comercial para soluciones personalizadas")

                '        Return

                '    End If

                '    ' Calcula el ancho de la hoja móvil para carpinterías AP94 con 2 hojas móviles

                '    AHM = (pl / 2) + 45
                '    AHMtext.Text = AHM.ToString()
                '    AHMtext.Update()

                '    ' Calcula la altura de la hoja móvil para carpinterías AP94 con 2 hojas móviles

                '    HHM = hh - 10
                '    HHMtext.Text = HHM.ToString()
                '    HHMtext.Update()

                '    ' Calcula el ancho del vidrio de la hoja móvil para carpinterías AP94 con 2 hojas móviles

                '    AVM = AHM - 174
                '    AVMtext.Text = AVM.ToString()
                '    AVMtext.Update()

                '    ' Calcula la altura del vidrio de la hoja móvil para carpinterías AP94 con 2 hojas móviles

                '    HVM = HHM - 159
                '    HVMtext.Text = HVM.ToString()
                '    HVMtext.Update()

                '    ' Calculo las medidas de los componentes de la carpintería AP94

                '    e12 = HHM - 41
                '    e13 = HHM - 54
                '    e14 = HHM - 54
                '    e16 = HHM - 141
                '    e17 = AHM - 19
                '    e18 = AHM - 155
                '    e19 = AHM - 160
                '    e20 = AHM - 155
                '    e21 = AHM - 200
                '    e22 = HHM - 41
                '    e23 = HHM - 41
                '    e27 = AHM + 70
                '    e28 = pl + 2 * AHM
                '    e29 = pl + 2 * AHM

                '    e1text.Text = "No hay h.fijas"
                '    e1text.Update()
                '    e2text.Text = "No hay h.fijas"
                '    e2text.Update()
                '    e3text.Text = "No hay h.fijas"
                '    e3text.Update()
                '    e4text.Text = "No hay h.fijas"
                '    e4text.Update()
                '    e5text.Text = "No hay h.fijas"
                '    e5text.Update()
                '    e6text.Text = "No hay h.fijas"
                '    e6text.Update()
                '    e7text.Text = "No hay h.fijas"
                '    e7text.Update()
                '    e8text.Text = "No hay h.fijas"
                '    e8text.Update()
                '    e9text.Text = "No hay h.fijas"
                '    e9text.Update()
                '    e10text.Text = "No hay h.fijas"
                '    e10text.Update()
                '    e11text.Text = "No hay h.fijas"
                '    e11text.Update()
                '    e12text.Text = e12.ToString()
                '    e12text.Update()
                '    e13text.Text = e13.ToString()
                '    e13text.Update()
                '    e14text.Text = e14.ToString()
                '    e14text.Update()
                '    e15text.Text = e15.ToString()
                '    e15text.Update()
                '    e16text.Text = e16.ToString()
                '    e16text.Update()
                '    e17text.Text = e17.ToString()
                '    e17text.Update()
                '    e18text.Text = e18.ToString()
                '    e18text.Update()
                '    e19text.Text = e19.ToString()
                '    e19text.Update()
                '    e20text.Text = e20.ToString()
                '    e20text.Update()
                '    e21text.Text = e21.ToString()
                '    e21text.Update()
                '    e22text.Text = e22.ToString()
                '    e22text.Update()
                '    e23text.Text = e23.ToString()
                '    e23text.Update()
                '    e27text.Text = e27.ToString()
                '    e27text.Update()
                '    e28text.Text = e28.ToString()
                '    e28text.Update()
                '    e29text.Text = e29.ToString()
                '    e29text.Update()

                '    eu12 = 2
                '    eu13 = 2
                '    eu14 = 2
                '    eu15 = 2
                '    eu16 = 4
                '    eu17 = 4
                '    eu18 = 2
                '    eu19 = 1
                '    eu20 = 2
                '    eu21 = 1
                '    eu22 = 1
                '    eu23 = 2
                '    eu24 = 1
                '    eu25 = 1
                '    eu26 = 2
                '    eu27 = 1
                '    eu28 = 1
                '    eu29 = 87

                '    eu1text.Text = "No hay h.fijas"
                '    eu1text.Update()
                '    eu2text.Text = "No hay h.fijas"
                '    eu2text.Update()
                '    eu3text.Text = "No hay h.fijas"
                '    eu3text.Update()
                '    eu4text.Text = "No hay h.fijas"
                '    eu4text.Update()
                '    eu5text.Text = "No hay h.fijas"
                '    eu5text.Update()
                '    eu6text.Text = "No hay h.fijas"
                '    eu6text.Update()
                '    eu7text.Text = "No hay h.fijas"
                '    eu7text.Update()
                '    eu8text.Text = "No hay h.fijas"
                '    eu8text.Update()
                '    eu9text.Text = "No hay h.fijas"
                '    eu9text.Update()
                '    eu10text.Text = "No hay h.fijas"
                '    eu10text.Update()
                '    eu11text.Text = "No hay h.fijas"
                '    eu11text.Update()
                '    eu12text.Text = eu12.ToString()
                '    eu12text.Update()
                '    eu13text.Text = eu13.ToString()
                '    eu13text.Update()
                '    eu14text.Text = eu14.ToString()
                '    eu14text.Update()
                '    eu15text.Text = eu15.ToString()
                '    eu15text.Update()
                '    eu16text.Text = eu16.ToString()
                '    eu16text.Update()
                '    eu17text.Text = eu17.ToString()
                '    eu17text.Update()
                '    eu18text.Text = eu18.ToString()
                '    eu18text.Update()
                '    eu19text.Text = eu19.ToString()
                '    eu19text.Update()
                '    eu20text.Text = eu20.ToString()
                '    eu20text.Update()
                '    eu21text.Text = eu21.ToString()
                '    eu21text.Update()
                '    eu22text.Text = eu22.ToString()
                '    eu22text.Update()
                '    eu23text.Text = eu23.ToString()
                '    eu23text.Update()
                '    eu27text.Text = eu27.ToString()
                '    eu27text.Update()
                '    eu28text.Text = eu28.ToString()
                '    eu28text.Update()
                '    eu29text.Text = eu29.ToString()
                '    eu29text.Update()

                '    Panel3.Visible = False
                '    Panel4.Visible = False
                '    Panel5.Visible = False
                '    Panel6.Visible = False
                '    Label105.Visible = False
                '    AHMLtext.Visible = False
                '    HHMLtext.Visible = False
                '    Label106.Visible = False
                '    AVMLtext.Visible = False
                '    HVMLtext.Visible = False
                '    Label133.Visible = False
                '    Label132.Visible = False
                '    Label144.Visible = False
                '    e24text.Visible = False
                '    e25text.Visible = False
                '    e26text.Visible = False
                '    eu24text.Visible = False
                '    eu25text.Visible = False
                '    eu26text.Visible = False
                '    Label107.Visible = False
                '    Label108.Visible = False
                '    Label143.Visible = True
                '    Label146.Visible = True
                '    Label136.Visible = False
                '    e15text.Visible = False
                '    eu15text.Visible = False
                '    e22text.Visible = True
                '    e23text.Visible = True
                '    eu22text.Visible = True
                '    eu23text.Visible = True
                '    TabControl2.SelectedIndex = 4


                '    '****************** AP94 - CASO III ****************************************************************************************************************************

                'ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("AP94") And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Fija + 1 Hoja Móvil") Then

                '    If (wh > 2190 Or pl > 1025) Then

                '        MsgBox("El ancho de hueco y/o el paso libre introducidos están fuera de rango. Consulte con su comercial para soluciones personalizadas")

                '        Return

                '    End If

                '    ' Calcula el ancho de la hoja móvil para carpinterías AP94 con 1 hoja fija y 1 móvil

                '    AHM = (wh / 2)
                '    AHMtext.Text = AHM.ToString()
                '    AHMtext.Update()

                '    ' Calcula el ancho de la hoja fija para carpinterías AP94 con 1 hoja fija y 1 móvil

                '    AHF = AHM + 70
                '    AHFtext.Text = AHF.ToString()
                '    AHFtext.Update()

                '    ' Calcula la altura de la hoja fija para carpinterías AP94 con 1 hoja fija y 1 móvil

                '    HHF = hh - 20
                '    HHFtext.Text = HHF.ToString()
                '    HHFtext.Update()

                '    ' Calcula la altura de la hoja móvil para carpinterías AP94 con 1 hoja fija y 1 móvil

                '    HHM = HHF - 25
                '    HHMtext.Text = HHM.ToString()
                '    HHMtext.Update()

                '    ' Calcula el ancho del vidrio de la hoja fija para carpinterías AP94 con 1 hoja fija y 1 móvil

                '    AVF = AHF - 209
                '    AVFtext.Text = AVF.ToString()
                '    AVFtext.Update()

                '    ' Calcula la altura del vidrio de la hoja fija para carpinterías AP94 con 1 hoja fija y 1 móvil

                '    HVF = HHF - 164
                '    HVFtext.Text = HVF.ToString()
                '    HVFtext.Update()

                '    ' Calcula el ancho del vidrio de la hoja móvil para carpinterías AP94 con 1 hoja fija y 1 móvil

                '    AVM = AHM - 127
                '    AVMtext.Text = AVM.ToString()
                '    AVMtext.Update()

                '    ' Calcula la altura del vidrio de la hoja móvil para carpinterías AP94 con 1 hoja fija y 1 móvil

                '    HVM = HHM - 159
                '    HVMtext.Text = HVM.ToString()
                '    HVMtext.Update()

                '    ' Calculo las medidas de los componentes de la carpintería AP94

                '    e1 = HHF
                '    e2 = HHF - 46
                '    e3 = HHF - 41
                '    e4 = AHF - 46
                '    e5 = AHF - 59
                '    e6 = HHF - 146
                '    e7 = AHF - 53
                '    e8 = AHF - 190
                '    e9 = AHF - 195
                '    e10 = AHF - 190
                '    e11 = AHF - 235
                '    e12 = HHM - 41
                '    e13 = HHM - 54
                '    e14 = HHM - 54
                '    e15 = HHM - 54
                '    e16 = HHM - 141
                '    e17 = AHM - 17
                '    e18 = AHM - 108
                '    e19 = AHM - 113
                '    e20 = AHM - 108
                '    e21 = AHM - 153
                '    e22 = HHM - 41
                '    e23 = HHM - 41
                '    e24 = wh
                '    e25 = wh - AHF
                '    e26 = wh - AHF
                '    e27 = AHF
                '    e28 = wh
                '    e29 = wh

                '    e1text.Text = e1.ToString()
                '    e1text.Update()
                '    e2text.Text = e2.ToString()
                '    e2text.Update()
                '    e3text.Text = e3.ToString()
                '    e3text.Update()
                '    e4text.Text = e4.ToString()
                '    e4text.Update()
                '    e5text.Text = e5.ToString()
                '    e5text.Update()
                '    e6text.Text = e6.ToString()
                '    e6text.Update()
                '    e7text.Text = e7.ToString()
                '    e7text.Update()
                '    e8text.Text = e8.ToString()
                '    e8text.Update()
                '    e9text.Text = e9.ToString()
                '    e9text.Update()
                '    e10text.Text = e10.ToString()
                '    e10text.Update()
                '    e11text.Text = e11.ToString()
                '    e11text.Update()
                '    e12text.Text = e12.ToString()
                '    e12text.Update()
                '    e13text.Text = e13.ToString()
                '    e13text.Update()
                '    e14text.Text = e14.ToString()
                '    e14text.Update()
                '    e15text.Text = e15.ToString()
                '    e15text.Update()
                '    e16text.Text = e16.ToString()
                '    e16text.Update()
                '    e17text.Text = e17.ToString()
                '    e17text.Update()
                '    e18text.Text = e18.ToString()
                '    e18text.Update()
                '    e19text.Text = e19.ToString()
                '    e19text.Update()
                '    e20text.Text = e20.ToString()
                '    e20text.Update()
                '    e21text.Text = e21.ToString()
                '    e21text.Update()
                '    e22text.Text = e22.ToString()
                '    e22text.Update()
                '    e23text.Text = e23.ToString()
                '    e23text.Update()
                '    e24text.Text = e24.ToString()
                '    e24text.Update()
                '    e25text.Text = e25.ToString()
                '    e25text.Update()
                '    e26text.Text = e26.ToString()
                '    e26text.Update()
                '    e27text.Text = e27.ToString()
                '    e27text.Update()
                '    e28text.Text = e28.ToString()
                '    e28text.Update()
                '    e29text.Text = e29.ToString()
                '    e29text.Update()

                '    eu1 = 4
                '    eu2 = 2
                '    eu3 = 2
                '    eu4 = 2
                '    eu5 = 2
                '    eu6 = 4
                '    eu7 = 4
                '    eu8 = 2
                '    eu9 = 4
                '    eu10 = 2
                '    eu11 = 2
                '    eu12 = 2
                '    eu13 = 2
                '    eu14 = 2
                '    eu15 = 2
                '    eu16 = 4
                '    eu17 = 4
                '    eu18 = 2
                '    eu19 = 1
                '    eu20 = 2
                '    eu21 = 1
                '    eu22 = 1
                '    eu23 = 2
                '    eu24 = 1
                '    eu25 = 1
                '    eu26 = 2
                '    eu27 = 1
                '    eu28 = 1
                '    eu29 = 87

                '    eu1text.Text = eu1.ToString()
                '    eu1text.Update()
                '    eu2text.Text = eu2.ToString()
                '    eu2text.Update()
                '    eu3text.Text = eu3.ToString()
                '    eu3text.Update()
                '    eu4text.Text = eu4.ToString()
                '    eu4text.Update()
                '    eu5text.Text = eu5.ToString()
                '    eu5text.Update()
                '    eu6text.Text = eu6.ToString()
                '    eu6text.Update()
                '    eu7text.Text = eu7.ToString()
                '    eu7text.Update()
                '    eu8text.Text = eu8.ToString()
                '    eu8text.Update()
                '    eu9text.Text = eu9.ToString()
                '    eu9text.Update()
                '    eu10text.Text = eu10.ToString()
                '    eu10text.Update()
                '    eu11text.Text = eu11.ToString()
                '    eu11text.Update()
                '    eu12text.Text = eu12.ToString()
                '    eu12text.Update()
                '    eu13text.Text = eu13.ToString()
                '    eu13text.Update()
                '    eu14text.Text = eu14.ToString()
                '    eu14text.Update()
                '    eu15text.Text = eu15.ToString()
                '    eu15text.Update()
                '    eu16text.Text = eu16.ToString()
                '    eu16text.Update()
                '    eu17text.Text = eu17.ToString()
                '    eu17text.Update()
                '    eu18text.Text = eu18.ToString()
                '    eu18text.Update()
                '    eu19text.Text = eu19.ToString()
                '    eu19text.Update()
                '    eu20text.Text = eu20.ToString()
                '    eu20text.Update()
                '    eu21text.Text = eu21.ToString()
                '    eu21text.Update()
                '    eu22text.Text = eu22.ToString()
                '    eu22text.Update()
                '    eu23text.Text = eu23.ToString()
                '    eu23text.Update()
                '    eu24text.Text = eu24.ToString()
                '    eu24text.Update()
                '    eu25text.Text = eu25.ToString()
                '    eu25text.Update()
                '    eu26text.Text = eu26.ToString()
                '    eu26text.Update()
                '    eu27text.Text = eu27.ToString()
                '    eu27text.Update()
                '    eu28text.Text = eu28.ToString()
                '    eu28text.Update()
                '    eu29text.Text = eu29.ToString()
                '    eu29text.Update()

                '    Panel3.Visible = False
                '    Panel4.Visible = False
                '    Panel5.Visible = False
                '    Panel6.Visible = False
                '    Label105.Visible = False
                '    AHMLtext.Visible = False
                '    HHMLtext.Visible = False
                '    Label106.Visible = False
                '    AVMLtext.Visible = False
                '    HVMLtext.Visible = False
                '    Label107.Visible = False
                '    Label108.Visible = False
                '    Label143.Visible = False
                '    Label146.Visible = False
                '    e22text.Visible = False
                '    e23text.Visible = False
                '    eu22text.Visible = False
                '    eu23text.Visible = False
                '    TabControl2.SelectedIndex = 4

                '    '****************** AP94 - CASO IV *****************************************************************************************************************************


                'ElseIf cmbmodcar.SelectedIndex = cmbmodcar.FindStringExact("AP94") And cmbconf.SelectedIndex = cmbconf.FindStringExact("1 Hoja Móvil") Then

                '    If (pl > 1025) Then

                '        MsgBox("El ancho de hueco y/o el paso libre introducidos están fuera de rango. Consulte con su comercial para soluciones personalizadas")

                '        Return

                '    End If

                '    ' Calcula el ancho de la hoja móvil para carpinterías AP94 con 1 hoja móvil

                '    AHM = pl / 2 + 45
                '    AHMtext.Text = AHM.ToString()
                '    AHMtext.Update()

                '    ' Calcula la altura de la hoja móvil para carpinterías AP94 con 1 hoja móvil

                '    HHM = hh - 10
                '    HHMtext.Text = HHM.ToString()
                '    HHMtext.Update()

                '    ' Calcula el ancho del vidrio de la hoja móvil para carpinterías AP94 con 1 hoja móvil

                '    AVM = AHM - 174
                '    AVMtext.Text = AVM.ToString()
                '    AVMtext.Update()

                '    ' Calcula la altura del vidrio de la hoja móvil para carpinterías AP94 con 1 hoja móvil

                '    HVM = HHM - 159
                '    HVMtext.Text = HVM.ToString()
                '    HVMtext.Update()

                '    ' Calculo las medidas de los componentes de la carpintería AP94

                '    e12 = HHM - 41
                '    e13 = HHM - 54
                '    e14 = HHM - 54
                '    e16 = HHM - 141
                '    e17 = AHM - 19
                '    e18 = AHM - 155
                '    e19 = AHM - 160
                '    e20 = AHM - 155
                '    e21 = AHM - 200
                '    e22 = HHM - 41
                '    e23 = HHM - 41
                '    e27 = AHM + 70
                '    e28 = pl + AHM
                '    e29 = pl + AHM

                '    e1text.Text = "No hay h.fijas"
                '    e1text.Update()
                '    e2text.Text = "No hay h.fijas"
                '    e2text.Update()
                '    e3text.Text = "No hay h.fijas"
                '    e3text.Update()
                '    e4text.Text = "No hay h.fijas"
                '    e4text.Update()
                '    e5text.Text = "No hay h.fijas"
                '    e5text.Update()
                '    e6text.Text = "No hay h.fijas"
                '    e6text.Update()
                '    e7text.Text = "No hay h.fijas"
                '    e7text.Update()
                '    e8text.Text = "No hay h.fijas"
                '    e8text.Update()
                '    e9text.Text = "No hay h.fijas"
                '    e9text.Update()
                '    e10text.Text = "No hay h.fijas"
                '    e10text.Update()
                '    e11text.Text = "No hay h.fijas"
                '    e11text.Update()
                '    e12text.Text = e12.ToString()
                '    e12text.Update()
                '    e13text.Text = e13.ToString()
                '    e13text.Update()
                '    e14text.Text = e14.ToString()
                '    e14text.Update()
                '    e15text.Text = e15.ToString()
                '    e15text.Update()
                '    e16text.Text = e16.ToString()
                '    e16text.Update()
                '    e17text.Text = e17.ToString()
                '    e17text.Update()
                '    e18text.Text = e18.ToString()
                '    e18text.Update()
                '    e19text.Text = e19.ToString()
                '    e19text.Update()
                '    e20text.Text = e20.ToString()
                '    e20text.Update()
                '    e21text.Text = e21.ToString()
                '    e21text.Update()
                '    e22text.Text = e22.ToString()
                '    e22text.Update()
                '    e23text.Text = e23.ToString()
                '    e23text.Update()
                '    e27text.Text = e27.ToString()
                '    e27text.Update()
                '    e28text.Text = e28.ToString()
                '    e28text.Update()
                '    e29text.Text = e29.ToString()
                '    e29text.Update()

                '    eu12 = 2
                '    eu13 = 2
                '    eu14 = 2
                '    eu15 = 2
                '    eu16 = 4
                '    eu17 = 4
                '    eu18 = 2
                '    eu19 = 1
                '    eu20 = 2
                '    eu21 = 1
                '    eu22 = 1
                '    eu23 = 2
                '    eu24 = 1
                '    eu25 = 1
                '    eu26 = 2
                '    eu27 = 1
                '    eu28 = 1
                '    eu29 = 87

                '    eu1text.Text = "No hay h.fijas"
                '    eu1text.Update()
                '    eu2text.Text = "No hay h.fijas"
                '    eu2text.Update()
                '    eu3text.Text = "No hay h.fijas"
                '    eu3text.Update()
                '    eu4text.Text = "No hay h.fijas"
                '    eu4text.Update()
                '    eu5text.Text = "No hay h.fijas"
                '    eu5text.Update()
                '    eu6text.Text = "No hay h.fijas"
                '    eu6text.Update()
                '    eu7text.Text = "No hay h.fijas"
                '    eu7text.Update()
                '    eu8text.Text = "No hay h.fijas"
                '    eu8text.Update()
                '    eu9text.Text = "No hay h.fijas"
                '    eu9text.Update()
                '    eu10text.Text = "No hay h.fijas"
                '    eu10text.Update()
                '    eu11text.Text = "No hay h.fijas"
                '    eu11text.Update()
                '    eu12text.Text = eu12.ToString()
                '    eu12text.Update()
                '    eu13text.Text = eu13.ToString()
                '    eu13text.Update()
                '    eu14text.Text = eu14.ToString()
                '    eu14text.Update()
                '    eu15text.Text = eu15.ToString()
                '    eu15text.Update()
                '    eu16text.Text = eu16.ToString()
                '    eu16text.Update()
                '    eu17text.Text = eu17.ToString()
                '    eu17text.Update()
                '    eu18text.Text = eu18.ToString()
                '    eu18text.Update()
                '    eu19text.Text = eu19.ToString()
                '    eu19text.Update()
                '    eu20text.Text = eu20.ToString()
                '    eu20text.Update()
                '    eu21text.Text = eu21.ToString()
                '    eu21text.Update()
                '    eu22text.Text = eu22.ToString()
                '    eu22text.Update()
                '    eu23text.Text = eu23.ToString()
                '    eu23text.Update()
                '    eu27text.Text = eu27.ToString()
                '    eu27text.Update()
                '    eu28text.Text = eu28.ToString()
                '    eu28text.Update()
                '    eu29text.Text = eu29.ToString()
                '    eu29text.Update()

                '    Panel3.Visible = False
                '    Panel4.Visible = False
                '    Panel5.Visible = False
                '    Panel6.Visible = False
                '    Label105.Visible = False
                '    AHMLtext.Visible = False
                '    HHMLtext.Visible = False
                '    Label106.Visible = False
                '    AVMLtext.Visible = False
                '    HVMLtext.Visible = False
                '    Label133.Visible = False
                '    Label132.Visible = False
                '    Label144.Visible = False
                '    e24text.Visible = False
                '    e25text.Visible = False
                '    e26text.Visible = False
                '    eu24text.Visible = False
                '    eu25text.Visible = False
                '    eu26text.Visible = False
                '    Label107.Visible = False
                '    Label108.Visible = False
                '    Label143.Visible = True
                '    Label146.Visible = True
                '    Label136.Visible = False
                '    e15text.Visible = False
                '    eu15text.Visible = False
                '    e22text.Visible = True
                '    e23text.Visible = True
                '    eu22text.Visible = True
                '    eu23text.Visible = True
                '    TabControl2.SelectedIndex = 4

            End If

        End While

    End Sub

    Private Sub casos_especiales()



    End Sub

    Private Sub cargaExcel()        ' función que carga la base de datos de excel **********************************************************************

        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet

        Dim xlApp2 As Excel.Application
        Dim xlWorkBook2 As Excel.Workbook
        Dim xlWorkSheet2 As Excel.Worksheet

        ' Variable para leer el registro de la base de datos a partir del código y variables para obtener el precio y la descripción *******************

        Dim codleido As Integer
        Dim precio As Double
        Dim descripcion As String
        Dim sIni As String = Application.StartupPath & "\Libroprecios.xls"

        fil = 10                 'fila anterior a la primera donde se inician los artículos 
        col = 7                  'columna de la tabla de datos donde se encuentra el código 

        xlApp = New Excel.Application
        xlWorkBook = xlApp.Workbooks.Open(sIni)
        xlWorkSheet = xlWorkBook.Worksheets("Hoja1")

        xlApp2 = New Excel.Application
        xlWorkBook2 = xlApp2.Workbooks.Open(sIni)
        xlWorkSheet2 = xlWorkBook2.Worksheets("Hoja2")

        ' LLenado de los Text Box con los datos del cliente ********************************************************************************************

        tb_nombre2.Text = tb_nombre.Text
        tb_dni2.Text = tb_dni.Text
        tb_tlf2.Text = tb_tlf.Text
        tb_cif2.Text = tb_cif.Text
        tb_direccion2.Text = tb_direccion.Text
        tb_email2.Text = tb_email.Text

        ' Bucle For con contador para leer la totalidad del registro ***********************************************************************************

        For i As Integer = 0 To 275
            fil = fil + 1
            codleido = xlWorkSheet.Cells(fil, col).value

            ' Bucles If para mostrar en presupuestos las opciones elegidas en los ComboBox de la pantalla inicial. Se añaden en bucles Try para evitar fallos *********

            Try

                If (cmbmodope.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = xlWorkSheet.Cells(fil, 4).value
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod1.Text = codleido.ToString
                    lbldes1.Text = descripcion
                    lbluds1.Text = 1
                    lblpre1.Text = precio
                    importe = xlWorkSheet.Cells(fil, 4).value * 1
                    lblimp1.Text = importe.ToString

                    lblcod1.Visible = True
                    lbldes1.Visible = True
                    lbluds1.Visible = True
                    lblpre1.Visible = True
                    lblimp1.Visible = True

                End If

            Catch ex As Exception

            End Try

            Try

                ' *********************** BUCLE IF PARA CALCULAR LOS COSTOS DE LOS FRENTES DE PUERTA *********************************************

                ' *********************** Carpintería MI con 2HF + 2HM *************************************************************************** 

                If ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 1) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double

                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    precio = ((xlWorkSheet2.Cells(12, 4).value * a2 * au2 * 1.1 / 1000) + (xlWorkSheet2.Cells(12, 4).value * a5 * au5 * 1.1 / 1000)) + ((xlWorkSheet2.Cells(13, 4).value * bver1 * 2 / 1000) + (xlWorkSheet2.Cells(13, 4).value * bver2 * 2 / 1000)) + ((xlWorkSheet2.Cells(14, 4).value * bver1 / 1000) + (xlWorkSheet2.Cells(14, 4).value * bver2 / 1000)) + (xlWorkSheet2.Cells(15, 4).value * a7 * au7 * 1.1 / 1000) + (xlWorkSheet2.Cells(16, 4).value * 4 * 4) + (xlWorkSheet2.Cells(17, 4).value * 2 * hh / 1000) + ((xlWorkSheet2.Cells(18, 4).value * 2 * 2 * (AHF + HHF) / 1000) + (xlWorkSheet2.Cells(18, 4).value * 2 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(19, 4).value * 4 * 4) + (xlWorkSheet2.Cells(20, 4).value * 2 * 2) + (xlWorkSheet2.Cells(21, 4).value * 2 * AHF / 1000) + (xlWorkSheet2.Cells(24, 4).value * 2 * 2) + (xlWorkSheet2.Cells(25, 4).value * 1 * 2) + (xlWorkSheet2.Cells(26, 4).value * a8 / 1000) + (xlWorkSheet2.Cells(27, 4).value * (hh * 4 + a8) / 1000) + (xlWorkSheet2.Cells(27, 4).value * (a8 + 100) / 1000) + (xlWorkSheet2.Cells(28, 4).value * 4) + (xlWorkSheet2.Cells(29, 4).value * 4))
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería MI con 2HM ******************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 2) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double

                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    precio = ((xlWorkSheet2.Cells(12, 4).value * a2 * au2 * 1.1 / 1000) + (xlWorkSheet2.Cells(12, 4).value * a5 * au5 * 1.1 / 1000) + (xlWorkSheet2.Cells(13, 4).value * bver2 * 2 / 1000) + (xlWorkSheet2.Cells(14, 4).value * bver2 / 1000) + (xlWorkSheet2.Cells(15, 4).value * a7 * au7 * 1.1 / 1000) + (xlWorkSheet2.Cells(16, 4).value * 4 * 2) + (xlWorkSheet2.Cells(17, 4).value * 2 * hh / 1000) + (xlWorkSheet2.Cells(18, 4).value * 2 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(19, 4).value * 4 * 2) + (xlWorkSheet2.Cells(24, 4).value * 2 * 2) + (xlWorkSheet2.Cells(25, 4).value * 1 * 2) + (xlWorkSheet2.Cells(26, 4).value * a8 / 1000) + (xlWorkSheet2.Cells(27, 4).value * (hh * 4 + a8) / 1000) + (xlWorkSheet2.Cells(27, 4).value * (a8 + 100) / 1000) + (xlWorkSheet2.Cells(28, 4).value * 2) + (xlWorkSheet2.Cells(29, 4).value * 2))
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería MI con 1HF + 1HM ***************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 3) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double

                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    precio = ((xlWorkSheet2.Cells(12, 4).value * a2 * au2 * 1.1 / 1000) + (xlWorkSheet2.Cells(12, 4).value * a5 * au5 * 1.1 / 1000)) + ((xlWorkSheet2.Cells(13, 4).value * bver1 / 1000) + (xlWorkSheet2.Cells(13, 4).value * bver2 / 1000)) + ((xlWorkSheet2.Cells(14, 4).value * bver1 * 0.5 / 1000) + (xlWorkSheet2.Cells(14, 4).value * bver2 * 0.5 / 1000)) + (xlWorkSheet2.Cells(15, 4).value * a7 * au7 * 1.1 / 1000) + (xlWorkSheet2.Cells(16, 4).value * 4 * 2) + (xlWorkSheet2.Cells(17, 4).value * 1 * hh / 1000) + ((xlWorkSheet2.Cells(18, 4).value * 1 * 2 * (AHF + HHF) / 1000) + (xlWorkSheet2.Cells(18, 4).value * 1 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(19, 4).value * 2 * 4) + (xlWorkSheet2.Cells(20, 4).value * 1 * 2) + (xlWorkSheet2.Cells(21, 4).value * 1 * AHF / 1000) + (xlWorkSheet2.Cells(24, 4).value * 1 * 2) + (xlWorkSheet2.Cells(25, 4).value * 1 * 1) + (xlWorkSheet2.Cells(26, 4).value * a8 / 1000) + (xlWorkSheet2.Cells(27, 4).value * (hh * 4 + a8) / 1000) + (xlWorkSheet2.Cells(27, 4).value * (a8 + 100) / 1000) + (xlWorkSheet2.Cells(28, 4).value * 2) + (xlWorkSheet2.Cells(29, 4).value * 2))
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería MI con 1HM *********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 4) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double

                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    precio = ((xlWorkSheet2.Cells(12, 4).value * a2 * au2 * 1.1 / 1000) + (xlWorkSheet2.Cells(12, 4).value * a5 * au5 * 1.1 / 1000) + (xlWorkSheet2.Cells(13, 4).value * bver2 / 1000) + (xlWorkSheet2.Cells(14, 4).value * bver2 * 0.5 / 1000)) + (xlWorkSheet2.Cells(15, 4).value * a7 * au7 * 1.1 / 1000) + (xlWorkSheet2.Cells(16, 4).value * 4 * 1) + (xlWorkSheet2.Cells(17, 4).value * 1 * hh / 1000) + (xlWorkSheet2.Cells(18, 4).value * 1 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(19, 4).value * 4 * 1) + (xlWorkSheet2.Cells(24, 4).value * 2 * 1) + (xlWorkSheet2.Cells(25, 4).value * 1 * 1) + (xlWorkSheet2.Cells(26, 4).value * a8 / 1000) + (xlWorkSheet2.Cells(27, 4).value * (hh * 4 + a8) / 1000) + (xlWorkSheet2.Cells(27, 4).value * ((a8 + 100) / 1000) + (xlWorkSheet2.Cells(28, 4).value * 1) + (xlWorkSheet2.Cells(29, 4).value * 1))
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería MI TES con 2HF + 4HM *********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 5) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double

                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    precio = ((xlWorkSheet2.Cells(12, 4).value * a2 * au2 * 1.1 / 1000) + (xlWorkSheet2.Cells(12, 4).value * a5 * au5 * 1.1 / 1000) + (xlWorkSheet2.Cells(13, 4).value * bver1 * 2 / 1000) + (xlWorkSheet2.Cells(13, 4).value * bver2 * 4 / 1000) + (xlWorkSheet2.Cells(14, 4).value * bver2 * 2 / 1000) + (xlWorkSheet2.Cells(15, 4).value * a7 * au7 * 1.1 / 1000) + (xlWorkSheet2.Cells(16, 4).value * 4 * 6) + (xlWorkSheet2.Cells(17, 4).value * 2 * hh / 1000) + (xlWorkSheet2.Cells(18, 4).value * 2 * 2 * (AHF + HHF) / 1000) + (xlWorkSheet2.Cells(18, 4).value * 2 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(18, 4).value * 2 * 2 * (AHML + HHML) / 1000) + (xlWorkSheet2.Cells(19, 4).value * 4 * 6) + (xlWorkSheet2.Cells(20, 4).value * 2 * 2) + (xlWorkSheet2.Cells(21, 4).value * 2 * AHF / 1000) + (xlWorkSheet2.Cells(23, 4).value * bver1 / 1000) + (xlWorkSheet2.Cells(23, 4).value * bver2 / 1000) + (xlWorkSheet2.Cells(24, 4).value * 2 * 4) + (xlWorkSheet2.Cells(25, 4).value * 1 * 4) + (xlWorkSheet2.Cells(26, 4).value * a8 / 1000) + (xlWorkSheet2.Cells(27, 4).value * (hh * 4 + a8) / 1000) + (xlWorkSheet2.Cells(27, 4).value * (a8 + 100) / 1000) + (xlWorkSheet2.Cells(28, 4).value * 6) + (xlWorkSheet2.Cells(29, 4).value * 6))
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True


                    ' *********************** Carpintería MI TES con 4HM ************************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 6) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double

                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    precio = ((xlWorkSheet2.Cells(12, 4).value * a2 * au2 * 1.1 / 1000) + (xlWorkSheet2.Cells(12, 4).value * a5 * au5 * 1.1 / 1000) + (xlWorkSheet2.Cells(13, 4).value * bver2 * 4 / 1000) + (xlWorkSheet2.Cells(14, 4).value * bver2 * 2 / 1000) + (xlWorkSheet2.Cells(15, 4).value * a7 * au7 * 1.1 / 1000) + (xlWorkSheet2.Cells(16, 4).value * 4 * 4) + (xlWorkSheet2.Cells(17, 4).value * 2 * hh / 1000) + (xlWorkSheet2.Cells(18, 4).value * 2 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(18, 4).value * 2 * 2 * (AHML + HHML) / 1000) + (xlWorkSheet2.Cells(19, 4).value * 4 * 4) + (xlWorkSheet2.Cells(23, 4).value * bver1 / 1000) + (xlWorkSheet2.Cells(23, 4).value * bver2 / 1000) + (xlWorkSheet2.Cells(24, 4).value * 2 * 4) + (xlWorkSheet2.Cells(25, 4).value * 1 * 4) + (xlWorkSheet2.Cells(26, 4).value * a8 / 1000) + (xlWorkSheet2.Cells(27, 4).value * (hh * 4 + a8) / 1000) + (xlWorkSheet2.Cells(27, 4).value * (a8 + 100) / 1000) + (xlWorkSheet2.Cells(28, 4).value * 4) + (xlWorkSheet2.Cells(29, 4).value * 4))
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería MI TES con 1HF + 2HM ********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 7) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double

                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    precio = ((xlWorkSheet2.Cells(12, 4).value * a2 * au2 * 1.1 / 1000) + (xlWorkSheet2.Cells(12, 4).value * a5 * au5 * 1.1 / 1000) + (xlWorkSheet2.Cells(13, 4).value * bver1 / 1000) + (xlWorkSheet2.Cells(13, 4).value * bver2 * 2 / 1000) + (xlWorkSheet2.Cells(14, 4).value * bver2 * 1 / 1000) + (xlWorkSheet2.Cells(15, 4).value * a7 * au7 * 1.1 / 1000) + (xlWorkSheet2.Cells(16, 4).value * 4 * 3) + (xlWorkSheet2.Cells(17, 4).value * 1 * hh / 1000) + (xlWorkSheet2.Cells(18, 4).value * 1 * 2 * (AHF + HHF) / 1000) + (xlWorkSheet2.Cells(18, 4).value * 1 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(18, 4).value * 1 * 2 * (AHML + HHML) / 1000) + (xlWorkSheet2.Cells(19, 4).value * 4 * 3) + (xlWorkSheet2.Cells(20, 4).value * 2 * 1) + (xlWorkSheet2.Cells(21, 4).value * 1 * AHF / 1000) + (xlWorkSheet2.Cells(23, 4).value * bver1 * 0.5 / 1000) + (xlWorkSheet2.Cells(23, 4).value * bver2 / 1000) + (xlWorkSheet2.Cells(24, 4).value * 2 * 2) + (xlWorkSheet2.Cells(25, 4).value * 1 * 2) + (xlWorkSheet2.Cells(26, 4).value * a8 / 1000) + (xlWorkSheet2.Cells(27, 4).value * (hh * 4 + a8) / 1000) + (xlWorkSheet2.Cells(27, 4).value * (a8 + 100) / 1000) + (xlWorkSheet2.Cells(28, 4).value * 3) + (xlWorkSheet2.Cells(29, 4).value * 3))
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería MI TES con 2HM ************************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 8) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double

                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    precio = ((xlWorkSheet2.Cells(12, 4).value * a2 * au2 * 1.1 / 1000) + (xlWorkSheet2.Cells(12, 4).value * a5 * au5 * 1.1 / 1000) + (xlWorkSheet2.Cells(13, 4).value * bver2 * 2 / 1000) + (xlWorkSheet2.Cells(14, 4).value * bver2 / 1000) + (xlWorkSheet2.Cells(15, 4).value * a7 * au7 * 1.1 / 1000) + (xlWorkSheet2.Cells(16, 4).value * 4 * 2) + (xlWorkSheet2.Cells(17, 4).value * 1 * hh / 1000) + (xlWorkSheet2.Cells(18, 4).value * 1 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(18, 4).value * 1 * 2 * (AHML + HHML) / 1000) + (xlWorkSheet2.Cells(19, 4).value * 4 * 2) + (xlWorkSheet2.Cells(23, 4).value * bver1 * 0.5 / 1000) + (xlWorkSheet2.Cells(23, 4).value * bver2 * 0.5 / 1000) + (xlWorkSheet2.Cells(24, 4).value * 2 * 2) + (xlWorkSheet2.Cells(25, 4).value * 1 * 2) + (xlWorkSheet2.Cells(26, 4).value * a8 / 1000) + (xlWorkSheet2.Cells(27, 4).value * (hh * 4 + a8) / 1000) + (xlWorkSheet2.Cells(27, 4).value * (a8 + 100) / 1000) + (xlWorkSheet2.Cells(28, 4).value * 2) + (xlWorkSheet2.Cells(29, 4).value * 2))
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                End If

            Catch ex As Exception

            End Try

            Try

                ' *********************** BUCLE IF PARA CALCULAR LOS COSTOS DE LOS FRENTES DE PUERTA *********************************************

                ' *********************** Carpintería Full Glass con 2HF + 2HM *************************************************************************** 

                If ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 1) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    precio = (xlWorkSheet2.Cells(35, 4).value * bver1 / 1000) + (xlWorkSheet2.Cells(34, 4).value * bver1 / 1000) + (xlWorkSheet2.Cells(35, 4).value * bver2 / 1000) + (xlWorkSheet2.Cells(34, 4).value * bver2 / 1000) + (xlWorkSheet2.Cells(33, 4).value * b3 * bu3 * 1.1 / 1000) + (xlWorkSheet2.Cells(32, 4).value * b6 * bu6 * 1.1 / 1000) + (xlWorkSheet2.Cells(33, 4).value * b7 * bu7 * 1.1 / 1000) + (xlWorkSheet2.Cells(36, 4).value * wh / 1000) + (xlWorkSheet2.Cells(37, 4).value * 2 * 2 * 2 * (AHF + HHF) / 1000) + (xlWorkSheet2.Cells(37, 4).value * 2 * 2 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(38, 4).value * 2 * hh / 1000) + (xlWorkSheet2.Cells(39, 4).value * (2 * hh + b8) / 1000) + (xlWorkSheet2.Cells(20, 4).value * 2 * 2) + (xlWorkSheet2.Cells(24, 4).value * 2 * 2) + (xlWorkSheet2.Cells(25, 4).value * 1 * 2) + (xlWorkSheet2.Cells(28, 4).value * 4) + (xlWorkSheet2.Cells(29, 4).value * 4)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería Full Glass con 2HM ******************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 2) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    precio = (xlWorkSheet2.Cells(35, 4).value * bver2 / 1000) + (xlWorkSheet2.Cells(34, 4).value * bver2 / 1000) + (xlWorkSheet2.Cells(32, 4).value * b6 * bu6 * 1.1 / 1000) + (xlWorkSheet2.Cells(33, 4).value * b7 * bu7 * 1.1 / 1000) + (xlWorkSheet2.Cells(36, 4).value * wh / 1000) + (xlWorkSheet2.Cells(37, 4).value * 2 * 2 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(38, 4).value * 2 * hh / 1000) + (xlWorkSheet2.Cells(39, 4).value * (2 * hh + b8) / 1000) + (xlWorkSheet2.Cells(24, 4).value * 2 * 2) + (xlWorkSheet2.Cells(25, 4).value * 1 * 2) + (xlWorkSheet2.Cells(28, 4).value * 2) + (xlWorkSheet2.Cells(29, 4).value * 2)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería Full Glass con 1HF + 1HM ***************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 3) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    precio = (xlWorkSheet2.Cells(35, 4).value * bver1 * 0.5 / 1000) + (xlWorkSheet2.Cells(34, 4).value * bver1 * 0.5 / 1000) + (xlWorkSheet2.Cells(35, 4).value * bver2 * 0.5 / 1000) + (xlWorkSheet2.Cells(34, 4).value * bver2 * 0.5 / 1000) + (xlWorkSheet2.Cells(33, 4).value * b3 * bu3 * 1.1 / 1000) + (xlWorkSheet2.Cells(32, 4).value * b6 * bu6 * 1.1 / 1000) + (xlWorkSheet2.Cells(33, 4).value * b7 * bu7 * 1.1 / 1000) + (xlWorkSheet2.Cells(36, 4).value * wh / 1000) + (xlWorkSheet2.Cells(37, 4).value * 1 * 2 * 2 * (AHF + HHF) / 1000) + (xlWorkSheet2.Cells(37, 4).value * 1 * 2 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(38, 4).value * 1 * hh / 1000) + (xlWorkSheet2.Cells(39, 4).value * hh / 1000) + (xlWorkSheet2.Cells(20, 4).value * 1 * 2) + (xlWorkSheet2.Cells(24, 4).value * 1 * 2) + (xlWorkSheet2.Cells(25, 4).value * 1 * 1) + (xlWorkSheet2.Cells(28, 4).value * 2) + (xlWorkSheet2.Cells(29, 4).value * 2)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería Full Glass con 1HM *********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 4) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    precio = (xlWorkSheet2.Cells(35, 4).value * bver2 * 0.5 / 1000) + (xlWorkSheet2.Cells(34, 4).value * bver2 * 0.5 / 1000) + (xlWorkSheet2.Cells(32, 4).value * b6 * bu6 * 1.1 / 1000) + (xlWorkSheet2.Cells(33, 4).value * b7 * bu7 * 1.1 / 1000) + (xlWorkSheet2.Cells(36, 4).value * wh / 1000) + (xlWorkSheet2.Cells(37, 4).value * 1 * 2 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(38, 4).value * 1 * hh / 1000) + (xlWorkSheet2.Cells(39, 4).value * hh / 1000) + (xlWorkSheet2.Cells(24, 4).value * 1 * 2) + (xlWorkSheet2.Cells(25, 4).value * 1 * 1) + (xlWorkSheet2.Cells(28, 4).value * 1) + (xlWorkSheet2.Cells(29, 4).value * 1)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería Full Glass TES con 2HF + 4HM *********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 5) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    precio = (xlWorkSheet2.Cells(35, 4).value * bver1 / 1000) + (xlWorkSheet2.Cells(34, 4).value * bver1 / 1000) + (xlWorkSheet2.Cells(35, 4).value * bver2 * 3 / 1000) + (xlWorkSheet2.Cells(34, 4).value * bver2 / 1000) + (xlWorkSheet2.Cells(33, 4).value * b3 * bu3 * 1.1 / 1000) + (xlWorkSheet2.Cells(32, 4).value * b6 * bu6 * 1.1 / 1000) + (xlWorkSheet2.Cells(33, 4).value * b7 * bu7 * 1.1 / 1000) + (xlWorkSheet2.Cells(36, 4).value * wh / 1000) + (xlWorkSheet2.Cells(37, 4).value * 2 * 2 * 2 * (AHF + HHF) / 1000) + (xlWorkSheet2.Cells(37, 4).value * 2 * 2 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(37, 4).value * 2 * 2 * 2 * (AHML + HHML) / 1000) + (xlWorkSheet2.Cells(38, 4).value * 2 * hh / 1000) + (xlWorkSheet2.Cells(39, 4).value * (2 * hh + b8) / 1000) + (xlWorkSheet2.Cells(20, 4).value * 2 * 2) + (xlWorkSheet2.Cells(24, 4).value * 2 * 4) + (xlWorkSheet2.Cells(25, 4).value * 1 * 4) + (xlWorkSheet2.Cells(28, 4).value * 6) + (xlWorkSheet2.Cells(29, 4).value * 6)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True


                    ' *********************** Carpintería Full Glass TES con 4HM ************************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 6) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    precio = (xlWorkSheet2.Cells(35, 4).value * bver2 * 3 / 1000) + (xlWorkSheet2.Cells(34, 4).value * bver2 / 1000) + (xlWorkSheet2.Cells(32, 4).value * b6 * bu6 * 1.1 / 1000) + (xlWorkSheet2.Cells(33, 4).value * b7 * bu7 * 1.1 / 1000) + (xlWorkSheet2.Cells(36, 4).value * wh / 1000) + (xlWorkSheet2.Cells(37, 4).value * 2 * 2 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(37, 4).value * 2 * 2 * 2 * (AHML + HHML) / 1000) + (xlWorkSheet2.Cells(38, 4).value * 2 * hh / 1000) + (xlWorkSheet2.Cells(39, 4).value * (2 * hh + b8) / 1000) + (xlWorkSheet2.Cells(24, 4).value * 2 * 4) + (xlWorkSheet2.Cells(25, 4).value * 1 * 4) + (xlWorkSheet2.Cells(28, 4).value * 4) + (xlWorkSheet2.Cells(29, 4).value * 4)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería Full Glass TES con 1HF + 2HM ********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 7) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    precio = (xlWorkSheet2.Cells(35, 4).value * bver1 * 0.5 / 1000) + (xlWorkSheet2.Cells(34, 4).value * bver1 * 0.5 / 1000) + (xlWorkSheet2.Cells(35, 4).value * bver2 * 1.5 / 1000) + (xlWorkSheet2.Cells(34, 4).value * bver2 * 0.5 / 1000) + (xlWorkSheet2.Cells(33, 4).value * b3 * bu3 * 1.1 / 1000) + (xlWorkSheet2.Cells(32, 4).value * b6 * bu6 * 1.1 / 1000) + (xlWorkSheet2.Cells(33, 4).value * b7 * bu7 * 1.1 / 1000) + (xlWorkSheet2.Cells(36, 4).value * wh / 1000) + (xlWorkSheet2.Cells(37, 4).value * 1 * 2 * 2 * (AHF + HHF) / 1000) + (xlWorkSheet2.Cells(37, 4).value * 1 * 2 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(37, 4).value * 1 * 2 * 2 * (AHML + HHML) / 1000) + (xlWorkSheet2.Cells(38, 4).value * 1 * hh / 1000) + (xlWorkSheet2.Cells(39, 4).value * hh / 1000) + (xlWorkSheet2.Cells(20, 4).value * 2 * 1) + (xlWorkSheet2.Cells(24, 4).value * 2 * 2) + (xlWorkSheet2.Cells(25, 4).value * 1 * 2) + (xlWorkSheet2.Cells(28, 4).value * 3) + (xlWorkSheet2.Cells(29, 4).value * 3)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería Full Glass TES con 2HM ************************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 8) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    precio = (xlWorkSheet2.Cells(35, 4).value * bver2 * 1.5 / 1000) + (xlWorkSheet2.Cells(34, 4).value * bver2 * 0.5 / 1000) + (xlWorkSheet2.Cells(32, 4).value * b6 * bu6 * 1.1 / 1000) + (xlWorkSheet2.Cells(33, 4).value * b7 * bu7 * 1.1 / 1000) + (xlWorkSheet2.Cells(36, 4).value * wh / 1000) + (xlWorkSheet2.Cells(37, 4).value * 1 * 2 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(37, 4).value * 1 * 2 * 2 * (AHML + HHML) / 1000) + (xlWorkSheet2.Cells(38, 4).value * 1 * hh / 1000) + (xlWorkSheet2.Cells(39, 4).value * hh / 1000) + (xlWorkSheet2.Cells(24, 4).value * 2 * 2) + (xlWorkSheet2.Cells(25, 4).value * 1 * 2) + (xlWorkSheet2.Cells(28, 4).value * 2) + (xlWorkSheet2.Cells(29, 4).value * 2)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                End If

            Catch ex As Exception

            End Try

            Try

                ' *********************** BUCLE IF PARA CALCULAR LOS COSTOS DE LOS FRENTES DE PUERTA *********************************************

                ' *********************** Carpintería Plintos silicona superior e inferior con 2HF + 2HM *************************************************************************** 

                If (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 1 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(47, 4).value * c3 * cu3 / 1000) + (xlWorkSheet2.Cells(46, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 4 * hh / 1000) + (xlWorkSheet2.Cells(20, 4).value * 2 * 2) + (xlWorkSheet2.Cells(25, 4).value * 1 * 2) + (xlWorkSheet2.Cells(28, 4).value * 1 * 4) + (xlWorkSheet2.Cells(29, 4).value * 1 * 4)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería Plintos silicona superior e inferior con 2HM ******************************************************************************

                ElseIf (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 2 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(47, 4).value * c3 * cu3 / 1000) + (xlWorkSheet2.Cells(46, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 2 * hh / 1000) + (xlWorkSheet2.Cells(25, 4).value * 1 * 2) + (xlWorkSheet2.Cells(28, 4).value * 1 * 2) + (xlWorkSheet2.Cells(29, 4).value * 1 * 2)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería Plintos silicona superior e inferior con 1HF + 1HM ***************************************************************************

                ElseIf (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 3 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(47, 4).value * c3 * cu3 / 1000) + (xlWorkSheet2.Cells(46, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 2 * hh / 1000) + (xlWorkSheet2.Cells(20, 4).value * 2 * 1) + (xlWorkSheet2.Cells(25, 4).value * 1 * 1) + (xlWorkSheet2.Cells(28, 4).value * 1 * 2) + (xlWorkSheet2.Cells(29, 4).value * 1 * 2)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería Plintos silicona superior e inferior con 1HM ******************************************

                ElseIf (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 4 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(47, 4).value * c3 * cu3 / 1000) + (xlWorkSheet2.Cells(46, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 1 * hh / 1000) + (xlWorkSheet2.Cells(25, 4).value * 1 * 1) + (xlWorkSheet2.Cells(28, 4).value * 1 * 1) + (xlWorkSheet2.Cells(29, 4).value * 1 * 1)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería Plintos silicona superior e inferior TES con 2HF + 4HM *********************************************************************************

                ElseIf (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 5 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(47, 4).value * c3 * cu3 / 1000) + (xlWorkSheet2.Cells(46, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 6 * hh / 1000) + (xlWorkSheet2.Cells(20, 4).value * 2 * 2) + (xlWorkSheet2.Cells(25, 4).value * 1 * 4) + (xlWorkSheet2.Cells(28, 4).value * 1 * 6) + (xlWorkSheet2.Cells(29, 4).value * 1 * 6)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True


                    ' *********************** Carpintería Plintos silicona superior e inferior TES con 4HM **************************************************

                ElseIf (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 6 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(47, 4).value * c3 * cu3 / 1000) + (xlWorkSheet2.Cells(46, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 4 * hh / 1000) + (xlWorkSheet2.Cells(25, 4).value * 1 * 4) + (xlWorkSheet2.Cells(28, 4).value * 1 * 4) + (xlWorkSheet2.Cells(29, 4).value * 1 * 4)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería Plintos silicona superior e inferior TES con 1HF + 2HM *****************************************

                ElseIf (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 7 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(47, 4).value * c3 * cu3 / 1000) + (xlWorkSheet2.Cells(46, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 3 * hh / 1000) + (xlWorkSheet2.Cells(20, 4).value * 2 * 2) + (xlWorkSheet2.Cells(25, 4).value * 1 * 2) + (xlWorkSheet2.Cells(28, 4).value * 1 * 3) + (xlWorkSheet2.Cells(29, 4).value * 1 * 3)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería Plintos silicona superior e inferior TES con 2HM *****************

                ElseIf (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 8 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(47, 4).value * c3 * cu3 / 1000) + (xlWorkSheet2.Cells(46, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 2 * hh / 1000) + (xlWorkSheet2.Cells(25, 4).value * 1 * 2) + (xlWorkSheet2.Cells(28, 4).value * 1 * 2) + (xlWorkSheet2.Cells(29, 4).value * 1 * 2)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                End If

            Catch ex As Exception

            End Try

            Try

                ' *********************** BUCLE IF PARA CALCULAR LOS COSTOS DE LOS FRENTES DE PUERTA ********************************

                ' *********************** Carpintería Plintos silicona solo superior  con 2HF + 2HM **************************************************** 

                If (cmbmodcar.SelectedIndex = 4 And cmbconf.SelectedIndex = 1 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(46, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 4 * hh / 1000) + (xlWorkSheet2.Cells(20, 4).value * 2 * 2) + (xlWorkSheet2.Cells(25, 4).value * 1 * 2) + (xlWorkSheet2.Cells(28, 4).value * 1 * 4) + (xlWorkSheet2.Cells(29, 4).value * 1 * 4)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería Plintos silicona solo superior con 2HM **************************************************

                ElseIf (cmbmodcar.SelectedIndex = 4 And cmbconf.SelectedIndex = 2 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(46, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 2 * hh / 1000) + (xlWorkSheet2.Cells(25, 4).value * 1 * 2) + (xlWorkSheet2.Cells(28, 4).value * 1 * 2) + (xlWorkSheet2.Cells(29, 4).value * 1 * 2)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería Plintos silicona solo superior con 1HF + 1HM **************************************************

                ElseIf (cmbmodcar.SelectedIndex = 4 And cmbconf.SelectedIndex = 3 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(46, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 2 * hh / 1000) + (xlWorkSheet2.Cells(20, 4).value * 2 * 1) + (xlWorkSheet2.Cells(25, 4).value * 1 * 1) + (xlWorkSheet2.Cells(28, 4).value * 1 * 2) + (xlWorkSheet2.Cells(29, 4).value * 1 * 2)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería Plintos silicona solo superior con 1HM ******************************************

                ElseIf (cmbmodcar.SelectedIndex = 4 And cmbconf.SelectedIndex = 4 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(46, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 1 * hh / 1000) + (xlWorkSheet2.Cells(25, 4).value * 1 * 1) + (xlWorkSheet2.Cells(28, 4).value * 1 * 1) + (xlWorkSheet2.Cells(29, 4).value * 1 * 1)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                End If

            Catch ex As Exception

            End Try

            Try

                ' *********************** BUCLE IF PARA CALCULAR LOS COSTOS DE LOS FRENTES DE PUERTA *********************************************

                ' *********************** Carpintería Plintos superior de pinza e inferior de silicona con 2HF + 2HM *************************************************************************** 

                If (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 1 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(47, 4).value * c3 * cu3 / 1000) + (xlWorkSheet2.Cells(48, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 4 * hh / 1000) + (xlWorkSheet2.Cells(20, 4).value * 2 * 2) + (xlWorkSheet2.Cells(25, 4).value * 1 * 2) + (xlWorkSheet2.Cells(28, 4).value * 1 * 4) + (xlWorkSheet2.Cells(29, 4).value * 1 * 4)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería Plintos superior de pinza e inferior de silicona con 2HM ******************************************************************************

                ElseIf (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 2 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(47, 4).value * c3 * cu3 / 1000) + (xlWorkSheet2.Cells(48, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 2 * hh / 1000) + (xlWorkSheet2.Cells(25, 4).value * 1 * 2) + (xlWorkSheet2.Cells(28, 4).value * 1 * 2) + (xlWorkSheet2.Cells(29, 4).value * 1 * 2)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería Plintos superior de pinza e inferior de silicona con 1HF + 1HM ***********************************************

                ElseIf (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 3 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(47, 4).value * c3 * cu3 / 1000) + (xlWorkSheet2.Cells(48, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 2 * hh / 1000) + (xlWorkSheet2.Cells(20, 4).value * 2 * 1) + (xlWorkSheet2.Cells(25, 4).value * 1 * 1) + (xlWorkSheet2.Cells(28, 4).value * 1 * 2) + (xlWorkSheet2.Cells(29, 4).value * 1 * 2)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería Plintos superior de pinza e inferior de silicona con 1HM ******************************************

                ElseIf (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 4 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(47, 4).value * c3 * cu3 / 1000) + (xlWorkSheet2.Cells(48, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 1 * hh / 1000) + (xlWorkSheet2.Cells(25, 4).value * 1 * 1) + (xlWorkSheet2.Cells(28, 4).value * 1 * 1) + (xlWorkSheet2.Cells(29, 4).value * 1 * 1)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería Plintos superior de pinza e inferior de silicona TES con 2HF + 4HM **************************************

                ElseIf (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 5 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(47, 4).value * c3 * cu3 / 1000) + (xlWorkSheet2.Cells(48, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 6 * hh / 1000) + (xlWorkSheet2.Cells(20, 4).value * 2 * 2) + (xlWorkSheet2.Cells(25, 4).value * 1 * 4) + (xlWorkSheet2.Cells(28, 4).value * 1 * 6) + (xlWorkSheet2.Cells(29, 4).value * 1 * 6)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True


                    ' *********************** Carpintería Plintos superior de pinza e inferior de silicona TES con 4HM ************************************

                ElseIf (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 6 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(47, 4).value * c3 * cu3 / 1000) + (xlWorkSheet2.Cells(48, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 4 * hh / 1000) + (xlWorkSheet2.Cells(25, 4).value * 1 * 4) + (xlWorkSheet2.Cells(28, 4).value * 1 * 4) + (xlWorkSheet2.Cells(29, 4).value * 1 * 4)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' Carpintería Plintos superior de pinza e inferior de silicona TES con 1HF + 2HM *****************************************

                ElseIf (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 7 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(47, 4).value * c3 * cu3 / 1000) + (xlWorkSheet2.Cells(48, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 3 * hh / 1000) + (xlWorkSheet2.Cells(20, 4).value * 2 * 2) + (xlWorkSheet2.Cells(25, 4).value * 1 * 2) + (xlWorkSheet2.Cells(28, 4).value * 1 * 3) + (xlWorkSheet2.Cells(29, 4).value * 1 * 3)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería Plintos superior de pinza e inferior de silicona TES con 2HM *****************

                ElseIf (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 8 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(47, 4).value * c3 * cu3 / 1000) + (xlWorkSheet2.Cells(48, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 2 * hh / 1000) + (xlWorkSheet2.Cells(25, 4).value * 1 * 2) + (xlWorkSheet2.Cells(28, 4).value * 1 * 2) + (xlWorkSheet2.Cells(29, 4).value * 1 * 2)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                End If

            Catch ex As Exception

            End Try

            Try

                ' *********************** BUCLE IF PARA CALCULAR LOS COSTOS DE LOS FRENTES DE PUERTA ********************************

                ' *********************** Carpintería Plintos de pinza solo superior  con 2HF + 2HM **************************************************** 

                If (cmbmodcar.SelectedIndex = 6 And cmbconf.SelectedIndex = 1 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(48, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 4 * hh / 1000) + (xlWorkSheet2.Cells(20, 4).value * 2 * 2) + (xlWorkSheet2.Cells(25, 4).value * 1 * 2) + (xlWorkSheet2.Cells(28, 4).value * 1 * 4) + (xlWorkSheet2.Cells(29, 4).value * 1 * 4)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería Plintos de pinza solo superior con 2HM **************************************************

                ElseIf (cmbmodcar.SelectedIndex = 6 And cmbconf.SelectedIndex = 2 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(48, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 2 * hh / 1000) + (xlWorkSheet2.Cells(25, 4).value * 1 * 2) + (xlWorkSheet2.Cells(28, 4).value * 1 * 2) + (xlWorkSheet2.Cells(29, 4).value * 1 * 2)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería Plintos de pinza solo superior con 1HF + 1HM **************************************************

                ElseIf (cmbmodcar.SelectedIndex = 6 And cmbconf.SelectedIndex = 3 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(48, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 2 * hh / 1000) + (xlWorkSheet2.Cells(20, 4).value * 2 * 1) + (xlWorkSheet2.Cells(25, 4).value * 1 * 1) + (xlWorkSheet2.Cells(28, 4).value * 1 * 2) + (xlWorkSheet2.Cells(29, 4).value * 1 * 2)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería Plintos de pinza solo superior con 1HM ******************************************

                ElseIf (cmbmodcar.SelectedIndex = 6 And cmbconf.SelectedIndex = 4 And cmbmodcar.SelectedValue = codleido.ToString) Then

                    Dim importe As Double

                    precio = (xlWorkSheet2.Cells(47, 4).value * c1 * cu1 * 1.1 / 1000) + (xlWorkSheet2.Cells(48, 4).value * c2 * cu2 / 1000) + (xlWorkSheet2.Cells(60, 4).value * 1 * hh / 1000) + (xlWorkSheet2.Cells(25, 4).value * 1 * 1) + (xlWorkSheet2.Cells(28, 4).value * 1 * 1) + (xlWorkSheet2.Cells(29, 4).value * 1 * 1)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                End If

            Catch ex As Exception

            End Try

            Try

                ' *********************** BUCLE IF PARA CALCULAR LOS COSTOS DE LOS FRENTES DE PUERTA *********************************************

                ' *********************** Carpintería BW52 con 2HF + 2HM *************************************************************************** 

                If ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 1) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double

                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 4700
                    End If

                    precio = (xlWorkSheet2.Cells(67, 4).value * d4 * du4 * 1.1 / 1000) + (xlWorkSheet2.Cells(67, 4).value * bver1 * 2 / 1000) + (xlWorkSheet2.Cells(68, 4).value * d5 * du5 * 1.1 / 1000) + (xlWorkSheet2.Cells(75, 4).value * bver3 * 2 / 1000) + (xlWorkSheet2.Cells(75, 4).value * d7 * du7 * 1.1 / 1000) + (xlWorkSheet2.Cells(69, 4).value * bver1 / 1000) + (xlWorkSheet2.Cells(70, 4).value * bver1 / 1000) + (xlWorkSheet2.Cells(73, 4).value * d8 * du8 * 1.1 / 1000) + (xlWorkSheet2.Cells(72, 4).value * d20 * du20 * 1.1 / 1000) + (xlWorkSheet2.Cells(79, 4).value * 4 * 4) + (xlWorkSheet2.Cells(81, 4).value * 2 * 2) + (xlWorkSheet2.Cells(76, 4).value * 2 * 2 * (AHF + HHF) / 1000) + (xlWorkSheet2.Cells(76, 4).value * 2 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(67, 4).value * d13 * du13 * 1.1 / 1000) + (xlWorkSheet2.Cells(67, 4).value * bver2 * 2 / 1000) + (xlWorkSheet2.Cells(68, 4).value * d14 * du14 * 1.1 / 1000) + (xlWorkSheet2.Cells(75, 4).value * bver4 * 2 / 1000) + (xlWorkSheet2.Cells(75, 4).value * d17 * du17 * 1.1 / 1000) + (xlWorkSheet2.Cells(74, 4).value * d18 * du18 * 1.1 / 1000) + (xlWorkSheet2.Cells(78, 4).value * d8 * du8 * 1.1 / 1000) + (xlWorkSheet2.Cells(69, 4).value * bver2 / 1000) + (xlWorkSheet2.Cells(71, 4).value * bver2 / 1000) + (xlWorkSheet2.Cells(77, 4).value * 2 * hh / 1000) + (xlWorkSheet2.Cells(70, 4).value * bver2 / 1000) + (xlWorkSheet2.Cells(83, 4).value * 2 * 2) + (xlWorkSheet2.Cells(28, 4).value * 1 * 4) + (xlWorkSheet2.Cells(29, 4).value * 1 * 4) + (xlWorkSheet2.Cells(84, 4).value * 1 * 2)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería BW52 con 2HM ******************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 2) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double

                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 4700
                    End If

                    precio = (xlWorkSheet2.Cells(79, 4).value * 4 * 2) + (xlWorkSheet2.Cells(76, 4).value * 2 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(67, 4).value * d13 * du13 * 1.1 / 1000) + (xlWorkSheet2.Cells(67, 4).value * bver2 * 2 / 1000) + (xlWorkSheet2.Cells(68, 4).value * d14 * du14 * 1.1 / 1000) + (xlWorkSheet2.Cells(75, 4).value * bver4 * 2 / 1000) + (xlWorkSheet2.Cells(75, 4).value * d17 * du17 * 1.1 / 1000) + (xlWorkSheet2.Cells(74, 4).value * d18 * du18 * 1.1 / 1000) + (xlWorkSheet2.Cells(78, 4).value * d8 * du8 * 1.1 / 1000) + (xlWorkSheet2.Cells(69, 4).value * bver2 / 1000) + (xlWorkSheet2.Cells(71, 4).value * bver2 / 1000) + (xlWorkSheet2.Cells(77, 4).value * 2 * hh / 1000) + (xlWorkSheet2.Cells(70, 4).value * bver2 / 1000) + (xlWorkSheet2.Cells(83, 4).value * 2 * 2) + (xlWorkSheet2.Cells(28, 4).value * 1 * 2) + (xlWorkSheet2.Cells(29, 4).value * 1 * 2) + (xlWorkSheet2.Cells(84, 4).value * 1 * 2)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería BW52 con 1HF + 1HM ***************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 3) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double

                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 4700
                    End If

                    precio = (xlWorkSheet2.Cells(67, 4).value * d4 * du4 * 1.1 / 1000) + (xlWorkSheet2.Cells(67, 4).value * bver1 * 1 / 1000) + (xlWorkSheet2.Cells(68, 4).value * d5 * du5 * 1.1 / 1000) + (xlWorkSheet2.Cells(75, 4).value * bver3 * 1 / 1000) + (xlWorkSheet2.Cells(75, 4).value * d7 * du7 * 1.1 / 1000) + (xlWorkSheet2.Cells(69, 4).value * bver1 * 0.5 / 1000) + (xlWorkSheet2.Cells(70, 4).value * bver1 * 0.5 / 1000) + (xlWorkSheet2.Cells(73, 4).value * d8 * du8 * 1.1 / 1000) + (xlWorkSheet2.Cells(72, 4).value * d20 * du20 * 1.1 / 1000) + (xlWorkSheet2.Cells(79, 4).value * 4 * 2) + (xlWorkSheet2.Cells(81, 4).value * 2 * 1) + (xlWorkSheet2.Cells(76, 4).value * 1 * 2 * (AHF + HHF) / 1000) + (xlWorkSheet2.Cells(76, 4).value * 1 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(67, 4).value * d13 * du13 * 1.1 / 1000) + (xlWorkSheet2.Cells(67, 4).value * bver2 * 1 / 1000) + (xlWorkSheet2.Cells(68, 4).value * d14 * du14 * 1.1 / 1000) + (xlWorkSheet2.Cells(75, 4).value * bver4 * 1 / 1000) + (xlWorkSheet2.Cells(75, 4).value * d17 * du17 * 1.1 / 1000) + (xlWorkSheet2.Cells(74, 4).value * d18 * du18 * 1.1 / 1000) + (xlWorkSheet2.Cells(78, 4).value * d8 * du8 * 1.1 / 1000) + (xlWorkSheet2.Cells(69, 4).value * bver2 * 0.5 / 1000) + (xlWorkSheet2.Cells(71, 4).value * bver2 * 0.5 / 1000) + (xlWorkSheet2.Cells(77, 4).value * 1 * hh / 1000) + (xlWorkSheet2.Cells(70, 4).value * bver2 * 0.5 / 1000) + (xlWorkSheet2.Cells(83, 4).value * 2 * 1) + (xlWorkSheet2.Cells(28, 4).value * 1 * 2) + (xlWorkSheet2.Cells(29, 4).value * 1 * 2) + (xlWorkSheet2.Cells(84, 4).value * 1 * 1)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería BW52 con 1HM *********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 4) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double

                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 4700
                    End If

                    precio = (xlWorkSheet2.Cells(79, 4).value * 4 * 1) + (xlWorkSheet2.Cells(76, 4).value * 1 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(67, 4).value * d13 * du13 * 1.1 / 1000) + (xlWorkSheet2.Cells(67, 4).value * bver2 / 1000) + (xlWorkSheet2.Cells(68, 4).value * d14 * du14 * 1.1 / 1000) + (xlWorkSheet2.Cells(75, 4).value * bver4 / 1000) + (xlWorkSheet2.Cells(75, 4).value * d17 * du17 * 1.1 / 1000) + (xlWorkSheet2.Cells(74, 4).value * d18 * du18 * 1.1 / 1000) + (xlWorkSheet2.Cells(78, 4).value * d8 * du8 * 1.1 / 1000) + (xlWorkSheet2.Cells(69, 4).value * bver2 * 0.5 / 1000) + (xlWorkSheet2.Cells(71, 4).value * bver2 * 0.5 / 1000) + (xlWorkSheet2.Cells(77, 4).value * 1 * hh / 1000) + (xlWorkSheet2.Cells(70, 4).value * bver2 * 0.5 / 1000) + (xlWorkSheet2.Cells(83, 4).value * 2 * 1) + (xlWorkSheet2.Cells(28, 4).value * 1 * 1) + (xlWorkSheet2.Cells(29, 4).value * 1 * 1) + (xlWorkSheet2.Cells(84, 4).value * 1 * 1)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True


                    ' *********************** Carpintería BW52 TES con 2HF + 4HM *******************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 5) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double

                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 4700
                    End If

                    precio = (xlWorkSheet2.Cells(67, 4).value * d4 * du4 * 1.1 / 1000) + (xlWorkSheet2.Cells(67, 4).value * bver1 * 2 / 1000) + (xlWorkSheet2.Cells(68, 4).value * d5 * du5 * 1.1 / 1000) + (xlWorkSheet2.Cells(75, 4).value * bver3 * 2 / 1000) + (xlWorkSheet2.Cells(75, 4).value * d7 * du7 * 1.1 / 1000) + (xlWorkSheet2.Cells(69, 4).value * bver1 / 1000) + (xlWorkSheet2.Cells(70, 4).value * bver1 / 1000) + (xlWorkSheet2.Cells(73, 4).value * d8 * du8 * 1.1 / 1000) + (xlWorkSheet2.Cells(72, 4).value * d20 * du20 * 1.1 / 1000) + (xlWorkSheet2.Cells(79, 4).value * 4 * 6) + (xlWorkSheet2.Cells(81, 4).value * 2 * 2) + (xlWorkSheet2.Cells(76, 4).value * 2 * 2 * (AHF + HHF) / 1000) + (xlWorkSheet2.Cells(76, 4).value * 2 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(76, 4).value * 2 * 2 * (AHML + HHML) / 1000) + (xlWorkSheet2.Cells(67, 4).value * d13 * du13 * 1.1 / 1000) + (xlWorkSheet2.Cells(67, 4).value * bver2 * 4 / 1000) + (xlWorkSheet2.Cells(68, 4).value * d14 * du14 * 1.1 / 1000) + (xlWorkSheet2.Cells(75, 4).value * bver4 * 4 / 1000) + (xlWorkSheet2.Cells(75, 4).value * d17 * du17 * 1.1 / 1000) + (xlWorkSheet2.Cells(74, 4).value * d18 * du18 * 1.1 / 1000) + (xlWorkSheet2.Cells(78, 4).value * d8 * du8 * 1.1 / 1000) + (xlWorkSheet2.Cells(69, 4).value * bver2 * 3 / 1000) + (xlWorkSheet2.Cells(71, 4).value * bver2 / 1000) + (xlWorkSheet2.Cells(77, 4).value * 2 * hh / 1000) + (xlWorkSheet2.Cells(70, 4).value * bver2 * 2 / 1000) + (xlWorkSheet2.Cells(83, 4).value * 2 * 4) + (xlWorkSheet2.Cells(28, 4).value * 1 * 6) + (xlWorkSheet2.Cells(29, 4).value * 1 * 6) + (xlWorkSheet2.Cells(84, 4).value * 1 * 4)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True


                    ' *********************** Carpintería BW52 TES con 4HM ************************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 6) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double

                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 4700
                    End If

                    precio = (xlWorkSheet2.Cells(79, 4).value * 4 * 4) + (xlWorkSheet2.Cells(76, 4).value * 2 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(76, 4).value * 2 * 2 * (AHML + HHML) / 1000) + (xlWorkSheet2.Cells(67, 4).value * d13 * du13 * 1.1 / 1000) + (xlWorkSheet2.Cells(67, 4).value * bver2 * 4 / 1000) + (xlWorkSheet2.Cells(68, 4).value * d14 * du14 * 1.1 / 1000) + (xlWorkSheet2.Cells(75, 4).value * bver4 * 4 / 1000) + (xlWorkSheet2.Cells(75, 4).value * d17 * du17 * 1.1 / 1000) + (xlWorkSheet2.Cells(74, 4).value * d18 * du18 * 1.1 / 1000) + (xlWorkSheet2.Cells(78, 4).value * d8 * du8 * 1.1 / 1000) + (xlWorkSheet2.Cells(69, 4).value * bver2 * 2 / 1000) + (xlWorkSheet2.Cells(71, 4).value * bver2 / 1000) + (xlWorkSheet2.Cells(77, 4).value * 2 * hh / 1000) + (xlWorkSheet2.Cells(70, 4).value * bver2 * 2 / 1000) + (xlWorkSheet2.Cells(83, 4).value * 2 * 4) + (xlWorkSheet2.Cells(28, 4).value * 1 * 4) + (xlWorkSheet2.Cells(29, 4).value * 1 * 4) + (xlWorkSheet2.Cells(84, 4).value * 1 * 4)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería BW52 TES con 1HF + 2HM ********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 7) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double

                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 4700
                    End If

                    precio = (xlWorkSheet2.Cells(67, 4).value * d4 * du4 * 1.1 / 1000) + (xlWorkSheet2.Cells(67, 4).value * bver1 / 1000) + (xlWorkSheet2.Cells(68, 4).value * d5 * du5 * 1.1 / 1000) + (xlWorkSheet2.Cells(75, 4).value * bver3 / 1000) + (xlWorkSheet2.Cells(75, 4).value * d7 * du7 * 1.1 / 1000) + (xlWorkSheet2.Cells(69, 4).value * bver1 * 0.5 / 1000) + (xlWorkSheet2.Cells(70, 4).value * bver1 * 0.5 / 1000) + (xlWorkSheet2.Cells(73, 4).value * d8 * du8 * 1.1 / 1000) + (xlWorkSheet2.Cells(72, 4).value * d20 * du20 * 1.1 / 1000) + (xlWorkSheet2.Cells(79, 4).value * 4 * 3) + (xlWorkSheet2.Cells(81, 4).value * 2 * 1) + (xlWorkSheet2.Cells(76, 4).value * 1 * 2 * (AHF + HHF) / 1000) + (xlWorkSheet2.Cells(76, 4).value * 1 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(76, 4).value * 1 * 2 * (AHML + HHML) / 1000) + (xlWorkSheet2.Cells(67, 4).value * d13 * du13 * 1.1 / 1000) + (xlWorkSheet2.Cells(67, 4).value * bver2 * 2 / 1000) + (xlWorkSheet2.Cells(68, 4).value * d14 * du14 * 1.1 / 1000) + (xlWorkSheet2.Cells(75, 4).value * bver4 * 2 / 1000) + (xlWorkSheet2.Cells(75, 4).value * d17 * du17 * 1.1 / 1000) + (xlWorkSheet2.Cells(74, 4).value * d18 * du18 * 1.1 / 1000) + (xlWorkSheet2.Cells(78, 4).value * d8 * du8 * 1.1 / 1000) + (xlWorkSheet2.Cells(69, 4).value * bver2 * 1.5 / 1000) + (xlWorkSheet2.Cells(71, 4).value * bver2 * 0.5 / 1000) + (xlWorkSheet2.Cells(77, 4).value * 1 * hh / 1000) + (xlWorkSheet2.Cells(70, 4).value * bver2 * 1 / 1000) + (xlWorkSheet2.Cells(83, 4).value * 2 * 2) + (xlWorkSheet2.Cells(28, 4).value * 1 * 3) + (xlWorkSheet2.Cells(29, 4).value * 1 * 3) + (xlWorkSheet2.Cells(84, 4).value * 1 * 3)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                    ' *********************** Carpintería BW52 TES con 2HM ************************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 8) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double

                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 4700
                    End If

                    precio = (xlWorkSheet2.Cells(79, 4).value * 4 * 2) + (xlWorkSheet2.Cells(76, 4).value * 1 * 2 * (AHM + HHM) / 1000) + (xlWorkSheet2.Cells(76, 4).value * 1 * 2 * (AHML + HHML) / 1000) + (xlWorkSheet2.Cells(67, 4).value * d13 * du13 * 1.1 / 1000) + (xlWorkSheet2.Cells(67, 4).value * bver2 * 2 / 1000) + (xlWorkSheet2.Cells(68, 4).value * d14 * du14 * 1.1 / 1000) + (xlWorkSheet2.Cells(75, 4).value * bver4 * 2 / 1000) + (xlWorkSheet2.Cells(75, 4).value * d17 * du17 * 1.1 / 1000) + (xlWorkSheet2.Cells(74, 4).value * d18 * du18 * 1.1 / 1000) + (xlWorkSheet2.Cells(78, 4).value * d8 * du8 * 1.1 / 1000) + (xlWorkSheet2.Cells(69, 4).value * bver2 * 1.5 / 1000) + (xlWorkSheet2.Cells(71, 4).value * bver2 * 0.5 / 1000) + (xlWorkSheet2.Cells(77, 4).value * 1 * hh / 1000) + (xlWorkSheet2.Cells(70, 4).value * bver2 * 1 / 1000) + (xlWorkSheet2.Cells(83, 4).value * 2 * 2) + (xlWorkSheet2.Cells(28, 4).value * 1 * 2) + (xlWorkSheet2.Cells(29, 4).value * 1 * 2) + (xlWorkSheet2.Cells(84, 4).value * 1 * 2)
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod2.Text = codleido.ToString
                    lbldes2.Text = descripcion
                    lbluds2.Text = 1
                    lblpre2.Text = precio
                    importe = precio * 1
                    lblimp2.Text = importe.ToString

                    lblcod2.Visible = True
                    lbldes2.Visible = True
                    lbluds2.Visible = True
                    lblpre2.Visible = True
                    lblimp2.Visible = True

                End If

            Catch ex As Exception

            End Try

            ' ******** TRATAMIENTO DE LA PERFILERÍA ***********************************************
            ' **** Condicional para asignar a cada grupo de RAL su tarifa correspondiente *********

            Try

                If (cmblac.SelectedIndex = 1) Then
                    pg = 1.781
                ElseIf (cmblac.SelectedIndex = 2) Then
                    pg = 1.893
                ElseIf (cmblac.SelectedIndex = 3) Then
                    pg = 2.319
                ElseIf (cmblac.SelectedIndex = 4) Then
                    pg = 3.313
                End If

            Catch ex As Exception

            End Try

            ' *********************** BUCLE IF PARA CALCULAR LOS COSTOS DE LOS LACADOS *********************************************

            ' *********************** Carpintería MI con 2HF + 2HM *************************************************************************** 
            Try

                If ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 1) And (cmblac.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double

                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (a2 * au2 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (a5 * au5 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If
                    If (a7 * au7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (a8 * au8 < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (2 * AHF < 4700) Then
                        cr5 = 1.2
                    Else : cr5 = 1.0
                    End If

                    If (a10 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(12, 10).value * a2 * au2 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(12, 10).value * a5 * au5 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(13, 10).value * bver1 * 2 / 1000) + (pg * xlWorkSheet2.Cells(13, 10).value * bver2 * 2 / 1000) + (pg * xlWorkSheet2.Cells(14, 10).value * bver1 / 1000) + (pg * xlWorkSheet2.Cells(14, 10).value * bver2 / 1000) + (pg * xlWorkSheet2.Cells(15, 10).value * a7 * au7 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(26, 10).value * a8 * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(21, 10).value * 2 * AHF * 1.25 * cr5 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * a10 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblcod3.Text = codleido.ToString
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblcod3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería MI con 2HM ******************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 2) And (cmblac.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double


                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (a2 * au2 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (a5 * au5 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If
                    If (a7 * au7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (a8 * au8 < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (2 * AHF < 4700) Then
                        cr5 = 1.2
                    Else : cr5 = 1.0
                    End If

                    If (a10 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(12, 10).value * a2 * au2 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(12, 10).value * a5 * au5 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(13, 10).value * bver2 * 2 / 1000) + (pg * xlWorkSheet2.Cells(14, 10).value * bver2 / 1000) + (pg * xlWorkSheet2.Cells(15, 10).value * a7 * au7 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(26, 10).value * a8 * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * a10 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblcod3.Text = codleido.ToString
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblcod3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería MI con 1HF + 1HM ***************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 3) And (cmblac.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double

                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (a2 * au2 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (a5 * au5 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If
                    If (a7 * au7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (a8 * au8 < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (AHF < 4700) Then
                        cr5 = 1.2
                    Else : cr5 = 1.0
                    End If

                    If (a10 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(12, 10).value * a2 * au2 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(12, 10).value * a5 * au5 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(13, 10).value * bver1 / 1000) + (pg * xlWorkSheet2.Cells(13, 10).value * bver2 / 1000) + (pg * xlWorkSheet2.Cells(14, 10).value * bver1 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(14, 10).value * bver2 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(15, 10).value * a7 * au7 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(21, 10).value * 1 * AHF * 1.25 * cr5 / 1000) + (pg * xlWorkSheet2.Cells(26, 10).value * a8 * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * a10 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblcod3.Text = codleido.ToString
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblcod3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería MI con 1HM *********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 4) And (cmblac.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double


                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (a2 * au2 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (a5 * au5 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If
                    If (a7 * au7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (a8 * au8 < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (AHF < 4700) Then
                        cr5 = 1.2
                    Else : cr5 = 1.0
                    End If

                    If (a10 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(12, 10).value * a2 * au2 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(12, 10).value * a5 * au5 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(13, 10).value * bver2 / 1000) + (pg * xlWorkSheet2.Cells(14, 10).value * bver2 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(15, 10).value * a7 * au7 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(26, 4).value * a8 * 1.25 * cr5 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * a10 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblcod3.Text = codleido.ToString
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblcod3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería MI TES con 2HF + 4HM *********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 5) And (cmblac.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double


                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (a2 * au2 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (a5 * au5 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If
                    If (a7 * au7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (a8 * au8 < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (2 * AHF < 4700) Then
                        cr5 = 1.2
                    Else : cr5 = 1.0
                    End If

                    If (a10 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(12, 10).value * a2 * au2 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(12, 10).value * a5 * au5 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(13, 10).value * bver1 * 2 / 1000) + (pg * xlWorkSheet2.Cells(13, 10).value * bver2 * 4 / 1000) + (pg * xlWorkSheet2.Cells(14, 10).value * bver2 * 2 / 1000) + (pg * xlWorkSheet2.Cells(15, 10).value * a7 * au7 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(21, 10).value * 2 * AHF * 1.25 * cr5 / 1000) + (pg * xlWorkSheet2.Cells(23, 10).value * bver1 / 1000) + (pg * xlWorkSheet2.Cells(23, 10).value * bver2 / 1000) + (pg * xlWorkSheet2.Cells(26, 10).value * a8 * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(86, 10).value * a10 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblcod3.Text = codleido.ToString
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True


                    ' *********************** Carpintería MI TES con 4HM ************************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 6) And (cmblac.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double


                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (a2 * au2 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (a5 * au5 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If
                    If (a7 * au7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (a8 * au8 < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (2 * AHF < 4700) Then
                        cr5 = 1.2
                    Else : cr5 = 1.0
                    End If

                    If (a10 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(12, 10).value * a2 * au2 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(12, 10).value * a5 * au5 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(13, 10).value * bver2 * 4 / 1000) + (pg * xlWorkSheet2.Cells(14, 10).value * bver2 * 2 / 1000) + (pg * xlWorkSheet2.Cells(15, 10).value * a7 * au7 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(23, 10).value * bver1 / 1000) + (pg * xlWorkSheet2.Cells(23, 10).value * bver2 / 1000) + (pg * xlWorkSheet2.Cells(26, 10).value * a8 * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(86, 10).value * a10 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblcod3.Text = codleido.ToString
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería MI TES con 1HF + 2HM ********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 7) And (cmblac.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double


                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (a2 * au2 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (a5 * au5 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If
                    If (a7 * au7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (a8 * au8 < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (AHF < 4700) Then
                        cr5 = 1.2
                    Else : cr5 = 1.0
                    End If

                    If (a10 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(12, 10).value * a2 * au2 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(12, 10).value * a5 * au5 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(13, 10).value * bver1 / 1000) + (pg * xlWorkSheet2.Cells(13, 10).value * bver2 * 2 / 1000) + (pg * xlWorkSheet2.Cells(14, 10).value * bver2 * 1 / 1000) + (pg * xlWorkSheet2.Cells(15, 10).value * a7 * au7 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(21, 10).value * 1 * AHF * 1.25 * cr5 / 1000) + (pg * xlWorkSheet2.Cells(23, 10).value * bver1 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(23, 10).value * bver2 / 1000) + (pg * xlWorkSheet2.Cells(26, 10).value * a8 * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(86, 10).value * a10 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería MI TES con 2HM TES*********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 8) And (cmblac.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double


                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (a2 * au2 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (a5 * au5 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If
                    If (a7 * au7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (a8 * au8 < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (2 * AHF < 4700) Then
                        cr5 = 1.2
                    Else : cr5 = 1.0
                    End If

                    If (a10 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(12, 10).value * a2 * au2 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(12, 10).value * a5 * au5 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(13, 10).value * bver2 * 2 / 1000) + (pg * xlWorkSheet2.Cells(14, 10).value * bver2 / 1000) + (pg * xlWorkSheet2.Cells(15, 10).value * a7 * au7 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(23, 10).value * bver2 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(26, 10).value * a8 * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(86, 10).value * a10 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                End If

            Catch ex As Exception

            End Try

            ' *********************** BUCLE IF PARA CALCULAR LOS COSTOS DE LOS FRENTES DE PUERTA *********************************************

            ' *********************** Carpintería Full Glass con 2HF + 2HM *************************************************************************** 

            Try

                If ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 1) And (cmblac.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (b3 * bu3 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (b6 * bu6 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (b7 * bu7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (wh < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (b11 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(35, 10).value * bver1 / 1000) + (pg * xlWorkSheet2.Cells(34, 10).value * bver1 / 1000) + (pg * xlWorkSheet2.Cells(35, 10).value * bver2 / 1000) + (pg * xlWorkSheet2.Cells(34, 10).value * bver2 / 1000) + (pg * xlWorkSheet2.Cells(33, 10).value * b3 * bu3 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(32, 10).value * b6 * bu6 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(33, 10).value * b7 * bu7 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(36, 10).value * wh * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * b11 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Full Glass con 2HM ******************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 2) And (cmblac.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (b3 * bu3 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (b6 * bu6 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (b7 * bu7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (wh < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (b11 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(35, 10).value * bver2 / 1000) + (pg * xlWorkSheet2.Cells(34, 10).value * bver2 / 1000) + (pg * xlWorkSheet2.Cells(32, 10).value * b6 * bu6 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(33, 10).value * b7 * bu7 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(36, 10).value * wh * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * b11 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Full Glass con 1HF + 1HM ***************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 3) And (cmblac.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (b3 * bu3 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (b6 * bu6 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (b7 * bu7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (wh < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (b11 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(35, 10).value * bver1 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(34, 10).value * bver1 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(35, 10).value * bver2 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(34, 10).value * bver2 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(33, 10).value * b3 * bu3 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(32, 10).value * b6 * bu6 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(33, 10).value * b7 * bu7 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(36, 4).value * wh * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * b11 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Full Glass con 1HM *********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 4) And (cmblac.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (b3 * bu3 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (b6 * bu6 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (b7 * bu7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (wh < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (b11 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(35, 10).value * bver2 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(34, 10).value * bver2 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(32, 10).value * b6 * bu6 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(33, 10).value * b7 * bu7 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(36, 10).value * wh * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * b11 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Full Glass TES con 2HF + 4HM *********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 5) And (cmblac.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim ff As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (b3 * bu3 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (b6 * bu6 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (b7 * bu7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (wh < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (b11 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(35, 10).value * bver1 / 1000) + (pg * xlWorkSheet2.Cells(34, 10).value * bver1 / 1000) + (pg * xlWorkSheet2.Cells(35, 10).value * bver2 * 3 / 1000) + (pg * xlWorkSheet2.Cells(34, 10).value * bver2 / 1000) + (pg * xlWorkSheet2.Cells(33, 10).value * b3 * bu3 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(32, 10).value * b6 * bu6 * 1.25 * cr2 / 1000) + (xlWorkSheet2.Cells(33, 10).value * b7 * bu7 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(36, 10).value * wh * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(86, 10).value * b11 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Full Glass TES con 4HM ************************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 6) And (cmblac.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (b3 * bu3 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (b6 * bu6 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (b7 * bu7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (wh < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (b11 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(35, 10).value * bver2 * 3 / 1000) + (pg * xlWorkSheet2.Cells(34, 10).value * bver2 / 1000) + (pg * xlWorkSheet2.Cells(32, 10).value * b6 * bu6 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(33, 10).value * b7 * bu7 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(36, 10).value * wh * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(86, 10).value * b11 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Full Glass TES con 1HF + 2HM ********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 7) And (cmblac.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (b3 * bu3 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (b6 * bu6 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (b7 * bu7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (wh < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (b11 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(35, 10).value * bver1 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(34, 10).value * bver1 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(35, 10).value * bver2 * 1.5 / 1000) + (pg * xlWorkSheet2.Cells(34, 10).value * bver2 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(33, 10).value * b3 * bu3 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(32, 10).value * b6 * bu6 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(33, 10).value * b7 * bu7 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(36, 10).value * wh * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(86, 10).value * b11 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Full Glass TES con 2HM TES *********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 8) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (b3 * bu3 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (b6 * bu6 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (b7 * bu7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (wh < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (b11 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(35, 10).value * bver2 * 1.5 / 1000) + (pg * xlWorkSheet2.Cells(34, 10).value * bver2 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(32, 10).value * b6 * bu6 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(33, 10).value * b7 * bu7 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(36, 10).value * wh * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(86, 10).value * b11 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                End If

            Catch ex As Exception

            End Try

            Try

                ' *********************** Carpintería Plintos silicona superior e inferior con 2HF + 2HM *************************************************************************** 

                If (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 1 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Plintos silicona superior e inferior con 2HM ******************************************************************************

                ElseIf (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 2 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(46, 4).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Plintos silicona superior e inferior con 1HF + 1HM ***************************************************************************

                ElseIf (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 3 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Plintos silicona superior e inferior con 1HM ******************************************

                ElseIf (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 4 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Plintos silicona superior e inferior TES con 2HF + 4HM *********************************************************************************

                ElseIf (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 5 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(86, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True


                    ' *********************** Carpintería Plintos silicona superior e inferior TES con 4HM **************************************************

                ElseIf (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 6 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(86, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Plintos silicona superior e inferior TES con 1HF + 2HM *****************************************

                ElseIf (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 7 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(86, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Plintos silicona superior e inferior TES con 2HM *****************

                ElseIf (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 8 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(86, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                End If

            Catch ex As Exception

            End Try

            Try

                ' *********************** Carpintería Plintos silicona solo superior  con 2HF + 2HM **************************************************** 

                If (cmbmodcar.SelectedIndex = 4 And cmbconf.SelectedIndex = 1 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Plintos silicona solo superior con 2HM **************************************************

                ElseIf (cmbmodcar.SelectedIndex = 4 And cmbconf.SelectedIndex = 2 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Plintos silicona solo superior con 1HF + 1HM **************************************************

                ElseIf (cmbmodcar.SelectedIndex = 4 And cmbconf.SelectedIndex = 3 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Plintos silicona solo superior con 1HM ******************************************

                ElseIf (cmbmodcar.SelectedIndex = 4 And cmbconf.SelectedIndex = 4 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                End If

            Catch ex As Exception

            End Try


            Try

                ' *********************** Carpintería Plintos  superior de pinza e inferior de silicona con 2HF + 2HM *************************************************************************** 

                If (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 1 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Plintos  superior de pinza e inferior de silicona con 2HM ******************************************************************************

                ElseIf (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 2 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(46, 4).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Plintos  superior de pinza e inferior de silicona con 1HF + 1HM ***************************************************************************

                ElseIf (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 3 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Plintos  superior de pinza e inferior de silicona con 1HM ******************************************

                ElseIf (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 4 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Plintos  superior de pinza e inferior de silicona TES con 2HF + 4HM *********************************************************************************

                ElseIf (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 5 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(86, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True


                    ' *********************** Carpintería Plintos superior de pinza e inferior de silicona TES con 4HM **************************************************

                ElseIf (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 6 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(86, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Plintos  superior de pinza e inferior de silicona TES con 1HF + 2HM *****************************************

                ElseIf (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 7 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(86, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Plintos  superior de pinza e inferior de silicona TES con 2HM *****************

                ElseIf (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 8 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(86, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                End If

            Catch ex As Exception

            End Try

            Try

                ' *********************** Carpintería Plintos pinza solo superior  con 2HF + 2HM **************************************************** 

                If (cmbmodcar.SelectedIndex = 6 And cmbconf.SelectedIndex = 1 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Plintos pinza solo superior con 2HM **************************************************

                ElseIf (cmbmodcar.SelectedIndex = 6 And cmbconf.SelectedIndex = 2 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Plintos pinza solo superior con 1HF + 1HM **************************************************

                ElseIf (cmbmodcar.SelectedIndex = 6 And cmbconf.SelectedIndex = 3 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería Plintos pinza solo superior con 1HM ******************************************

                ElseIf (cmbmodcar.SelectedIndex = 6 And cmbconf.SelectedIndex = 4 And cmblac.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                End If

            Catch ex As Exception

            End Try


            Try

                ' *********************** Carpintería BW52 con 2HF + 2HM *************************************************************************** 

                If ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 1) And (cmblac.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double
                    Dim cr6 As Double
                    Dim cr7 As Double
                    Dim cr8 As Double
                    Dim cr9 As Double

                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 1.2 * 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 1.2 * 4700
                    End If

                    If (d4 * du4 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (d5 * du5 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (d7 * du7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (d8 * du8 < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (d20 * du20 < 4700) Then
                        cr5 = 1.2
                    Else : cr5 = 1.0
                    End If

                    If (d13 * du13 < 4700) Then
                        cr6 = 1.2
                    Else : cr6 = 1.0
                    End If

                    If (d14 * du14 < 4700) Then
                        cr7 = 1.2
                    Else : cr7 = 1.0
                    End If

                    If (d17 * du17 < 4700) Then
                        cr8 = 1.2
                    Else : cr8 = 1.0
                    End If

                    If (d18 * du18 < 4700) Then
                        cr9 = 1.2
                    Else : cr9 = 1.0
                    End If

                    If (d22 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(67, 10).value * d4 * du4 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(67, 10).value * bver1 * 2 / 1000) + (pg * xlWorkSheet2.Cells(68, 10).value * d5 * du5 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * bver3 * 2 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * d7 * du7 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(69, 10).value * bver1 / 1000) + (pg * xlWorkSheet2.Cells(70, 10).value * bver1 / 1000) + (pg * xlWorkSheet2.Cells(73, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(72, 10).value * d20 * du20 * 1.25 * cr5 / 1000) + (pg * xlWorkSheet2.Cells(67, 10).value * d13 * du13 * 1.25 * cr6 / 1000) + (pg * xlWorkSheet2.Cells(67, 10).value * bver2 * 2 / 1000) + (pg * xlWorkSheet2.Cells(68, 10).value * d14 * du14 * 1.25 * cr7 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * bver4 * 2 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * d17 * du17 * 1.25 * cr8 / 1000) + (pg * xlWorkSheet2.Cells(74, 10).value * d18 * du18 * 1.25 * cr9 / 1000) + (pg * xlWorkSheet2.Cells(78, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(69, 10).value * bver2 / 1000) + (pg * xlWorkSheet2.Cells(71, 10).value * bver2 / 1000) + (pg * xlWorkSheet2.Cells(70, 10).value * bver2 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * d22 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería BW52 con 2HM ******************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 2) And (cmblac.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double
                    Dim cr6 As Double
                    Dim cr7 As Double
                    Dim cr8 As Double
                    Dim cr9 As Double


                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 1.2 * 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 1.2 * 4700
                    End If

                    If (d4 * du4 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (d5 * du5 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (d7 * du7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (d8 * du8 < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (d20 * du20 < 4700) Then
                        cr5 = 1.2
                    Else : cr5 = 1.0
                    End If

                    If (d13 * du13 < 4700) Then
                        cr6 = 1.2
                    Else : cr6 = 1.0
                    End If

                    If (d14 * du14 < 4700) Then
                        cr7 = 1.2
                    Else : cr7 = 1.0
                    End If

                    If (d17 * du17 < 4700) Then
                        cr8 = 1.2
                    Else : cr8 = 1.0
                    End If

                    If (d18 * du18 < 4700) Then
                        cr9 = 1.2
                    Else : cr9 = 1.0
                    End If

                    If (d22 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(67, 10).value * d13 * du13 * 1.25 * cr6 / 1000) + (pg * xlWorkSheet2.Cells(67, 10).value * bver2 * 2 / 1000) + (pg * xlWorkSheet2.Cells(68, 10).value * d14 * du14 * 1.25 * cr7 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * bver4 * 2 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * d17 * du17 * 1.25 * cr8 / 1000) + (pg * xlWorkSheet2.Cells(74, 10).value * d18 * du18 * 1.25 * cr9 / 1000) + (pg * xlWorkSheet2.Cells(78, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(69, 10).value * bver2 / 1000) + (pg * xlWorkSheet2.Cells(71, 10).value * bver2 / 1000) + (pg * xlWorkSheet2.Cells(70, 10).value * bver2 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * d22 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería BW52 con 1HF + 1HM ***************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 3) And (cmblac.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double
                    Dim cr6 As Double
                    Dim cr7 As Double
                    Dim cr8 As Double
                    Dim cr9 As Double

                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 1.2 * 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 1.2 * 4700
                    End If

                    If (d4 * du4 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (d5 * du5 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (d7 * du7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (d8 * du8 < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (d20 * du20 < 4700) Then
                        cr5 = 1.2
                    Else : cr5 = 1.0
                    End If

                    If (d13 * du13 < 4700) Then
                        cr6 = 1.2
                    Else : cr6 = 1.0
                    End If

                    If (d14 * du14 < 4700) Then
                        cr7 = 1.2
                    Else : cr7 = 1.0
                    End If

                    If (d17 * du17 < 4700) Then
                        cr8 = 1.2
                    Else : cr8 = 1.0
                    End If

                    If (d18 * du18 < 4700) Then
                        cr9 = 1.2
                    Else : cr9 = 1.0
                    End If

                    If (d22 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(67, 10).value * d4 * du4 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(67, 10).value * bver1 * 1 / 1000) + (pg * xlWorkSheet2.Cells(68, 10).value * d5 * du5 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * bver3 * 1 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * d7 * du7 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(69, 10).value * bver1 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(70, 10).value * bver1 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(73, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(72, 10).value * d20 * du20 * 1.25 * cr5 / 1000) + (pg * xlWorkSheet2.Cells(67, 10).value * d13 * du13 * 1.25 * cr6 / 1000) + (pg * xlWorkSheet2.Cells(67, 10).value * bver2 * 1 / 1000) + (pg * xlWorkSheet2.Cells(68, 10).value * d14 * du14 * 1.25 * cr7 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * bver4 * 1 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * d17 * du17 * 1.25 * cr8 / 1000) + (pg * xlWorkSheet2.Cells(74, 10).value * d18 * du18 * 1.25 * cr9 / 1000) + (pg * xlWorkSheet2.Cells(78, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(69, 10).value * bver2 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(71, 10).value * bver2 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(70, 10).value * bver2 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * d22 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería BW52 con 1HM *********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 4) And (cmblac.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double
                    Dim cr6 As Double
                    Dim cr7 As Double
                    Dim cr8 As Double
                    Dim cr9 As Double

                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 1.2 * 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 1.2 * 4700
                    End If

                    If (d4 * du4 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (d5 * du5 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (d7 * du7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (d8 * du8 < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (d20 * du20 < 4700) Then
                        cr5 = 1.2
                    Else : cr5 = 1.0
                    End If

                    If (d13 * du13 < 4700) Then
                        cr6 = 1.2
                    Else : cr6 = 1.0
                    End If

                    If (d14 * du14 < 4700) Then
                        cr7 = 1.2
                    Else : cr7 = 1.0
                    End If

                    If (d17 * du17 < 4700) Then
                        cr8 = 1.2
                    Else : cr8 = 1.0
                    End If

                    If (d18 * du18 < 4700) Then
                        cr9 = 1.2
                    Else : cr9 = 1.0
                    End If

                    If (d22 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(67, 10).value * d13 * du13 * 1.25 * cr6 / 1000) + (pg * xlWorkSheet2.Cells(67, 10).value * bver2 / 1000) + (pg * xlWorkSheet2.Cells(68, 10).value * d14 * du14 * 1.25 * cr7 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * bver4 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * d17 * du17 * 1.25 * cr8 / 1000) + (pg * xlWorkSheet2.Cells(74, 10).value * d18 * du18 * 1.25 * cr9 / 1000) + (pg * xlWorkSheet2.Cells(78, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(69, 10).value * bver2 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(71, 10).value * bver2 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(70, 10).value * bver2 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(85, 10).value * d22 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería BW52 TES con 2HF + 4HM *******************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 5) And (cmblac.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double
                    Dim cr6 As Double
                    Dim cr7 As Double
                    Dim cr8 As Double
                    Dim cr9 As Double

                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 1.2 * 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 1.2 * 4700
                    End If

                    If (d4 * du4 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (d5 * du5 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (d7 * du7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (d8 * du8 < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (d20 * du20 < 4700) Then
                        cr5 = 1.2
                    Else : cr5 = 1.0
                    End If

                    If (d13 * du13 < 4700) Then
                        cr6 = 1.2
                    Else : cr6 = 1.0
                    End If

                    If (d14 * du14 < 4700) Then
                        cr7 = 1.2
                    Else : cr7 = 1.0
                    End If

                    If (d17 * du17 < 4700) Then
                        cr8 = 1.2
                    Else : cr8 = 1.0
                    End If

                    If (d18 * du18 < 4700) Then
                        cr9 = 1.2
                    Else : cr9 = 1.0
                    End If

                    If (d22 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(67, 10).value * d4 * du4 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(67, 10).value * bver1 * 2 / 1000) + (pg * xlWorkSheet2.Cells(68, 10).value * d5 * du5 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * bver3 * 2 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * d7 * du7 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(69, 10).value * bver1 / 1000) + (pg * xlWorkSheet2.Cells(70, 10).value * bver1 / 1000) + (pg * xlWorkSheet2.Cells(73, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(72, 10).value * d20 * du20 * 1.25 * cr5 / 1000) + (pg * xlWorkSheet2.Cells(67, 10).value * d13 * du13 * 1.25 * cr6 / 1000) + (pg * xlWorkSheet2.Cells(67, 10).value * bver2 * 4 / 1000) + (pg * xlWorkSheet2.Cells(68, 10).value * d14 * du14 * 1.25 * cr7 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * bver4 * 4 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * d17 * du17 * 1.25 * cr8 / 1000) + (pg * xlWorkSheet2.Cells(74, 10).value * d18 * du18 * 1.25 * cr9 / 1000) + (pg * xlWorkSheet2.Cells(78, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(69, 10).value * bver2 * 3 / 1000) + (pg * xlWorkSheet2.Cells(71, 10).value * bver2 / 1000) + (pg * xlWorkSheet2.Cells(70, 10).value * bver2 * 2 / 1000) + (pg * xlWorkSheet2.Cells(86, 10).value * d22 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería BW52 TES con 4HM ************************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 6) And (cmblac.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double
                    Dim cr6 As Double
                    Dim cr7 As Double
                    Dim cr8 As Double
                    Dim cr9 As Double

                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 1.2 * 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 1.2 * 4700
                    End If

                    If (d4 * du4 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (d5 * du5 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (d7 * du7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (d8 * du8 < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (d20 * du20 < 4700) Then
                        cr5 = 1.2
                    Else : cr5 = 1.0
                    End If

                    If (d13 * du13 < 4700) Then
                        cr6 = 1.2
                    Else : cr6 = 1.0
                    End If

                    If (d14 * du14 < 4700) Then
                        cr7 = 1.2
                    Else : cr7 = 1.0
                    End If

                    If (d17 * du17 < 4700) Then
                        cr8 = 1.2
                    Else : cr8 = 1.0
                    End If

                    If (d18 * du18 < 4700) Then
                        cr9 = 1.2
                    Else : cr9 = 1.0
                    End If

                    If (d22 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(67, 10).value * d13 * du13 * 1.25 * cr6 / 1000) + (pg * xlWorkSheet2.Cells(67, 10).value * bver2 * 4 / 1000) + (pg * xlWorkSheet2.Cells(68, 10).value * d14 * du14 * 1.25 * cr7 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * bver4 * 4 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * d17 * du17 * 1.25 * cr8 / 1000) + (pg * xlWorkSheet2.Cells(74, 10).value * d18 * du18 * 1.25 * cr9 / 1000) + (pg * xlWorkSheet2.Cells(78, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(69, 10).value * bver2 * 2 / 1000) + (pg * xlWorkSheet2.Cells(71, 10).value * bver2 / 1000) + (pg * xlWorkSheet2.Cells(70, 10).value * bver2 * 2 / 1000) + (pg * xlWorkSheet2.Cells(86, 10).value * d22 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería BW52 TES con 1HF + 2HM ********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 7) And (cmblac.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double
                    Dim cr6 As Double
                    Dim cr7 As Double
                    Dim cr8 As Double
                    Dim cr9 As Double

                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 1.2 * 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 1.2 * 4700
                    End If

                    If (d4 * du4 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (d5 * du5 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (d7 * du7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (d8 * du8 < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (d20 * du20 < 4700) Then
                        cr5 = 1.2
                    Else : cr5 = 1.0
                    End If

                    If (d13 * du13 < 4700) Then
                        cr6 = 1.2
                    Else : cr6 = 1.0
                    End If

                    If (d14 * du14 < 4700) Then
                        cr7 = 1.2
                    Else : cr7 = 1.0
                    End If

                    If (d17 * du17 < 4700) Then
                        cr8 = 1.2
                    Else : cr8 = 1.0
                    End If

                    If (d18 * du18 < 4700) Then
                        cr9 = 1.2
                    Else : cr9 = 1.0
                    End If

                    If (d22 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(67, 10).value * d4 * du4 * 1.25 * cr1 / 1000) + (pg * xlWorkSheet2.Cells(67, 10).value * bver1 / 1000) + (pg * xlWorkSheet2.Cells(68, 10).value * d5 * du5 * 1.25 * cr2 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * bver3 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * d7 * du7 * 1.25 * cr3 / 1000) + (pg * xlWorkSheet2.Cells(69, 10).value * bver1 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(70, 10).value * bver1 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(73, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(72, 10).value * d20 * du20 * 1.25 * cr5 / 1000) + (pg * xlWorkSheet2.Cells(67, 10).value * d13 * du13 * 1.25 * cr6 / 1000) + (pg * xlWorkSheet2.Cells(67, 10).value * bver2 * 2 / 1000) + (pg * xlWorkSheet2.Cells(68, 10).value * d14 * du14 * 1.25 * cr7 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * bver4 * 2 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * d17 * du17 * 1.25 * cr8 / 1000) + (pg * xlWorkSheet2.Cells(74, 10).value * d18 * du18 * 1.25 * cr9 / 1000) + (pg * xlWorkSheet2.Cells(78, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(69, 10).value * bver2 * 1.5 / 1000) + (pg * xlWorkSheet2.Cells(71, 10).value * bver2 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(70, 10).value * bver2 * 1 / 1000) + (pg * xlWorkSheet2.Cells(86, 10).value * d22 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                    ' *********************** Carpintería BW52 TES con 2HM ************************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 8) And (cmblac.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim ff As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double
                    Dim cr6 As Double
                    Dim cr7 As Double
                    Dim cr8 As Double
                    Dim cr9 As Double

                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 1.2 * 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 1.2 * 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 1.2 * 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 1.2 * 4700
                    End If

                    If (d4 * du4 < 4700) Then
                        cr1 = 1.2
                    Else : cr1 = 1.0
                    End If

                    If (d5 * du5 < 4700) Then
                        cr2 = 1.2
                    Else : cr2 = 1.0
                    End If

                    If (d7 * du7 < 4700) Then
                        cr3 = 1.2
                    Else : cr3 = 1.0
                    End If

                    If (d8 * du8 < 4700) Then
                        cr4 = 1.2
                    Else : cr4 = 1.0
                    End If

                    If (d20 * du20 < 4700) Then
                        cr5 = 1.2
                    Else : cr5 = 1.0
                    End If

                    If (d13 * du13 < 4700) Then
                        cr6 = 1.2
                    Else : cr6 = 1.0
                    End If

                    If (d14 * du14 < 4700) Then
                        cr7 = 1.2
                    Else : cr7 = 1.0
                    End If

                    If (d17 * du17 < 4700) Then
                        cr8 = 1.2
                    Else : cr8 = 1.0
                    End If

                    If (d18 * du18 < 4700) Then
                        cr9 = 1.2
                    Else : cr9 = 1.0
                    End If

                    If (d22 < 4700) Then
                        cpl = 1.2
                    Else : cpl = 1.0
                    End If

                    ff = (pg * xlWorkSheet2.Cells(67, 10).value * d13 * du13 * 1.25 * cr6 / 1000) + (pg * xlWorkSheet2.Cells(67, 10).value * bver2 * 2 / 1000) + (pg * xlWorkSheet2.Cells(68, 10).value * d14 * du14 * 1.25 * cr7 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * bver4 * 2 / 1000) + (pg * xlWorkSheet2.Cells(75, 10).value * d17 * du17 * 1.25 * cr8 / 1000) + (pg * xlWorkSheet2.Cells(74, 10).value * d18 * du18 * 1.25 * cr9 / 1000) + (pg * xlWorkSheet2.Cells(78, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pg * xlWorkSheet2.Cells(69, 10).value * bver2 * 1.5 / 1000) + (pg * xlWorkSheet2.Cells(71, 10).value * bver2 * 0.5 / 1000) + (pg * xlWorkSheet2.Cells(70, 10).value * bver2 * 1 / 1000) + (pg * xlWorkSheet2.Cells(86, 10).value * d22 * 1.25 * cpl / 1000)

                    If (cmblac.SelectedIndex = 1) Then
                        precio = ff * 2
                    ElseIf (cmblac.SelectedIndex = 2 And ff < 100) Then
                        precio = 100 * 2
                    ElseIf (cmblac.SelectedIndex = 3 And ff < 150) Then
                        precio = 150 * 2
                    ElseIf (cmblac.SelectedIndex = 4 And ff < 200) Then
                        precio = 200 * 2
                    End If

                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod3.Text = codleido.ToString
                    lbldes3.Text = descripcion
                    lbluds3.Text = 1
                    lblpre3.Text = precio
                    importe = precio * 1
                    lblimp3.Text = importe.ToString

                    lblcod3.Visible = True
                    lbldes3.Visible = True
                    lbluds3.Visible = True
                    lblpre3.Visible = True
                    lblimp3.Visible = True

                End If

            Catch ex As Exception

            End Try


            ' ******** ANODIZADO *****************************************************************************
            ' **** Condicional para asignar a cada tipo de anodizados su tarifa correspondiente *********

            Try

                If (cmbano.SelectedIndex = 1) Then
                    pa = 1.46
                ElseIf (cmbano.SelectedIndex = 2) Then
                    pa = 1.6
                ElseIf (cmbano.SelectedIndex = 3) Then
                    pa = 2.33
                ElseIf (cmbano.SelectedIndex = 4) Then
                    pa = 3.56
                End If

            Catch ex As Exception

            End Try

            ' *********************** BUCLE IF PARA CALCULAR LOS COSTOS DE LOS ANODIZADOS *********************************************

            ' *********************** Carpintería MI con 2HF + 2HM *************************************************************************** 

            Try

                If ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 1) And (cmbano.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double

                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (a2 * au2 < 4500) And (a2 * au2 > 3000) Then
                        cr1 = 1.4
                    ElseIf (a2 * au2 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (a5 * au5 < 4500) And (a5 * au5 > 3000) Then
                        cr2 = 1.4
                    ElseIf (a5 * au5 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (a7 * au7 < 4500) And (a7 * au7 > 3000) Then
                        cr3 = 1.4
                    ElseIf (a7 * au7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (a8 * au8 < 4500) And (a8 * au8 > 3000) Then
                        cr4 = 1.4
                    ElseIf (a8 * au8 < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (2 * AHF < 4500) And (2 * AHF > 3000) Then
                        cr5 = 1.4
                    ElseIf (2 * AHF < 3000) Then
                        cr5 = 1.8
                    Else : cr5 = 1.0
                    End If

                    If (a10 < 4500) And (a10 > 3000) Then
                        cpl = 1.4
                    ElseIf (a10 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(12, 10).value * a2 * au2 * 1.25 * cr1 / 1000) + (pgr * a2 * au2 * 1.25 * xlWorkSheet2.Cells(12, 13).value * xlWorkSheet2.Cells(12, 12).value / 1000) + (pa * xlWorkSheet2.Cells(12, 10).value * a5 * au5 * 1.25 * cr2 / 1000) + (pgr * a5 * au5 * 1.25 * xlWorkSheet2.Cells(12, 13).value * xlWorkSheet2.Cells(12, 12).value / 1000) + (pa * xlWorkSheet2.Cells(13, 10).value * bver1 * 2 / 1000) + (pgr * bver1 * xlWorkSheet2.Cells(13, 12).value * xlWorkSheet2.Cells(13, 13).value / 1000) + (pa * xlWorkSheet2.Cells(13, 10).value * bver2 * 2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(13, 12).value * xlWorkSheet2.Cells(13, 13).value / 1000) + (pa * xlWorkSheet2.Cells(14, 10).value * bver1 / 1000) + (pgr * bver1 * xlWorkSheet2.Cells(14, 12).value * xlWorkSheet2.Cells(14, 13).value / 1000) + (pa * xlWorkSheet2.Cells(14, 10).value * bver2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(14, 12).value * xlWorkSheet2.Cells(14, 13).value / 1000) + (pa * xlWorkSheet2.Cells(15, 10).value * a7 * au7 * 1.25 * cr3 / 1000) + (pgr * a7 * au7 * 1.25 * xlWorkSheet2.Cells(15, 13).value * xlWorkSheet2.Cells(15, 12).value / 1000) + (pa * xlWorkSheet2.Cells(26, 10).value * a8 * 1.25 * cr4 / 1000) + (pgr * a8 * 1.25 * xlWorkSheet2.Cells(26, 13).value * xlWorkSheet2.Cells(26, 12).value / 1000) + (pa * xlWorkSheet2.Cells(21, 10).value * 2 * AHF * 1.25 * cr5 / 1000) + (pgr * 2 * AHF * 1.25 * xlWorkSheet2.Cells(21, 13).value * xlWorkSheet2.Cells(21, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * a10 * 1.25 * cpl / 1000) + (pgr * a10 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblcod21.Text = codleido.ToString
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblcod21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería MI con 2HM ******************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 2) And (cmbano.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double

                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (a2 * au2 < 4500) And (a2 * au2 > 3000) Then
                        cr1 = 1.4
                    ElseIf (a2 * au2 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (a5 * au5 < 4500) And (a5 * au5 > 3000) Then
                        cr2 = 1.4
                    ElseIf (a5 * au5 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (a7 * au7 < 4500) And (a7 * au7 > 3000) Then
                        cr3 = 1.4
                    ElseIf (a7 * au7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (a8 * au8 < 4500) And (a8 * au8 > 3000) Then
                        cr4 = 1.4
                    ElseIf (a8 * au8 < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (2 * AHF < 4500) And (2 * AHF > 3000) Then
                        cr5 = 1.4
                    ElseIf (2 * AHF < 3000) Then
                        cr5 = 1.8
                    Else : cr5 = 1.0
                    End If

                    If (a10 < 4500) And (a10 > 3000) Then
                        cpl = 1.4
                    ElseIf (a10 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(12, 10).value * a2 * au2 * 1.25 * cr1 / 1000) + (pgr * a2 * au2 * 1.25 * xlWorkSheet2.Cells(12, 13).value * xlWorkSheet2.Cells(12, 12).value / 1000) + (pa * xlWorkSheet2.Cells(12, 10).value * a5 * au5 * 1.25 * cr2 / 1000) + (pgr * a5 * au5 * 1.25 * xlWorkSheet2.Cells(12, 13).value * xlWorkSheet2.Cells(12, 12).value / 1000) + (pa * xlWorkSheet2.Cells(13, 10).value * bver2 * 2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(13, 12).value * xlWorkSheet2.Cells(13, 13).value / 1000) + (pa * xlWorkSheet2.Cells(14, 10).value * bver2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(14, 12).value * xlWorkSheet2.Cells(14, 13).value / 1000) + (pa * xlWorkSheet2.Cells(15, 10).value * a7 * au7 * 1.25 * cr3 / 1000) + (pgr * a7 * au7 * 1.25 * xlWorkSheet2.Cells(15, 13).value * xlWorkSheet2.Cells(15, 12).value / 1000) + (pa * xlWorkSheet2.Cells(26, 10).value * a8 * 1.25 * cr4 / 1000) + (pgr * a8 * 1.25 * xlWorkSheet2.Cells(26, 13).value * xlWorkSheet2.Cells(26, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * a10 * 1.25 * cpl / 1000) + (pgr * a10 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblcod21.Text = codleido.ToString
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblcod21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería MI con 1HF + 1HM ***************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 3) And (cmbano.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double


                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (a2 * au2 < 4500) And (a2 * au2 > 3000) Then
                        cr1 = 1.4
                    ElseIf (a2 * au2 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (a5 * au5 < 4500) And (a5 * au5 > 3000) Then
                        cr2 = 1.4
                    ElseIf (a5 * au5 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (a7 * au7 < 4500) And (a7 * au7 > 3000) Then
                        cr3 = 1.4
                    ElseIf (a7 * au7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (a8 * au8 < 4500) And (a8 * au8 > 3000) Then
                        cr4 = 1.4
                    ElseIf (a8 * au8 < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (1 * AHF < 4500) And (2 * AHF > 3000) Then
                        cr5 = 1.4
                    ElseIf (1 * AHF < 3000) Then
                        cr5 = 1.8
                    Else : cr5 = 1.0
                    End If

                    If (a10 < 4500) And (a10 > 3000) Then
                        cpl = 1.4
                    ElseIf (a10 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(12, 10).value * a2 * au2 * 1.25 * cr1 / 1000) + (pgr * a2 * au2 * 1.25 * xlWorkSheet2.Cells(12, 13).value * xlWorkSheet2.Cells(12, 12).value / 1000) + (pa * xlWorkSheet2.Cells(12, 10).value * a5 * au5 * 1.25 * cr2 / 1000) + (pgr * a5 * au5 * 1.25 * xlWorkSheet2.Cells(12, 13).value * xlWorkSheet2.Cells(12, 12).value / 1000) + (pa * xlWorkSheet2.Cells(13, 10).value * bver1 / 1000) + (pgr * bver1 * xlWorkSheet2.Cells(13, 12).value * xlWorkSheet2.Cells(13, 13).value / 1000) + (pa * xlWorkSheet2.Cells(13, 10).value * bver2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(13, 12).value * xlWorkSheet2.Cells(13, 13).value / 1000) + (pa * xlWorkSheet2.Cells(14, 10).value * bver1 * 0.5 / 1000) + (pgr * 0.5 * bver1 * xlWorkSheet2.Cells(14, 12).value * xlWorkSheet2.Cells(14, 13).value / 1000) + (pa * xlWorkSheet2.Cells(14, 10).value * bver2 * 0.5 / 1000) + (pgr * bver2 * 0.5 * xlWorkSheet2.Cells(14, 12).value * xlWorkSheet2.Cells(14, 13).value / 1000) + (pa * xlWorkSheet2.Cells(15, 10).value * a7 * au7 * 1.25 * cr3 / 1000) + (pgr * a7 * au7 * 1.25 * xlWorkSheet2.Cells(15, 13).value * xlWorkSheet2.Cells(15, 12).value / 1000) + (pa * xlWorkSheet2.Cells(21, 10).value * 1 * AHF * 1.25 * cr5 / 1000) + (pgr * 1 * AHF * 1.25 * xlWorkSheet2.Cells(21, 13).value * xlWorkSheet2.Cells(21, 12).value / 1000) + (pa * xlWorkSheet2.Cells(26, 10).value * a8 * 1.25 * cr4 / 1000) + (pgr * a8 * 1.25 * xlWorkSheet2.Cells(26, 13).value * xlWorkSheet2.Cells(26, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * a10 * 1.25 * cpl / 1000) + (pgr * a10 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblcod21.Text = codleido.ToString
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblcod21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería MI con 1HM *********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 4) And (cmbano.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double


                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (a2 * au2 < 4500) And (a2 * au2 > 3000) Then
                        cr1 = 1.4
                    ElseIf (a2 * au2 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (a5 * au5 < 4500) And (a5 * au5 > 3000) Then
                        cr2 = 1.4
                    ElseIf (a5 * au5 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (a7 * au7 < 4500) And (a7 * au7 > 3000) Then
                        cr3 = 1.4
                    ElseIf (a7 * au7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (a8 * au8 < 4500) And (a8 * au8 > 3000) Then
                        cr4 = 1.4
                    ElseIf (a8 * au8 < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (1 * AHF < 4500) And (2 * AHF > 3000) Then
                        cr5 = 1.4
                    ElseIf (1 * AHF < 3000) Then
                        cr5 = 1.8
                    Else : cr5 = 1.0
                    End If

                    If (a10 < 4500) And (a10 > 3000) Then
                        cpl = 1.4
                    ElseIf (a10 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(12, 10).value * a2 * au2 * 1.25 * cr1 / 1000) + (pgr * a2 * au2 * 1.25 * xlWorkSheet2.Cells(12, 13).value * xlWorkSheet2.Cells(12, 12).value / 1000) + (pgr * a5 * au5 * 1.25 * xlWorkSheet2.Cells(12, 13).value * xlWorkSheet2.Cells(12, 12).value / 1000) + (pa * xlWorkSheet2.Cells(12, 10).value * a5 * au5 * 1.25 * cr2 / 1000) + (pa * xlWorkSheet2.Cells(13, 10).value * bver2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(13, 12).value * xlWorkSheet2.Cells(13, 13).value / 1000) + (pa * xlWorkSheet2.Cells(14, 10).value * bver2 * 0.5 / 1000) + (pgr * bver2 * 0.5 * xlWorkSheet2.Cells(14, 12).value * xlWorkSheet2.Cells(14, 13).value / 1000) + (pa * xlWorkSheet2.Cells(15, 10).value * a7 * au7 * 1.25 * cr3 / 1000) + (pgr * a7 * au7 * 1.25 * xlWorkSheet2.Cells(15, 13).value * xlWorkSheet2.Cells(15, 12).value / 1000) + (pa * xlWorkSheet2.Cells(26, 4).value * a8 * 1.25 * cr5 / 1000) + (pgr * a8 * 1.25 * xlWorkSheet2.Cells(26, 13).value * xlWorkSheet2.Cells(26, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * a10 * 1.25 * cpl / 1000) + (pgr * a10 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblcod21.Text = codleido.ToString
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblcod21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería MI TES con 2HF + 4HM *********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 5) And (cmbano.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double


                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (a2 * au2 < 4500) And (a2 * au2 > 3000) Then
                        cr1 = 1.4
                    ElseIf (a2 * au2 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (a5 * au5 < 4500) And (a5 * au5 > 3000) Then
                        cr2 = 1.4
                    ElseIf (a5 * au5 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (a7 * au7 < 4500) And (a7 * au7 > 3000) Then
                        cr3 = 1.4
                    ElseIf (a7 * au7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (a8 * au8 < 4500) And (a8 * au8 > 3000) Then
                        cr4 = 1.4
                    ElseIf (a8 * au8 < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (2 * AHF < 4500) And (2 * AHF > 3000) Then
                        cr5 = 1.4
                    ElseIf (2 * AHF < 3000) Then
                        cr5 = 1.8
                    Else : cr5 = 1.0
                    End If

                    If (a10 < 4500) And (a10 > 3000) Then
                        cpl = 1.4
                    ElseIf (a10 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(12, 10).value * a2 * au2 * 1.25 * cr1 / 1000) + (pgr * a2 * au2 * 1.25 * xlWorkSheet2.Cells(12, 13).value * xlWorkSheet2.Cells(12, 12).value / 1000) + (pa * xlWorkSheet2.Cells(12, 10).value * a5 * au5 * 1.25 * cr2 / 1000) + (pgr * a5 * au5 * 1.25 * xlWorkSheet2.Cells(12, 13).value * xlWorkSheet2.Cells(12, 12).value / 1000) + (pa * xlWorkSheet2.Cells(13, 10).value * bver1 * 2 / 1000) + (pgr * bver1 * 2 * xlWorkSheet2.Cells(13, 12).value * xlWorkSheet2.Cells(13, 13).value / 1000) + (pa * xlWorkSheet2.Cells(13, 10).value * bver2 * 4 / 1000) + (pgr * bver2 * 4 * xlWorkSheet2.Cells(13, 12).value * xlWorkSheet2.Cells(13, 13).value / 1000) + (pa * xlWorkSheet2.Cells(14, 10).value * bver2 * 2 / 1000) + (pgr * bver2 * 2 * xlWorkSheet2.Cells(14, 12).value * xlWorkSheet2.Cells(14, 13).value / 1000) + (pa * xlWorkSheet2.Cells(15, 10).value * a7 * au7 * 1.25 * cr3 / 1000) + (pgr * a7 * au7 * 1.25 * xlWorkSheet2.Cells(15, 13).value * xlWorkSheet2.Cells(15, 12).value / 1000) + (pa * xlWorkSheet2.Cells(21, 10).value * 2 * AHF * 1.25 * cr5 / 1000) + (pgr * 2 * AHF * 1.25 * xlWorkSheet2.Cells(21, 13).value * xlWorkSheet2.Cells(21, 12).value / 1000) + (pa * xlWorkSheet2.Cells(23, 10).value * bver1 / 1000) + (pgr * bver1 * xlWorkSheet2.Cells(23, 12).value * xlWorkSheet2.Cells(23, 13).value / 1000) + (pa * xlWorkSheet2.Cells(23, 10).value * bver2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(23, 12).value * xlWorkSheet2.Cells(23, 13).value / 1000) + (pa * xlWorkSheet2.Cells(26, 10).value * a8 * 1.25 * cr4 / 1000) + (pgr * a8 * 1.25 * xlWorkSheet2.Cells(26, 13).value * xlWorkSheet2.Cells(26, 12).value / 1000) + (pa * xlWorkSheet2.Cells(86, 10).value * a10 * 1.25 * cpl / 1000) + (pgr * a10 * 1.25 * xlWorkSheet2.Cells(86, 13).value * xlWorkSheet2.Cells(86, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblcod21.Text = codleido.ToString
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería MI TES con 4HM ************************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 6) And (cmbano.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double

                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (a2 * au2 < 4500) And (a2 * au2 > 3000) Then
                        cr1 = 1.4
                    ElseIf (a2 * au2 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (a5 * au5 < 4500) And (a5 * au5 > 3000) Then
                        cr2 = 1.4
                    ElseIf (a5 * au5 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (a7 * au7 < 4500) And (a7 * au7 > 3000) Then
                        cr3 = 1.4
                    ElseIf (a7 * au7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (a8 * au8 < 4500) And (a8 * au8 > 3000) Then
                        cr4 = 1.4
                    ElseIf (a8 * au8 < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (2 * AHF < 4500) And (2 * AHF > 3000) Then
                        cr5 = 1.4
                    ElseIf (2 * AHF < 3000) Then
                        cr5 = 1.8
                    Else : cr5 = 1.0
                    End If

                    If (a10 < 4500) And (a10 > 3000) Then
                        cpl = 1.4
                    ElseIf (a10 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(12, 10).value * a2 * au2 * 1.25 * cr1 / 1000) + (pgr * a2 * au2 * 1.25 * xlWorkSheet2.Cells(12, 13).value * xlWorkSheet2.Cells(12, 12).value / 1000) + (pa * xlWorkSheet2.Cells(12, 10).value * a5 * au5 * 1.25 * cr2 / 1000) + (pgr * a5 * au5 * 1.25 * xlWorkSheet2.Cells(12, 13).value * xlWorkSheet2.Cells(12, 12).value / 1000) + (pa * xlWorkSheet2.Cells(13, 10).value * bver2 * 4 / 1000) + (pgr * 4 * bver2 * xlWorkSheet2.Cells(13, 12).value * xlWorkSheet2.Cells(13, 13).value / 1000) + (pa * xlWorkSheet2.Cells(14, 10).value * bver2 * 2 / 1000) + (pgr * bver2 * 2 * xlWorkSheet2.Cells(14, 12).value * xlWorkSheet2.Cells(14, 13).value / 1000) + (pa * xlWorkSheet2.Cells(15, 10).value * a7 * au7 * 1.25 * cr3 / 1000) + (pgr * a7 * au7 * 1.25 * xlWorkSheet2.Cells(15, 13).value * xlWorkSheet2.Cells(15, 12).value / 1000) + (pa * xlWorkSheet2.Cells(23, 10).value * bver1 / 1000) + (pgr * bver1 * xlWorkSheet2.Cells(23, 12).value * xlWorkSheet2.Cells(23, 13).value / 1000) + (pa * xlWorkSheet2.Cells(23, 10).value * bver2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(23, 12).value * xlWorkSheet2.Cells(23, 13).value / 1000) + (pa * xlWorkSheet2.Cells(26, 10).value * a8 * 1.25 * cr4 / 1000) + (pgr * a8 * 1.25 * xlWorkSheet2.Cells(26, 13).value * xlWorkSheet2.Cells(26, 12).value / 1000) + (pa * xlWorkSheet2.Cells(86, 10).value * a10 * 1.25 * cpl / 1000) + (pgr * a10 * 1.25 * xlWorkSheet2.Cells(86, 13).value * xlWorkSheet2.Cells(86, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblcod21.Text = codleido.ToString
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería MI TES con 1HF + 2HM ********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 7) And (cmbano.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double

                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (a2 * au2 < 4500) And (a2 * au2 > 3000) Then
                        cr1 = 1.4
                    ElseIf (a2 * au2 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (a5 * au5 < 4500) And (a5 * au5 > 3000) Then
                        cr2 = 1.4
                    ElseIf (a5 * au5 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (a7 * au7 < 4500) And (a7 * au7 > 3000) Then
                        cr3 = 1.4
                    ElseIf (a7 * au7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (a8 * au8 < 4500) And (a8 * au8 > 3000) Then
                        cr4 = 1.4
                    ElseIf (a8 * au8 < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (1 * AHF < 4500) And (2 * AHF > 3000) Then
                        cr5 = 1.4
                    ElseIf (1 * AHF < 3000) Then
                        cr5 = 1.8
                    Else : cr5 = 1.0
                    End If

                    If (a10 < 4500) And (a10 > 3000) Then
                        cpl = 1.4
                    ElseIf (a10 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(12, 10).value * a2 * au2 * 1.25 * cr1 / 1000) + (pgr * a2 * au2 * 1.25 * xlWorkSheet2.Cells(12, 13).value * xlWorkSheet2.Cells(12, 12).value / 1000) + (pa * xlWorkSheet2.Cells(12, 10).value * a5 * au5 * 1.25 * cr2 / 1000) + (pgr * a5 * au5 * 1.25 * xlWorkSheet2.Cells(12, 13).value * xlWorkSheet2.Cells(12, 12).value / 1000) + (pa * xlWorkSheet2.Cells(13, 10).value * bver1 / 1000) + (pgr * bver1 * xlWorkSheet2.Cells(13, 12).value * xlWorkSheet2.Cells(13, 13).value / 1000) + (pa * xlWorkSheet2.Cells(13, 10).value * bver2 * 2 / 1000) + (pgr * bver2 * 2 * xlWorkSheet2.Cells(13, 12).value * xlWorkSheet2.Cells(13, 13).value / 1000) + (pa * xlWorkSheet2.Cells(14, 10).value * bver2 * 1 / 1000) + (pgr * bver2 * 1 * xlWorkSheet2.Cells(14, 12).value * xlWorkSheet2.Cells(14, 13).value / 1000) + (pa * xlWorkSheet2.Cells(15, 10).value * a7 * au7 * 1.25 * cr3 / 1000) + (pgr * a7 * au7 * 1.25 * xlWorkSheet2.Cells(15, 13).value * xlWorkSheet2.Cells(15, 12).value / 1000) + (pa * xlWorkSheet2.Cells(21, 10).value * 1 * AHF * 1.25 * cr5 / 1000) + (pgr * 1 * AHF * 1.25 * xlWorkSheet2.Cells(21, 13).value * xlWorkSheet2.Cells(21, 12).value / 1000) + (pa * xlWorkSheet2.Cells(23, 10).value * bver1 * 0.5 / 1000) + (pgr * bver1 * 0.5 * xlWorkSheet2.Cells(23, 12).value * xlWorkSheet2.Cells(23, 13).value / 1000) + (pa * xlWorkSheet2.Cells(23, 10).value * bver2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(23, 12).value * xlWorkSheet2.Cells(23, 13).value / 1000) + (pa * xlWorkSheet2.Cells(26, 10).value * a8 * 1.25 * cr4 / 1000) + (pgr * a8 * 1.25 * xlWorkSheet2.Cells(26, 13).value * xlWorkSheet2.Cells(26, 12).value / 1000) + (pa * xlWorkSheet2.Cells(86, 10).value * a10 * 1.25 * cpl / 1000) + (pgr * a10 * 1.25 * xlWorkSheet2.Cells(86, 13).value * xlWorkSheet2.Cells(86, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería MI TES con 2HM TES*********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 1) And (cmbconf.SelectedIndex = 8) And (cmbano.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double

                    If (a1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (a4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (a2 * au2 < 4500) And (a2 * au2 > 3000) Then
                        cr1 = 1.4
                    ElseIf (a2 * au2 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (a5 * au5 < 4500) And (a5 * au5 > 3000) Then
                        cr2 = 1.4
                    ElseIf (a5 * au5 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (a7 * au7 < 4500) And (a7 * au7 > 3000) Then
                        cr3 = 1.4
                    ElseIf (a7 * au7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (a8 * au8 < 4500) And (a8 * au8 > 3000) Then
                        cr4 = 1.4
                    ElseIf (a8 * au8 < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (2 * AHF < 4500) And (2 * AHF > 3000) Then
                        cr5 = 1.4
                    ElseIf (2 * AHF < 3000) Then
                        cr5 = 1.8
                    Else : cr5 = 1.0
                    End If

                    If (a10 < 4500) And (a10 > 3000) Then
                        cpl = 1.4
                    ElseIf (a10 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(12, 10).value * a2 * au2 * 1.25 * cr1 / 1000) + (pgr * a2 * au2 * 1.25 * xlWorkSheet2.Cells(12, 13).value * xlWorkSheet2.Cells(12, 12).value / 1000) + (pa * xlWorkSheet2.Cells(12, 10).value * a5 * au5 * 1.25 * cr2 / 1000) + (pgr * a5 * au5 * 1.25 * xlWorkSheet2.Cells(12, 13).value * xlWorkSheet2.Cells(12, 12).value / 1000) + (pa * xlWorkSheet2.Cells(13, 10).value * bver2 * 2 / 1000) + (pgr * 2 * bver2 * xlWorkSheet2.Cells(13, 12).value * xlWorkSheet2.Cells(13, 13).value / 1000) + (pa * xlWorkSheet2.Cells(14, 10).value * bver2 / 1000) + (pgr * bver2 * 1 * xlWorkSheet2.Cells(14, 12).value * xlWorkSheet2.Cells(14, 13).value / 1000) + (pa * xlWorkSheet2.Cells(15, 10).value * a7 * au7 * 1.25 * cr3 / 1000) + (pgr * a7 * au7 * 1.25 * xlWorkSheet2.Cells(15, 13).value * xlWorkSheet2.Cells(15, 12).value / 1000) + (pa * xlWorkSheet2.Cells(23, 10).value * bver2 * 0.5 / 1000) + (pgr * bver2 * 0.5 * xlWorkSheet2.Cells(23, 12).value * xlWorkSheet2.Cells(23, 13).value / 1000) + (pa * xlWorkSheet2.Cells(26, 10).value * a8 * 1.25 * cr4 / 1000) + (pgr * a8 * 1.25 * xlWorkSheet2.Cells(26, 13).value * xlWorkSheet2.Cells(26, 12).value / 1000) + (pa * xlWorkSheet2.Cells(86, 10).value * a10 * 1.25 * cpl / 1000) + (pgr * a10 * 1.25 * xlWorkSheet2.Cells(86, 13).value * xlWorkSheet2.Cells(86, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                End If

            Catch ex As Exception

            End Try


                    ' *********************** BUCLE IF PARA CALCULAR LOS COSTOS DE LOS FRENTES DE PUERTA *********************************************

                    ' *********************** Carpintería Full Glass con 2HF + 2HM *************************************************************************** 

            Try

                If ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 1) And (cmbano.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (b3 * bu3 < 4500) And (b3 * bu3 > 3000) Then
                        cr1 = 1.4
                    ElseIf (b3 * bu3 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (b6 * bu6 < 4500) And (b6 * bu6 > 3000) Then
                        cr2 = 1.4
                    ElseIf (b6 * bu6 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (b7 * bu7 < 4500) And (b7 * bu7 > 3000) Then
                        cr3 = 1.4
                    ElseIf (b7 * bu7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (wh < 4500) And (wh > 3000) Then
                        cr4 = 1.4
                    ElseIf (wh < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (b11 < 4500) And (b11 > 3000) Then
                        cpl = 1.4
                    ElseIf (b11 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(35, 10).value * bver1 / 1000) + (pgr * bver1 * xlWorkSheet2.Cells(35, 13).value * xlWorkSheet2.Cells(35, 12).value / 1000) + (pa * xlWorkSheet2.Cells(34, 10).value * bver1 / 1000) + (pgr * bver1 * xlWorkSheet2.Cells(34, 13).value * xlWorkSheet2.Cells(34, 12).value / 1000) + (pa * xlWorkSheet2.Cells(35, 10).value * bver2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(35, 13).value * xlWorkSheet2.Cells(35, 12).value / 1000) + (pa * xlWorkSheet2.Cells(34, 10).value * bver2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(34, 13).value * xlWorkSheet2.Cells(34, 12).value / 1000) + (pa * xlWorkSheet2.Cells(33, 10).value * b3 * bu3 * 1.25 * cr1 / 1000) + (pgr * b3 * bu3 * 1.25 * xlWorkSheet2.Cells(33, 13).value * xlWorkSheet2.Cells(33, 12).value / 1000) + (pa * xlWorkSheet2.Cells(32, 10).value * b6 * bu6 * 1.25 * cr2 / 1000) + (pgr * b6 * bu6 * 1.25 * xlWorkSheet2.Cells(32, 13).value * xlWorkSheet2.Cells(32, 12).value / 1000) + (pa * xlWorkSheet2.Cells(33, 10).value * b7 * bu7 * 1.25 * cr3 / 1000) + (pgr * b7 * bu7 * 1.25 * xlWorkSheet2.Cells(33, 13).value * xlWorkSheet2.Cells(33, 12).value / 1000) + (pa * xlWorkSheet2.Cells(36, 10).value * wh * 1.25 * cr4 / 1000) + (pgr * wh * 1.25 * xlWorkSheet2.Cells(36, 13).value * xlWorkSheet2.Cells(36, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * b11 * 1.25 * cpl / 1000) + (pgr * b11 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True


                    ' *********************** Carpintería Full Glass con 2HM ******************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 2) And (cmbano.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (b3 * bu3 < 4500) And (b3 * bu3 > 3000) Then
                        cr1 = 1.4
                    ElseIf (b3 * bu3 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (b6 * bu6 < 4500) And (b6 * bu6 > 3000) Then
                        cr2 = 1.4
                    ElseIf (b6 * bu6 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (b7 * bu7 < 4500) And (b7 * bu7 > 3000) Then
                        cr3 = 1.4
                    ElseIf (b7 * bu7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (wh < 4500) And (wh > 3000) Then
                        cr4 = 1.4
                    ElseIf (wh < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (b11 < 4500) And (b11 > 3000) Then
                        cpl = 1.4
                    ElseIf (b11 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(35, 10).value * bver2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(35, 13).value * xlWorkSheet2.Cells(35, 12).value / 1000) + (pa * xlWorkSheet2.Cells(34, 10).value * bver2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(34, 13).value * xlWorkSheet2.Cells(34, 12).value / 1000) + (pa * xlWorkSheet2.Cells(32, 10).value * b6 * bu6 * 1.25 * cr2 / 1000) + (pgr * b6 * bu6 * 1.25 * xlWorkSheet2.Cells(32, 13).value * xlWorkSheet2.Cells(32, 12).value / 1000) + (pa * xlWorkSheet2.Cells(33, 10).value * b7 * bu7 * 1.25 * cr3 / 1000) + (pgr * b7 * bu7 * 1.25 * xlWorkSheet2.Cells(33, 13).value * xlWorkSheet2.Cells(33, 12).value / 1000) + (pg * xlWorkSheet2.Cells(36, 10).value * wh * 1.25 * cr4 / 1000) + (pgr * wh * 1.25 * xlWorkSheet2.Cells(36, 13).value * xlWorkSheet2.Cells(36, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * b11 * 1.25 * cpl / 1000) + (pgr * b11 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Full Glass con 1HF + 1HM ***************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 3) And (cmbano.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (b3 * bu3 < 4500) And (b3 * bu3 > 3000) Then
                        cr1 = 1.4
                    ElseIf (b3 * bu3 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (b6 * bu6 < 4500) And (b6 * bu6 > 3000) Then
                        cr2 = 1.4
                    ElseIf (b6 * bu6 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (b7 * bu7 < 4500) And (b7 * bu7 > 3000) Then
                        cr3 = 1.4
                    ElseIf (b7 * bu7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (wh < 4500) And (wh > 3000) Then
                        cr4 = 1.4
                    ElseIf (wh < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (b11 < 4500) And (b11 > 3000) Then
                        cpl = 1.4
                    ElseIf (b11 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(35, 10).value * bver1 * 0.5 / 1000) + (pgr * 0.5 * bver1 * xlWorkSheet2.Cells(35, 13).value * xlWorkSheet2.Cells(35, 12).value / 1000) + (pa * xlWorkSheet2.Cells(34, 10).value * bver1 * 0.5 / 1000) + (pgr * bver1 * 0.5 * xlWorkSheet2.Cells(34, 13).value * xlWorkSheet2.Cells(34, 12).value / 1000) + (pa * xlWorkSheet2.Cells(35, 10).value * bver2 * 0.5 / 1000) + (pgr * bver2 * 0.5 * xlWorkSheet2.Cells(35, 13).value * xlWorkSheet2.Cells(35, 12).value / 1000) + (pa * xlWorkSheet2.Cells(34, 10).value * bver2 * 0.5 / 1000) + (pgr * bver2 * 0.5 * xlWorkSheet2.Cells(34, 13).value * xlWorkSheet2.Cells(34, 12).value / 1000) + (pa * xlWorkSheet2.Cells(33, 10).value * b3 * bu3 * 1.25 * cr1 / 1000) + (pgr * b3 * bu3 * 1.25 * xlWorkSheet2.Cells(33, 13).value * xlWorkSheet2.Cells(33, 12).value / 1000) + (pa * xlWorkSheet2.Cells(32, 10).value * b6 * bu6 * 1.25 * cr2 / 1000) + (pgr * b6 * bu6 * 1.25 * xlWorkSheet2.Cells(32, 13).value * xlWorkSheet2.Cells(32, 12).value / 1000) + (pa * xlWorkSheet2.Cells(33, 10).value * b7 * bu7 * 1.25 * cr3 / 1000) + (pgr * b7 * bu7 * 1.25 * xlWorkSheet2.Cells(33, 13).value * xlWorkSheet2.Cells(33, 12).value / 1000) + (pg * xlWorkSheet2.Cells(36, 4).value * wh * 1.25 * cr4 / 1000) + (pgr * wh * 1.25 * xlWorkSheet2.Cells(36, 13).value * xlWorkSheet2.Cells(36, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * b11 * 1.25 * cpl / 1000) + (pgr * b11 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True


                    ' *********************** Carpintería Full Glass con 1HM *********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 4) And (cmbano.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (b3 * bu3 < 4500) And (b3 * bu3 > 3000) Then
                        cr1 = 1.4
                    ElseIf (b3 * bu3 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (b6 * bu6 < 4500) And (b6 * bu6 > 3000) Then
                        cr2 = 1.4
                    ElseIf (b6 * bu6 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (b7 * bu7 < 4500) And (b7 * bu7 > 3000) Then
                        cr3 = 1.4
                    ElseIf (b7 * bu7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (wh < 4500) And (wh > 3000) Then
                        cr4 = 1.4
                    ElseIf (wh < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (b11 < 4500) And (b11 > 3000) Then
                        cpl = 1.4
                    ElseIf (b11 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(35, 10).value * bver2 * 0.5 / 1000) + (pgr * bver2 * 0.5 * xlWorkSheet2.Cells(35, 13).value * xlWorkSheet2.Cells(35, 12).value / 1000) + (pa * xlWorkSheet2.Cells(34, 10).value * bver2 * 0.5 / 1000) + (pgr * bver2 * 0.5 * xlWorkSheet2.Cells(34, 13).value * xlWorkSheet2.Cells(34, 12).value / 1000) + (pa * xlWorkSheet2.Cells(32, 10).value * b6 * bu6 * 1.25 * cr2 / 1000) + (pgr * b6 * bu6 * 1.25 * xlWorkSheet2.Cells(32, 13).value * xlWorkSheet2.Cells(32, 12).value / 1000) + (pa * xlWorkSheet2.Cells(33, 10).value * b7 * bu7 * 1.25 * cr3 / 1000) + (pgr * b7 * bu7 * 1.25 * xlWorkSheet2.Cells(33, 13).value * xlWorkSheet2.Cells(33, 12).value / 1000) + (pg * xlWorkSheet2.Cells(36, 10).value * wh * 1.25 * cr4 / 1000) + (pgr * wh * 1.25 * xlWorkSheet2.Cells(36, 13).value * xlWorkSheet2.Cells(36, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * b11 * 1.25 * cpl / 1000) + (pgr * b11 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Full Glass TES con 2HF + 4HM *********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 5) And (cmbano.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (b3 * bu3 < 4500) And (b3 * bu3 > 3000) Then
                        cr1 = 1.4
                    ElseIf (b3 * bu3 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (b6 * bu6 < 4500) And (b6 * bu6 > 3000) Then
                        cr2 = 1.4
                    ElseIf (b6 * bu6 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (b7 * bu7 < 4500) And (b7 * bu7 > 3000) Then
                        cr3 = 1.4
                    ElseIf (b7 * bu7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (wh < 4500) And (wh > 3000) Then
                        cr4 = 1.4
                    ElseIf (wh < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (b11 < 4500) And (b11 > 3000) Then
                        cpl = 1.4
                    ElseIf (b11 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(35, 10).value * bver1 / 1000) + (pgr * bver1 * xlWorkSheet2.Cells(35, 13).value * xlWorkSheet2.Cells(35, 12).value / 1000) + (pa * xlWorkSheet2.Cells(34, 10).value * bver1 / 1000) + (pgr * bver1 * xlWorkSheet2.Cells(34, 13).value * xlWorkSheet2.Cells(34, 12).value / 1000) + (pa * xlWorkSheet2.Cells(35, 10).value * bver2 * 3 / 1000) + (pgr * bver2 * 3 * xlWorkSheet2.Cells(35, 13).value * xlWorkSheet2.Cells(35, 12).value / 1000) + (pa * xlWorkSheet2.Cells(34, 10).value * bver2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(34, 13).value * xlWorkSheet2.Cells(34, 12).value / 1000) + (pa * xlWorkSheet2.Cells(33, 10).value * b3 * bu3 * 1.25 * cr1 / 1000) + (pgr * b3 * bu3 * 1.25 * xlWorkSheet2.Cells(33, 13).value * xlWorkSheet2.Cells(33, 12).value / 1000) + (pa * xlWorkSheet2.Cells(32, 10).value * b6 * bu6 * 1.25 * cr2 / 1000) + (pgr * b6 * bu6 * 1.25 * xlWorkSheet2.Cells(32, 13).value * xlWorkSheet2.Cells(32, 12).value / 1000) + (pa * xlWorkSheet2.Cells(33, 10).value * b7 * bu7 * 1.25 * cr3 / 1000) + (pgr * b7 * bu7 * 1.25 * xlWorkSheet2.Cells(33, 13).value * xlWorkSheet2.Cells(33, 12).value / 1000) + (pg * xlWorkSheet2.Cells(36, 10).value * wh * 1.25 * cr4 / 1000) + (pgr * wh * 1.25 * xlWorkSheet2.Cells(36, 13).value * xlWorkSheet2.Cells(36, 12).value / 1000) + (pa * xlWorkSheet2.Cells(86, 10).value * b11 * 1.25 * cpl / 1000) + (pgr * b11 * 1.25 * xlWorkSheet2.Cells(86, 13).value * xlWorkSheet2.Cells(86, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Full Glass TES con 4HM ************************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 6) And (cmbano.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (b3 * bu3 < 4500) And (b3 * bu3 > 3000) Then
                        cr1 = 1.4
                    ElseIf (b3 * bu3 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (b6 * bu6 < 4500) And (b6 * bu6 > 3000) Then
                        cr2 = 1.4
                    ElseIf (b6 * bu6 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (b7 * bu7 < 4500) And (b7 * bu7 > 3000) Then
                        cr3 = 1.4
                    ElseIf (b7 * bu7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (wh < 4500) And (wh > 3000) Then
                        cr4 = 1.4
                    ElseIf (wh < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (b11 < 4500) And (b11 > 3000) Then
                        cpl = 1.4
                    ElseIf (b11 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(35, 10).value * bver2 * 3 / 1000) + (pgr * bver2 * 3 * xlWorkSheet2.Cells(35, 13).value * xlWorkSheet2.Cells(35, 12).value / 1000) + (pa * xlWorkSheet2.Cells(34, 10).value * bver2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(34, 13).value * xlWorkSheet2.Cells(34, 12).value / 1000) + (pa * xlWorkSheet2.Cells(32, 10).value * b6 * bu6 * 1.25 * cr2 / 1000) + (pgr * b6 * bu6 * 1.25 * xlWorkSheet2.Cells(32, 13).value * xlWorkSheet2.Cells(32, 12).value / 1000) + (pa * xlWorkSheet2.Cells(33, 10).value * b7 * bu7 * 1.25 * cr3 / 1000) + (pgr * b7 * bu7 * 1.25 * xlWorkSheet2.Cells(33, 13).value * xlWorkSheet2.Cells(33, 12).value / 1000) + (pg * xlWorkSheet2.Cells(36, 10).value * wh * 1.25 * cr4 / 1000) + (pgr * wh * 1.25 * xlWorkSheet2.Cells(36, 13).value * xlWorkSheet2.Cells(36, 12).value / 1000) + (pa * xlWorkSheet2.Cells(86, 10).value * b11 * 1.25 * cpl / 1000) + (pgr * b11 * 1.25 * xlWorkSheet2.Cells(86, 13).value * xlWorkSheet2.Cells(86, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Full Glass TES con 1HF + 2HM ********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 7) And (cmbano.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (b3 * bu3 < 4500) And (b3 * bu3 > 3000) Then
                        cr1 = 1.4
                    ElseIf (b3 * bu3 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (b6 * bu6 < 4500) And (b6 * bu6 > 3000) Then
                        cr2 = 1.4
                    ElseIf (b6 * bu6 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (b7 * bu7 < 4500) And (b7 * bu7 > 3000) Then
                        cr3 = 1.4
                    ElseIf (b7 * bu7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (wh < 4500) And (wh > 3000) Then
                        cr4 = 1.4
                    ElseIf (wh < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (b11 < 4500) And (b11 > 3000) Then
                        cpl = 1.4
                    ElseIf (b11 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(35, 10).value * bver1 * 0.5 / 1000) + (pgr * bver1 * 0.5 * xlWorkSheet2.Cells(35, 13).value * xlWorkSheet2.Cells(35, 12).value / 1000) + (pa * xlWorkSheet2.Cells(34, 10).value * bver1 * 0.5 / 1000) + (pgr * bver1 * 0.5 * xlWorkSheet2.Cells(34, 13).value * xlWorkSheet2.Cells(34, 12).value / 1000) + (pa * xlWorkSheet2.Cells(35, 10).value * bver2 * 1.5 / 1000) + (pgr * bver2 * 1.5 * xlWorkSheet2.Cells(35, 13).value * xlWorkSheet2.Cells(35, 12).value / 1000) + (pa * xlWorkSheet2.Cells(34, 10).value * bver2 * 0.5 / 1000) + (pgr * bver2 * 0.5 * xlWorkSheet2.Cells(34, 13).value * xlWorkSheet2.Cells(34, 12).value / 1000) + (pa * xlWorkSheet2.Cells(33, 10).value * b3 * bu3 * 1.25 * cr1 / 1000) + (pgr * b3 * bu3 * 1.25 * xlWorkSheet2.Cells(33, 13).value * xlWorkSheet2.Cells(33, 12).value / 1000) + (pa * xlWorkSheet2.Cells(32, 10).value * b6 * bu6 * 1.25 * cr2 / 1000) + (pgr * b6 * bu6 * 1.25 * xlWorkSheet2.Cells(32, 13).value * xlWorkSheet2.Cells(32, 12).value / 1000) + (pa * xlWorkSheet2.Cells(33, 10).value * b7 * bu7 * 1.25 * cr3 / 1000) + (pgr * b7 * bu7 * 1.25 * xlWorkSheet2.Cells(33, 13).value * xlWorkSheet2.Cells(33, 12).value / 1000) + (pg * xlWorkSheet2.Cells(36, 10).value * wh * 1.25 * cr4 / 1000) + (pgr * wh * 1.25 * xlWorkSheet2.Cells(36, 13).value * xlWorkSheet2.Cells(36, 12).value / 1000) + (pa * xlWorkSheet2.Cells(86, 10).value * b11 * 1.25 * cpl / 1000) + (pgr * b11 * 1.25 * xlWorkSheet2.Cells(86, 13).value * xlWorkSheet2.Cells(86, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Full Glass TES con 2HM TES *********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 2) And (cmbconf.SelectedIndex = 8) And (cmbmodcar.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double

                    If (b1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (b4 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (b3 * bu3 < 4500) And (b3 * bu3 > 3000) Then
                        cr1 = 1.4
                    ElseIf (b3 * bu3 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (b6 * bu6 < 4500) And (b6 * bu6 > 3000) Then
                        cr2 = 1.4
                    ElseIf (b6 * bu6 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (b7 * bu7 < 4500) And (b7 * bu7 > 3000) Then
                        cr3 = 1.4
                    ElseIf (b7 * bu7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (wh < 4500) And (wh > 3000) Then
                        cr4 = 1.4
                    ElseIf (wh < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (b11 < 4500) And (b11 > 3000) Then
                        cpl = 1.4
                    ElseIf (b11 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(35, 10).value * bver2 * 1.5 / 1000) + (pgr * bver2 * 1.5 * xlWorkSheet2.Cells(35, 13).value * xlWorkSheet2.Cells(35, 12).value / 1000) + (pa * xlWorkSheet2.Cells(34, 10).value * bver2 * 0.5 / 1000) + (pgr * bver2 * 0.5 * xlWorkSheet2.Cells(34, 13).value * xlWorkSheet2.Cells(34, 12).value / 1000) + (pa * xlWorkSheet2.Cells(32, 10).value * b6 * bu6 * 1.25 * cr2 / 1000) + (pgr * b6 * bu6 * 1.25 * xlWorkSheet2.Cells(32, 13).value * xlWorkSheet2.Cells(32, 12).value / 1000) + (pa * xlWorkSheet2.Cells(33, 10).value * b7 * bu7 * 1.25 * cr3 / 1000) + (pgr * b7 * bu7 * 1.25 * xlWorkSheet2.Cells(33, 13).value * xlWorkSheet2.Cells(33, 12).value / 1000) + (pg * xlWorkSheet2.Cells(36, 10).value * wh * 1.25 * cr4 / 1000) + (pgr * wh * 1.25 * xlWorkSheet2.Cells(36, 13).value * xlWorkSheet2.Cells(36, 12).value / 1000) + (pa * xlWorkSheet2.Cells(86, 10).value * b11 * 1.25 * cpl / 1000) + (pgr * b11 * 1.25 * xlWorkSheet2.Cells(86, 13).value * xlWorkSheet2.Cells(86, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                End If

            Catch ex As Exception

            End Try

            Try

                ' *********************** Carpintería Plintos silicona superior e inferior con 2HF + 2HM *************************************************************************** 

                If (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 1 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pgr * c3 * cu3 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Plintos silicona superior e inferior con 2HM ******************************************************************************

                ElseIf (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 2 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pgr * c3 * cu3 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 4).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Plintos silicona superior e inferior con 1HF + 1HM ***************************************************************************

                ElseIf (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 3 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pgr * c3 * cu3 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Plintos silicona superior e inferior con 1HM ******************************************

                ElseIf (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 4 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pgr * c3 * cu3 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Plintos silicona superior e inferior TES con 2HF + 4HM *********************************************************************************

                ElseIf (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 5 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pgr * c3 * cu3 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(86, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(86, 13).value * xlWorkSheet2.Cells(86, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Plintos silicona superior e inferior TES con 4HM **************************************************

                ElseIf (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 6 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pgr * c3 * cu3 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(86, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(86, 13).value * xlWorkSheet2.Cells(86, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Plintos silicona superior e inferior TES con 1HF + 2HM *****************************************

                ElseIf (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 7 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pgr * c3 * cu3 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(86, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(86, 13).value * xlWorkSheet2.Cells(86, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Plintos silicona superior e inferior TES con 2HM *****************

                ElseIf (cmbmodcar.SelectedIndex = 3 And cmbconf.SelectedIndex = 8 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pgr * c3 * cu3 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(86, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(86, 13).value * xlWorkSheet2.Cells(86, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                End If

            Catch ex As Exception

            End Try

            Try

                ' *********************** Carpintería Plintos silicona solo superior con 2HF + 2HM **************************************************** 

                If (cmbmodcar.SelectedIndex = 4 And cmbconf.SelectedIndex = 1 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Plintos silicona solo superior con 2HM **************************************************

                ElseIf (cmbmodcar.SelectedIndex = 4 And cmbconf.SelectedIndex = 2 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Plintos silicona solo superior con 1HF + 1HM **************************************************

                ElseIf (cmbmodcar.SelectedIndex = 4 And cmbconf.SelectedIndex = 3 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Plintos silicona solo superior con 1HM ******************************************

                ElseIf (cmbmodcar.SelectedIndex = 4 And cmbconf.SelectedIndex = 4 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                End If

            Catch ex As Exception

            End Try

            Try

                ' *********************** Carpintería Plintos superior de pinza e inferior de silicona con 2HF + 2HM *************************************************************************** 

                If (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 1 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pgr * c3 * cu3 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Plintos  superior de pinza e inferior de silicona con 2HM ******************************************************************************

                ElseIf (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 2 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pgr * c3 * cu3 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 4).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Plintos  superior de pinza e inferior de silicona con 1HF + 1HM ***************************************************************************

                ElseIf (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 3 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pgr * c3 * cu3 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Plintos  superior de pinza e inferior de silicona con 1HM ******************************************

                ElseIf (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 4 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pgr * c3 * cu3 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Plintos  superior de pinza e inferior de silicona TES con 2HF + 4HM *********************************************************************************

                ElseIf (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 5 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pgr * c3 * cu3 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(86, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(86, 13).value * xlWorkSheet2.Cells(86, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Plintos  superior de pinza e inferior de silicona TES con 4HM **************************************************

                ElseIf (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 6 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pgr * c3 * cu3 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(86, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(86, 13).value * xlWorkSheet2.Cells(86, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Plintos  superior de pinza e inferior de silicona TES con 1HF + 2HM *****************************************

                ElseIf (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 7 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pgr * c3 * cu3 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(86, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(86, 13).value * xlWorkSheet2.Cells(86, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Plintos  superior de pinza e inferior de silicona TES con 2HM *****************

                ElseIf (cmbmodcar.SelectedIndex = 5 And cmbconf.SelectedIndex = 8 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(47, 10).value * c3 * cu3 * 1.25 * cr3 / 1000) + (pgr * c3 * cu3 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(86, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(86, 13).value * xlWorkSheet2.Cells(86, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                End If

            Catch ex As Exception

            End Try

            Try

                ' *********************** Carpintería Plintos silicona solo superior  con 2HF + 2HM **************************************************** 

                If (cmbmodcar.SelectedIndex = 6 And cmbconf.SelectedIndex = 1 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Plintos silicona solo superior con 2HM **************************************************

                ElseIf (cmbmodcar.SelectedIndex = 6 And cmbconf.SelectedIndex = 2 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Plintos silicona solo superior con 1HF + 1HM **************************************************

                ElseIf (cmbmodcar.SelectedIndex = 6 And cmbconf.SelectedIndex = 3 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería Plintos silicona solo superior con 1HM ******************************************

                ElseIf (cmbmodcar.SelectedIndex = 6 And cmbconf.SelectedIndex = 4 And cmbano.SelectedValue = codleido.ToString) Then

                    Dim importe As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double

                    If (c1 * cu1 < 4500) And (c1 * cu1 > 3000) Then
                        cr1 = 1.4
                    ElseIf (c1 * cu1 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (c2 * cu2 < 4500) And (c2 * cu2 > 3000) Then
                        cr2 = 1.4
                    ElseIf (c2 * cu2 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (c3 * cu3 < 4500) And (c3 * cu3 > 3000) Then
                        cr3 = 1.2
                    ElseIf (c3 * cu3 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (c6 < 4500) And (c6 > 3000) Then
                        cpl = 1.4
                    ElseIf (c6 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(47, 10).value * c1 * cu1 * 1.25 * cr1 / 1000) + (pgr * c1 * cu1 * 1.25 * xlWorkSheet2.Cells(47, 13).value * xlWorkSheet2.Cells(47, 12).value / 1000) + (pa * xlWorkSheet2.Cells(46, 10).value * c2 * cu2 * 1.25 * cr2 / 1000) + (pgr * c2 * cu2 * 1.25 * xlWorkSheet2.Cells(46, 13).value * xlWorkSheet2.Cells(46, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * c6 * 1.25 * cpl / 1000) + (pgr * c6 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                End If

            Catch ex As Exception

            End Try


            Try

                ' *********************** Carpintería BW52 con 2HF + 2HM *************************************************************************** 

                If ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 1) And (cmbano.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double
                    Dim cr6 As Double
                    Dim cr7 As Double
                    Dim cr8 As Double
                    Dim cr9 As Double

                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 4700
                    End If

                    If (d4 * du4 < 4500) And (d4 * du4 > 3000) Then
                        cr1 = 1.4
                    ElseIf (d4 * du4 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (d5 * du5 < 4500) And (d5 * du5 > 3000) Then
                        cr2 = 1.2
                    ElseIf (d5 * du5 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (d7 * du7 < 4500) And (d7 * du7 > 3000) Then
                        cr3 = 1.2
                    ElseIf (d7 * du7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (d8 * du8 < 4500) And (d8 * du8 > 3000) Then
                        cr4 = 1.2
                    ElseIf (d8 * du8 < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (d20 * du20 < 4500) And (d20 * du20 > 3000) Then
                        cr5 = 1.2
                    ElseIf (d20 * du20 < 3000) Then
                        cr5 = 1.8
                    Else : cr5 = 1.0
                    End If

                    If (d13 * du13 < 4500) And (d13 * du13 > 3000) Then
                        cr6 = 1.2
                    ElseIf (d13 * du13 < 3000) Then
                        cr6 = 1.8
                    Else : cr6 = 1.0
                    End If

                    If (d14 * du14 < 4500) And (d14 * du14 > 3000) Then
                        cr7 = 1.2
                    ElseIf (d14 * du14 < 3000) Then
                        cr7 = 1.8
                    Else : cr7 = 1.0
                    End If

                    If (d17 * du17 < 4500) And (d17 * du17 > 3000) Then
                        cr8 = 1.2
                    ElseIf (d17 * du17 < 3000) Then
                        cr8 = 1.8
                    Else : cr8 = 1.0
                    End If

                    If (d18 * du18 < 4700) And (d18 * du18 > 3000) Then
                        cr9 = 1.2
                    ElseIf (d18 * du18 < 3000) Then
                        cr9 = 1.8
                    Else : cr9 = 1.0
                    End If

                    If (d22 < 4500) And (d22 > 3000) Then
                        cpl = 1.4
                    ElseIf (d22 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(67, 10).value * d4 * du4 * 1.25 * cr1 / 1000) + (pgr * d4 * du4 * 1.25 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(67, 10).value * bver1 * 2 / 1000) + (pgr * bver1 * 2 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(68, 10).value * d5 * du5 * 1.25 * cr2 / 1000) + (pgr * d5 * du5 * 1.25 * xlWorkSheet2.Cells(68, 13).value * xlWorkSheet2.Cells(68, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * bver3 * 2 / 1000) + (pgr * bver3 * 2 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * d7 * du7 * 1.25 * cr3 / 1000) + (pgr * d7 * du7 * 1.25 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(69, 10).value * bver1 / 1000) + (pgr * bver1 * xlWorkSheet2.Cells(69, 13).value * xlWorkSheet2.Cells(69, 12).value / 1000) + (pa * xlWorkSheet2.Cells(70, 10).value * bver1 / 1000) + (pgr * bver1 * xlWorkSheet2.Cells(70, 13).value * xlWorkSheet2.Cells(70, 12).value / 1000) + (pa * xlWorkSheet2.Cells(73, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pgr * d8 * du8 * 1.25 * xlWorkSheet2.Cells(73, 13).value * xlWorkSheet2.Cells(73, 12).value / 1000) + (pa * xlWorkSheet2.Cells(72, 10).value * d20 * du20 * 1.25 * cr5 / 1000) + (pgr * d20 * du20 * 1.25 * xlWorkSheet2.Cells(72, 13).value * xlWorkSheet2.Cells(72, 12).value / 1000) + (pa * xlWorkSheet2.Cells(67, 10).value * d13 * du13 * 1.25 * cr6 / 1000) + (pgr * d13 * du13 * 1.25 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(67, 10).value * bver2 * 2 / 1000) + (pgr * bver2 * 2 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(68, 10).value * d14 * du14 * 1.25 * cr7 / 1000) + (pgr * d14 * du14 * 1.25 * xlWorkSheet2.Cells(68, 13).value * xlWorkSheet2.Cells(68, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * bver4 * 2 / 1000) + (pgr * bver4 * 2 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * d17 * du17 * 1.25 * cr8 / 1000) + (pgr * d17 * du17 * 1.25 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(74, 10).value * d18 * du18 * 1.25 * cr9 / 1000) + (pgr * d18 * du18 * 1.25 * xlWorkSheet2.Cells(74, 13).value * xlWorkSheet2.Cells(74, 12).value / 1000) + (pa * xlWorkSheet2.Cells(78, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pgr * d8 * du8 * 1.25 * xlWorkSheet2.Cells(78, 13).value * xlWorkSheet2.Cells(78, 12).value / 1000) + (pa * xlWorkSheet2.Cells(69, 10).value * bver2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(69, 13).value * xlWorkSheet2.Cells(69, 12).value / 1000) + (pa * xlWorkSheet2.Cells(71, 10).value * bver2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(71, 13).value * xlWorkSheet2.Cells(71, 12).value / 1000) + (pa * xlWorkSheet2.Cells(70, 10).value * bver2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(70, 13).value * xlWorkSheet2.Cells(70, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * d22 * 1.25 * cpl / 1000) + (pgr * d22 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería BW52 con 2HM ******************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 2) And (cmbano.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double
                    Dim cr6 As Double
                    Dim cr7 As Double
                    Dim cr8 As Double
                    Dim cr9 As Double

                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 4700
                    End If

                    If (d4 * du4 < 4500) And (d4 * du4 > 3000) Then
                        cr1 = 1.4
                    ElseIf (d4 * du4 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (d5 * du5 < 4500) And (d5 * du5 > 3000) Then
                        cr2 = 1.2
                    ElseIf (d5 * du5 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (d7 * du7 < 4500) And (d7 * du7 > 3000) Then
                        cr3 = 1.2
                    ElseIf (d7 * du7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (d8 * du8 < 4500) And (d8 * du8 > 3000) Then
                        cr4 = 1.2
                    ElseIf (d8 * du8 < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (d20 * du20 < 4500) And (d20 * du20 > 3000) Then
                        cr5 = 1.2
                    ElseIf (d20 * du20 < 3000) Then
                        cr5 = 1.8
                    Else : cr5 = 1.0
                    End If

                    If (d13 * du13 < 4500) And (d13 * du13 > 3000) Then
                        cr6 = 1.2
                    ElseIf (d13 * du13 < 3000) Then
                        cr6 = 1.8
                    Else : cr6 = 1.0
                    End If

                    If (d14 * du14 < 4500) And (d14 * du14 > 3000) Then
                        cr7 = 1.2
                    ElseIf (d14 * du14 < 3000) Then
                        cr7 = 1.8
                    Else : cr7 = 1.0
                    End If

                    If (d17 * du17 < 4500) And (d17 * du17 > 3000) Then
                        cr8 = 1.2
                    ElseIf (d17 * du17 < 3000) Then
                        cr8 = 1.8
                    Else : cr8 = 1.0
                    End If

                    If (d18 * du18 < 4700) And (d18 * du18 > 3000) Then
                        cr9 = 1.2
                    ElseIf (d18 * du18 < 3000) Then
                        cr9 = 1.8
                    Else : cr9 = 1.0
                    End If

                    If (d22 < 4500) And (d22 > 3000) Then
                        cpl = 1.4
                    ElseIf (d22 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(67, 10).value * d13 * du13 * 1.25 * cr6 / 1000) + (pgr * d13 * du13 * 1.25 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(67, 10).value * bver2 * 2 / 1000) + (pgr * bver2 * 2 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(68, 10).value * d14 * du14 * 1.25 * cr7 / 1000) + (pgr * d14 * du14 * 1.25 * xlWorkSheet2.Cells(68, 13).value * xlWorkSheet2.Cells(68, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * bver4 * 2 / 1000) + (pgr * bver4 * 2 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * d17 * du17 * 1.25 * cr8 / 1000) + (pgr * d17 * du17 * 1.25 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(74, 10).value * d18 * du18 * 1.25 * cr9 / 1000) + (pgr * d18 * du18 * 1.25 * xlWorkSheet2.Cells(74, 13).value * xlWorkSheet2.Cells(74, 12).value / 1000) + (pa * xlWorkSheet2.Cells(78, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pgr * d8 * du8 * 1.25 * xlWorkSheet2.Cells(78, 13).value * xlWorkSheet2.Cells(78, 12).value / 1000) + (pa * xlWorkSheet2.Cells(69, 10).value * bver2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(69, 13).value * xlWorkSheet2.Cells(69, 12).value / 1000) + (pa * xlWorkSheet2.Cells(71, 10).value * bver2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(71, 13).value * xlWorkSheet2.Cells(71, 12).value / 1000) + (pa * xlWorkSheet2.Cells(70, 10).value * bver2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(70, 13).value * xlWorkSheet2.Cells(70, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * d22 * 1.25 * cpl / 1000) + (pgr * d22 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True


                    ' *********************** Carpintería BW52 con 1HF + 1HM ***************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 3) And (cmbano.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double
                    Dim cr6 As Double
                    Dim cr7 As Double
                    Dim cr8 As Double
                    Dim cr9 As Double

                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 4700
                    End If

                    If (d4 * du4 < 4500) And (d4 * du4 > 3000) Then
                        cr1 = 1.4
                    ElseIf (d4 * du4 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (d5 * du5 < 4500) And (d5 * du5 > 3000) Then
                        cr2 = 1.2
                    ElseIf (d5 * du5 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (d7 * du7 < 4500) And (d7 * du7 > 3000) Then
                        cr3 = 1.2
                    ElseIf (d7 * du7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (d8 * du8 < 4500) And (d8 * du8 > 3000) Then
                        cr4 = 1.2
                    ElseIf (d8 * du8 < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (d20 * du20 < 4500) And (d20 * du20 > 3000) Then
                        cr5 = 1.2
                    ElseIf (d20 * du20 < 3000) Then
                        cr5 = 1.8
                    Else : cr5 = 1.0
                    End If

                    If (d13 * du13 < 4500) And (d13 * du13 > 3000) Then
                        cr6 = 1.2
                    ElseIf (d13 * du13 < 3000) Then
                        cr6 = 1.8
                    Else : cr6 = 1.0
                    End If

                    If (d14 * du14 < 4500) And (d14 * du14 > 3000) Then
                        cr7 = 1.2
                    ElseIf (d14 * du14 < 3000) Then
                        cr7 = 1.8
                    Else : cr7 = 1.0
                    End If

                    If (d17 * du17 < 4500) And (d17 * du17 > 3000) Then
                        cr8 = 1.2
                    ElseIf (d17 * du17 < 3000) Then
                        cr8 = 1.8
                    Else : cr8 = 1.0
                    End If

                    If (d18 * du18 < 4700) And (d18 * du18 > 3000) Then
                        cr9 = 1.2
                    ElseIf (d18 * du18 < 3000) Then
                        cr9 = 1.8
                    Else : cr9 = 1.0
                    End If

                    If (d22 < 4500) And (d22 > 3000) Then
                        cpl = 1.4
                    ElseIf (d22 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(67, 10).value * d4 * du4 * 1.25 * cr1 / 1000) + (pgr * d4 * du4 * 1.25 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(67, 10).value * bver1 * 1 / 1000) + (pgr * bver1 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(68, 10).value * d5 * du5 * 1.25 * cr2 / 1000) + (pgr * d5 * du5 * 1.25 * xlWorkSheet2.Cells(68, 13).value * xlWorkSheet2.Cells(68, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * bver3 * 1 / 1000) + (pgr * bver3 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * d7 * du7 * 1.25 * cr3 / 1000) + (pgr * d7 * du7 * 1.25 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(69, 10).value * bver1 * 0.5 / 1000) + (pgr * bver1 * 0.5 * xlWorkSheet2.Cells(69, 13).value * xlWorkSheet2.Cells(69, 12).value / 1000) + (pa * xlWorkSheet2.Cells(70, 10).value * bver1 * 0.5 / 1000) + (pgr * bver1 * 0.5 * xlWorkSheet2.Cells(70, 13).value * xlWorkSheet2.Cells(70, 12).value / 1000) + (pa * xlWorkSheet2.Cells(73, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pgr * d8 * du8 * 1.25 * xlWorkSheet2.Cells(73, 13).value * xlWorkSheet2.Cells(73, 12).value / 1000) + (pa * xlWorkSheet2.Cells(72, 10).value * d20 * du20 * 1.25 * cr5 / 1000) + (pgr * d20 * du20 * 1.25 * xlWorkSheet2.Cells(72, 13).value * xlWorkSheet2.Cells(72, 12).value / 1000) + (pa * xlWorkSheet2.Cells(67, 10).value * d13 * du13 * 1.25 * cr6 / 1000) + (pgr * d13 * du13 * 1.25 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(67, 10).value * bver2 * 1 / 1000) + (pgr * bver2 * 1 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(68, 10).value * d14 * du14 * 1.25 * cr7 / 1000) + (pgr * d14 * du14 * 1.25 * xlWorkSheet2.Cells(68, 13).value * xlWorkSheet2.Cells(68, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * bver4 * 1 / 1000) + (pgr * bver4 * 1 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * d17 * du17 * 1.25 * cr8 / 1000) + (pgr * d17 * du17 * 1.25 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(74, 10).value * d18 * du18 * 1.25 * cr9 / 1000) + (pgr * d18 * du18 * 1.25 * xlWorkSheet2.Cells(74, 13).value * xlWorkSheet2.Cells(74, 12).value / 1000) + (pa * xlWorkSheet2.Cells(78, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pgr * d8 * du8 * 1.25 * xlWorkSheet2.Cells(78, 13).value * xlWorkSheet2.Cells(78, 12).value / 1000) + (pa * xlWorkSheet2.Cells(69, 10).value * bver2 * 0.5 / 1000) + (pgr * bver2 * 0.5 * xlWorkSheet2.Cells(69, 13).value * xlWorkSheet2.Cells(69, 12).value / 1000) + (pa * xlWorkSheet2.Cells(71, 10).value * bver2 * 0.5 / 1000) + (pgr * bver2 * 0.5 * xlWorkSheet2.Cells(71, 13).value * xlWorkSheet2.Cells(71, 12).value / 1000) + (pa * xlWorkSheet2.Cells(70, 10).value * bver2 * 0.5 / 1000) + (pgr * bver2 * 0.5 * xlWorkSheet2.Cells(70, 13).value * xlWorkSheet2.Cells(70, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * d22 * 1.25 * cpl / 1000) + (pgr * d22 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True


                    ' *********************** Carpintería BW52 con 1HM *********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 4) And (cmbano.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double
                    Dim cr6 As Double
                    Dim cr7 As Double
                    Dim cr8 As Double
                    Dim cr9 As Double

                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 4700
                    End If

                    If (d4 * du4 < 4500) And (d4 * du4 > 3000) Then
                        cr1 = 1.4
                    ElseIf (d4 * du4 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (d5 * du5 < 4500) And (d5 * du5 > 3000) Then
                        cr2 = 1.2
                    ElseIf (d5 * du5 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (d7 * du7 < 4500) And (d7 * du7 > 3000) Then
                        cr3 = 1.2
                    ElseIf (d7 * du7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (d8 * du8 < 4500) And (d8 * du8 > 3000) Then
                        cr4 = 1.2
                    ElseIf (d8 * du8 < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (d20 * du20 < 4500) And (d20 * du20 > 3000) Then
                        cr5 = 1.2
                    ElseIf (d20 * du20 < 3000) Then
                        cr5 = 1.8
                    Else : cr5 = 1.0
                    End If

                    If (d13 * du13 < 4500) And (d13 * du13 > 3000) Then
                        cr6 = 1.2
                    ElseIf (d13 * du13 < 3000) Then
                        cr6 = 1.8
                    Else : cr6 = 1.0
                    End If

                    If (d14 * du14 < 4500) And (d14 * du14 > 3000) Then
                        cr7 = 1.2
                    ElseIf (d14 * du14 < 3000) Then
                        cr7 = 1.8
                    Else : cr7 = 1.0
                    End If

                    If (d17 * du17 < 4500) And (d17 * du17 > 3000) Then
                        cr8 = 1.2
                    ElseIf (d17 * du17 < 3000) Then
                        cr8 = 1.8
                    Else : cr8 = 1.0
                    End If

                    If (d18 * du18 < 4700) And (d18 * du18 > 3000) Then
                        cr9 = 1.2
                    ElseIf (d18 * du18 < 3000) Then
                        cr9 = 1.8
                    Else : cr9 = 1.0
                    End If

                    If (d22 < 4500) And (d22 > 3000) Then
                        cpl = 1.4
                    ElseIf (d22 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(67, 10).value * d13 * du13 * 1.25 * cr6 / 1000) + (pgr * d13 * du13 * 1.25 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(67, 10).value * bver2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(68, 10).value * d14 * du14 * 1.25 * cr7 / 1000) + (pgr * d14 * du14 * 1.25 * xlWorkSheet2.Cells(68, 13).value * xlWorkSheet2.Cells(68, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * bver4 / 1000) + (pgr * bver4 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * d17 * du17 * 1.25 * cr8 / 1000) + (pgr * d17 * du17 * 1.25 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(74, 10).value * d18 * du18 * 1.25 * cr9 / 1000) + (pgr * d18 * du18 * 1.25 * xlWorkSheet2.Cells(74, 13).value * xlWorkSheet2.Cells(74, 12).value / 1000) + (pa * xlWorkSheet2.Cells(78, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pgr * d8 * du8 * 1.25 * xlWorkSheet2.Cells(78, 13).value * xlWorkSheet2.Cells(78, 12).value / 1000) + (pa * xlWorkSheet2.Cells(69, 10).value * bver2 * 0.5 / 1000) + (pgr * bver2 * 0.5 * xlWorkSheet2.Cells(69, 13).value * xlWorkSheet2.Cells(69, 12).value / 1000) + (pa * xlWorkSheet2.Cells(71, 10).value * bver2 * 0.5 / 1000) + (pgr * bver2 * 0.5 * xlWorkSheet2.Cells(71, 13).value * xlWorkSheet2.Cells(71, 12).value / 1000) + (pa * xlWorkSheet2.Cells(70, 10).value * bver2 * 0.5 / 1000) + (pgr * bver2 * 0.5 * xlWorkSheet2.Cells(70, 13).value * xlWorkSheet2.Cells(70, 12).value / 1000) + (pa * xlWorkSheet2.Cells(85, 10).value * d22 * 1.25 * cpl / 1000) + (pgr * d22 * 1.25 * xlWorkSheet2.Cells(85, 13).value * xlWorkSheet2.Cells(85, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería BW52 TES con 2HF + 4HM *******************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 5) And (cmbano.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double
                    Dim cr6 As Double
                    Dim cr7 As Double
                    Dim cr8 As Double
                    Dim cr9 As Double

                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 4700
                    End If

                    If (d4 * du4 < 4500) And (d4 * du4 > 3000) Then
                        cr1 = 1.4
                    ElseIf (d4 * du4 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (d5 * du5 < 4500) And (d5 * du5 > 3000) Then
                        cr2 = 1.2
                    ElseIf (d5 * du5 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (d7 * du7 < 4500) And (d7 * du7 > 3000) Then
                        cr3 = 1.2
                    ElseIf (d7 * du7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (d8 * du8 < 4500) And (d8 * du8 > 3000) Then
                        cr4 = 1.2
                    ElseIf (d8 * du8 < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (d20 * du20 < 4500) And (d20 * du20 > 3000) Then
                        cr5 = 1.2
                    ElseIf (d20 * du20 < 3000) Then
                        cr5 = 1.8
                    Else : cr5 = 1.0
                    End If

                    If (d13 * du13 < 4500) And (d13 * du13 > 3000) Then
                        cr6 = 1.2
                    ElseIf (d13 * du13 < 3000) Then
                        cr6 = 1.8
                    Else : cr6 = 1.0
                    End If

                    If (d14 * du14 < 4500) And (d14 * du14 > 3000) Then
                        cr7 = 1.2
                    ElseIf (d14 * du14 < 3000) Then
                        cr7 = 1.8
                    Else : cr7 = 1.0
                    End If

                    If (d17 * du17 < 4500) And (d17 * du17 > 3000) Then
                        cr8 = 1.2
                    ElseIf (d17 * du17 < 3000) Then
                        cr8 = 1.8
                    Else : cr8 = 1.0
                    End If

                    If (d18 * du18 < 4700) And (d18 * du18 > 3000) Then
                        cr9 = 1.2
                    ElseIf (d18 * du18 < 3000) Then
                        cr9 = 1.8
                    Else : cr9 = 1.0
                    End If

                    If (d22 < 4500) And (d22 > 3000) Then
                        cpl = 1.4
                    ElseIf (d22 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(67, 10).value * d4 * du4 * 1.25 * cr1 / 1000) + (pgr * d4 * du4 * 1.25 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(67, 10).value * bver1 * 2 / 1000) + (pgr * bver1 * 2 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(68, 10).value * d5 * du5 * 1.25 * cr2 / 1000) + (pgr * d5 * du5 * 1.25 * xlWorkSheet2.Cells(68, 13).value * xlWorkSheet2.Cells(68, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * bver3 * 2 / 1000) + (pgr * bver3 * 2 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * d7 * du7 * 1.25 * cr3 / 1000) + (pgr * d7 * du7 * 1.25 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(69, 10).value * bver1 / 1000) + (pgr * bver1 * xlWorkSheet2.Cells(69, 13).value * xlWorkSheet2.Cells(69, 12).value / 1000) + (pa * xlWorkSheet2.Cells(70, 10).value * bver1 / 1000) + (pgr * bver1 * xlWorkSheet2.Cells(70, 13).value * xlWorkSheet2.Cells(70, 12).value / 1000) + (pa * xlWorkSheet2.Cells(73, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pgr * d8 * du8 * 1.25 * xlWorkSheet2.Cells(73, 13).value * xlWorkSheet2.Cells(73, 12).value / 1000) + (pa * xlWorkSheet2.Cells(72, 10).value * d20 * du20 * 1.25 * cr5 / 1000) + (pgr * d20 * du20 * 1.25 * xlWorkSheet2.Cells(72, 13).value * xlWorkSheet2.Cells(72, 12).value / 1000) + (pa * xlWorkSheet2.Cells(67, 10).value * d13 * du13 * 1.25 * cr6 / 1000) + (pgr * d13 * du13 * 1.25 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(67, 10).value * bver2 * 4 / 1000) + (pgr * bver2 * 4 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(68, 10).value * d14 * du14 * 1.25 * cr7 / 1000) + (pgr * d14 * du14 * 1.25 * xlWorkSheet2.Cells(68, 13).value * xlWorkSheet2.Cells(68, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * bver4 * 4 / 1000) + (pgr * bver4 * 4 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * d17 * du17 * 1.25 * cr8 / 1000) + (pgr * d17 * du17 * 1.25 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(74, 10).value * d18 * du18 * 1.25 * cr9 / 1000) + (pgr * d18 * du18 * 1.25 * xlWorkSheet2.Cells(74, 13).value * xlWorkSheet2.Cells(74, 12).value / 1000) + (pa * xlWorkSheet2.Cells(78, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pgr * d8 * du8 * 1.25 * xlWorkSheet2.Cells(78, 13).value * xlWorkSheet2.Cells(78, 12).value / 1000) + (pa * xlWorkSheet2.Cells(69, 10).value * bver2 * 3 / 1000) + (pgr * bver2 * 3 * xlWorkSheet2.Cells(69, 13).value * xlWorkSheet2.Cells(69, 12).value / 1000) + (pa * xlWorkSheet2.Cells(71, 10).value * bver2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(71, 13).value * xlWorkSheet2.Cells(71, 12).value / 1000) + (pa * xlWorkSheet2.Cells(70, 10).value * bver2 * 2 / 1000) + (pgr * bver2 * 2 * xlWorkSheet2.Cells(70, 13).value * xlWorkSheet2.Cells(70, 12).value / 1000) + (pa * xlWorkSheet2.Cells(86, 10).value * d22 * 1.25 * cpl / 1000) + (pgr * d22 * 1.25 * xlWorkSheet2.Cells(86, 13).value * xlWorkSheet2.Cells(86, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería BW52 TES con 4HM ************************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 6) And (cmbano.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double
                    Dim cr6 As Double
                    Dim cr7 As Double
                    Dim cr8 As Double
                    Dim cr9 As Double

                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 4700
                    End If

                    If (d4 * du4 < 4500) And (d4 * du4 > 3000) Then
                        cr1 = 1.4
                    ElseIf (d4 * du4 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (d5 * du5 < 4500) And (d5 * du5 > 3000) Then
                        cr2 = 1.2
                    ElseIf (d5 * du5 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (d7 * du7 < 4500) And (d7 * du7 > 3000) Then
                        cr3 = 1.2
                    ElseIf (d7 * du7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (d8 * du8 < 4500) And (d8 * du8 > 3000) Then
                        cr4 = 1.2
                    ElseIf (d8 * du8 < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (d20 * du20 < 4500) And (d20 * du20 > 3000) Then
                        cr5 = 1.2
                    ElseIf (d20 * du20 < 3000) Then
                        cr5 = 1.8
                    Else : cr5 = 1.0
                    End If

                    If (d13 * du13 < 4500) And (d13 * du13 > 3000) Then
                        cr6 = 1.2
                    ElseIf (d13 * du13 < 3000) Then
                        cr6 = 1.8
                    Else : cr6 = 1.0
                    End If

                    If (d14 * du14 < 4500) And (d14 * du14 > 3000) Then
                        cr7 = 1.2
                    ElseIf (d14 * du14 < 3000) Then
                        cr7 = 1.8
                    Else : cr7 = 1.0
                    End If

                    If (d17 * du17 < 4500) And (d17 * du17 > 3000) Then
                        cr8 = 1.2
                    ElseIf (d17 * du17 < 3000) Then
                        cr8 = 1.8
                    Else : cr8 = 1.0
                    End If

                    If (d18 * du18 < 4700) And (d18 * du18 > 3000) Then
                        cr9 = 1.2
                    ElseIf (d18 * du18 < 3000) Then
                        cr9 = 1.8
                    Else : cr9 = 1.0
                    End If

                    If (d22 < 4500) And (d22 > 3000) Then
                        cpl = 1.4
                    ElseIf (d22 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(67, 10).value * d13 * du13 * 1.25 * cr6 / 1000) + (pgr * d13 * du13 * 1.25 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(67, 10).value * bver2 * 4 / 1000) + (pgr * bver2 * 4 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(68, 10).value * d14 * du14 * 1.25 * cr7 / 1000) + (pgr * d14 * du14 * 1.25 * xlWorkSheet2.Cells(68, 13).value * xlWorkSheet2.Cells(68, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * bver4 * 4 / 1000) + (pgr * bver4 * 4 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * d17 * du17 * 1.25 * cr8 / 1000) + (pgr * d17 * du17 * 1.25 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(74, 10).value * d18 * du18 * 1.25 * cr9 / 1000) + (pgr * d18 * du18 * 1.25 * xlWorkSheet2.Cells(74, 13).value * xlWorkSheet2.Cells(74, 12).value / 1000) + (pa * xlWorkSheet2.Cells(78, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pgr * d8 * du8 * 1.25 * xlWorkSheet2.Cells(78, 13).value * xlWorkSheet2.Cells(78, 12).value / 1000) + (pa * xlWorkSheet2.Cells(69, 10).value * bver2 * 2 / 1000) + (pgr * bver2 * 2 * xlWorkSheet2.Cells(69, 13).value * xlWorkSheet2.Cells(69, 12).value / 1000) + (pa * xlWorkSheet2.Cells(71, 10).value * bver2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(71, 13).value * xlWorkSheet2.Cells(71, 12).value / 1000) + (pa * xlWorkSheet2.Cells(70, 10).value * bver2 * 2 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(70, 13).value * xlWorkSheet2.Cells(70, 12).value / 1000) + (pa * xlWorkSheet2.Cells(86, 10).value * d22 * 1.25 * cpl / 1000) + (pgr * d22 * 1.25 * xlWorkSheet2.Cells(86, 13).value * xlWorkSheet2.Cells(86, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería BW52 TES con 1HF + 2HM ********************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 7) And (cmbano.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double
                    Dim cr6 As Double
                    Dim cr7 As Double
                    Dim cr8 As Double
                    Dim cr9 As Double

                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 4700
                    End If

                    If (d4 * du4 < 4500) And (d4 * du4 > 3000) Then
                        cr1 = 1.4
                    ElseIf (d4 * du4 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (d5 * du5 < 4500) And (d5 * du5 > 3000) Then
                        cr2 = 1.2
                    ElseIf (d5 * du5 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (d7 * du7 < 4500) And (d7 * du7 > 3000) Then
                        cr3 = 1.2
                    ElseIf (d7 * du7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (d8 * du8 < 4500) And (d8 * du8 > 3000) Then
                        cr4 = 1.2
                    ElseIf (d8 * du8 < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (d20 * du20 < 4500) And (d20 * du20 > 3000) Then
                        cr5 = 1.2
                    ElseIf (d20 * du20 < 3000) Then
                        cr5 = 1.8
                    Else : cr5 = 1.0
                    End If

                    If (d13 * du13 < 4500) And (d13 * du13 > 3000) Then
                        cr6 = 1.2
                    ElseIf (d13 * du13 < 3000) Then
                        cr6 = 1.8
                    Else : cr6 = 1.0
                    End If

                    If (d14 * du14 < 4500) And (d14 * du14 > 3000) Then
                        cr7 = 1.2
                    ElseIf (d14 * du14 < 3000) Then
                        cr7 = 1.8
                    Else : cr7 = 1.0
                    End If

                    If (d17 * du17 < 4500) And (d17 * du17 > 3000) Then
                        cr8 = 1.2
                    ElseIf (d17 * du17 < 3000) Then
                        cr8 = 1.8
                    Else : cr8 = 1.0
                    End If

                    If (d18 * du18 < 4700) And (d18 * du18 > 3000) Then
                        cr9 = 1.2
                    ElseIf (d18 * du18 < 3000) Then
                        cr9 = 1.8
                    Else : cr9 = 1.0
                    End If

                    If (d22 < 4500) And (d22 > 3000) Then
                        cpl = 1.4
                    ElseIf (d22 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(67, 10).value * d4 * du4 * 1.25 * cr1 / 1000) + (pgr * d4 * du4 * 1.25 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(67, 10).value * bver1 / 1000) + (pgr * bver1 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(68, 10).value * d5 * du5 * 1.25 * cr2 / 1000) + (pgr * d5 * du5 * 1.25 * xlWorkSheet2.Cells(68, 13).value * xlWorkSheet2.Cells(68, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * bver3 / 1000) + (pgr * bver3 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * d7 * du7 * 1.25 * cr3 / 1000) + (pgr * d7 * du7 * 1.25 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(69, 10).value * bver1 * 0.5 / 1000) + (pgr * bver1 * 0.5 * xlWorkSheet2.Cells(69, 13).value * xlWorkSheet2.Cells(69, 12).value / 1000) + (pa * xlWorkSheet2.Cells(70, 10).value * bver1 * 0.5 / 1000) + (pgr * bver1 * 0.5 * xlWorkSheet2.Cells(70, 13).value * xlWorkSheet2.Cells(70, 12).value / 1000) + (pa * xlWorkSheet2.Cells(73, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pgr * d8 * du8 * 1.25 * xlWorkSheet2.Cells(73, 13).value * xlWorkSheet2.Cells(73, 12).value / 1000) + (pa * xlWorkSheet2.Cells(72, 10).value * d20 * du20 * 1.25 * cr5 / 1000) + (pgr * d20 * du20 * 1.25 * xlWorkSheet2.Cells(72, 13).value * xlWorkSheet2.Cells(72, 12).value / 1000) + (pa * xlWorkSheet2.Cells(67, 10).value * d13 * du13 * 1.25 * cr6 / 1000) + (pgr * d13 * du13 * 1.25 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(67, 10).value * bver2 * 2 / 1000) + (pgr * bver2 * 2 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(68, 10).value * d14 * du14 * 1.25 * cr7 / 1000) + (pgr * d14 * du14 * 1.25 * xlWorkSheet2.Cells(68, 13).value * xlWorkSheet2.Cells(68, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * bver4 * 2 / 1000) + (pgr * bver4 * 2 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * d17 * du17 * 1.25 * cr8 / 1000) + (pgr * d17 * du17 * 1.25 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(74, 10).value * d18 * du18 * 1.25 * cr9 / 1000) + (pgr * d18 * du18 * 1.25 * xlWorkSheet2.Cells(74, 13).value * xlWorkSheet2.Cells(74, 12).value / 1000) + (pa * xlWorkSheet2.Cells(78, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pgr * d8 * du8 * 1.25 * xlWorkSheet2.Cells(78, 13).value * xlWorkSheet2.Cells(78, 12).value / 1000) + (pa * xlWorkSheet2.Cells(69, 10).value * bver2 * 1.5 / 1000) + (pgr * bver2 * 1.5 * xlWorkSheet2.Cells(69, 13).value * xlWorkSheet2.Cells(69, 12).value / 1000) + (pa * xlWorkSheet2.Cells(71, 10).value * bver2 * 0.5 / 1000) + (pgr * bver2 * 0.5 * xlWorkSheet2.Cells(71, 13).value * xlWorkSheet2.Cells(71, 12).value / 1000) + (pa * xlWorkSheet2.Cells(70, 10).value * bver2 * 1 / 1000) + (pgr * bver2 * 1 * xlWorkSheet2.Cells(70, 13).value * xlWorkSheet2.Cells(70, 12).value / 1000) + (pa * xlWorkSheet2.Cells(86, 10).value * d22 * 1.25 * cpl / 1000) + (pgr * d22 * 1.25 * xlWorkSheet2.Cells(86, 13).value * xlWorkSheet2.Cells(86, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                    ' *********************** Carpintería BW52 TES con 2HM ************************************************************************************

                ElseIf ((cmbmodcar.SelectedIndex = 7) And (cmbconf.SelectedIndex = 8) And (cmbano.SelectedValue = codleido.ToString)) Then

                    Dim importe As Double
                    Dim bver1 As Double
                    Dim bver2 As Double
                    Dim bver3 As Double
                    Dim bver4 As Double
                    Dim cr1 As Double
                    Dim cr2 As Double
                    Dim cr3 As Double
                    Dim cr4 As Double
                    Dim cr5 As Double
                    Dim cr6 As Double
                    Dim cr7 As Double
                    Dim cr8 As Double
                    Dim cr9 As Double

                    If (d1 * 2 > 4600) Then
                        bver1 = 6000
                    Else : bver1 = 4700
                    End If

                    If (d9 * 2 > 4600) Then
                        bver2 = 6000
                    Else : bver2 = 4700
                    End If

                    If (d6 * 2 > 4600) Then
                        bver3 = 6000
                    Else : bver3 = 4700
                    End If

                    If (d16 * 2 > 4600) Then
                        bver4 = 6000
                    Else : bver4 = 4700
                    End If

                    If (d4 * du4 < 4500) And (d4 * du4 > 3000) Then
                        cr1 = 1.4
                    ElseIf (d4 * du4 < 3000) Then
                        cr1 = 1.8
                    Else : cr1 = 1.0
                    End If

                    If (d5 * du5 < 4500) And (d5 * du5 > 3000) Then
                        cr2 = 1.2
                    ElseIf (d5 * du5 < 3000) Then
                        cr2 = 1.8
                    Else : cr2 = 1.0
                    End If

                    If (d7 * du7 < 4500) And (d7 * du7 > 3000) Then
                        cr3 = 1.2
                    ElseIf (d7 * du7 < 3000) Then
                        cr3 = 1.8
                    Else : cr3 = 1.0
                    End If

                    If (d8 * du8 < 4500) And (d8 * du8 > 3000) Then
                        cr4 = 1.2
                    ElseIf (d8 * du8 < 3000) Then
                        cr4 = 1.8
                    Else : cr4 = 1.0
                    End If

                    If (d20 * du20 < 4500) And (d20 * du20 > 3000) Then
                        cr5 = 1.2
                    ElseIf (d20 * du20 < 3000) Then
                        cr5 = 1.8
                    Else : cr5 = 1.0
                    End If

                    If (d13 * du13 < 4500) And (d13 * du13 > 3000) Then
                        cr6 = 1.2
                    ElseIf (d13 * du13 < 3000) Then
                        cr6 = 1.8
                    Else : cr6 = 1.0
                    End If

                    If (d14 * du14 < 4500) And (d14 * du14 > 3000) Then
                        cr7 = 1.2
                    ElseIf (d14 * du14 < 3000) Then
                        cr7 = 1.8
                    Else : cr7 = 1.0
                    End If

                    If (d17 * du17 < 4500) And (d17 * du17 > 3000) Then
                        cr8 = 1.2
                    ElseIf (d17 * du17 < 3000) Then
                        cr8 = 1.8
                    Else : cr8 = 1.0
                    End If

                    If (d18 * du18 < 4700) And (d18 * du18 > 3000) Then
                        cr9 = 1.2
                    ElseIf (d18 * du18 < 3000) Then
                        cr9 = 1.8
                    Else : cr9 = 1.0
                    End If

                    If (d22 < 4500) And (d22 > 3000) Then
                        cpl = 1.4
                    ElseIf (d22 < 3000) Then
                        cpl = 1.8
                    Else : cpl = 1.0
                    End If

                    precio = ((pa * xlWorkSheet2.Cells(67, 10).value * d13 * du13 * 1.25 * cr6 / 1000) + (pgr * d13 * du13 * 1.25 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(67, 10).value * bver2 * 2 / 1000) + (pgr * bver2 * 2 * xlWorkSheet2.Cells(67, 13).value * xlWorkSheet2.Cells(67, 12).value / 1000) + (pa * xlWorkSheet2.Cells(68, 10).value * d14 * du14 * 1.25 * cr7 / 1000) + (pgr * d14 * du14 * 1.25 * xlWorkSheet2.Cells(68, 13).value * xlWorkSheet2.Cells(68, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * bver4 * 2 / 1000) + (pgr * bver4 * 2 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(75, 10).value * d17 * du17 * 1.25 * cr8 / 1000) + (pgr * d17 * du17 * 1.25 * xlWorkSheet2.Cells(75, 13).value * xlWorkSheet2.Cells(75, 12).value / 1000) + (pa * xlWorkSheet2.Cells(74, 10).value * d18 * du18 * 1.25 * cr9 / 1000) + (pgr * d18 * du18 * 1.25 * xlWorkSheet2.Cells(74, 13).value * xlWorkSheet2.Cells(74, 12).value / 1000) + (pa * xlWorkSheet2.Cells(78, 10).value * d8 * du8 * 1.25 * cr4 / 1000) + (pgr * d8 * du8 * 1.25 * xlWorkSheet2.Cells(78, 13).value * xlWorkSheet2.Cells(78, 12).value / 1000) + (pa * xlWorkSheet2.Cells(69, 10).value * bver2 * 1.5 / 1000) + (pgr * bver2 * 1.5 * xlWorkSheet2.Cells(69, 13).value * xlWorkSheet2.Cells(69, 12).value / 1000) + (pa * xlWorkSheet2.Cells(71, 10).value * bver2 * 0.5 / 1000) + (pgr * bver2 * 0.5 * xlWorkSheet2.Cells(71, 13).value * xlWorkSheet2.Cells(71, 12).value / 1000) + (pa * xlWorkSheet2.Cells(70, 10).value * bver2 * 1 / 1000) + (pgr * bver2 * xlWorkSheet2.Cells(70, 13).value * xlWorkSheet2.Cells(70, 12).value / 1000) + (pa * xlWorkSheet2.Cells(86, 10).value * d22 * 1.25 * cpl / 1000) + (pgr * d22 * 1.25 * xlWorkSheet2.Cells(86, 13).value * xlWorkSheet2.Cells(86, 12).value / 1000)) * 2
                    descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                    lblcod21.Text = codleido.ToString
                    lbldes21.Text = descripcion
                    lbluds21.Text = 1
                    lblpre21.Text = precio
                    importe = precio * 1 * 2
                    lblimp21.Text = importe.ToString

                    lblcod21.Visible = True
                    lbldes21.Visible = True
                    lbluds21.Visible = True
                    lblpre21.Visible = True
                    lblimp21.Visible = True

                End If

            Catch ex As Exception

            End Try

            ' ******** FORRO VIGA PORTAOPERADOR *****************************************

                    Try

                        If (cmbfvig.SelectedValue = codleido.ToString) Then

                            Dim importe As Double

                            precio = xlWorkSheet.Cells(fil, 4).value
                            descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                            lblcod4.Text = codleido.ToString
                            lbldes4.Text = descripcion
                            lbluds4.Text = cmbfvigu.SelectedItem.ToString
                            lblpre4.Text = precio
                            importe = xlWorkSheet.Cells(fil, 4).value * Val(cmbfvigu.Text)
                            lblimp4.Text = importe.ToString

                            lblcod4.Visible = True
                            lbldes4.Visible = True
                            lbluds4.Visible = True
                            lblpre4.Visible = True
                            lblimp4.Visible = True

                        End If

                    Catch ex As Exception

                    End Try

                    ' ******** FORRO POSTE VERTICAL *****************************************

                    Try

                        If (cmbfpv.SelectedValue = codleido.ToString) Then

                            Dim importe As Double

                            precio = xlWorkSheet.Cells(fil, 4).value
                            descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                            lblcod5.Text = codleido.ToString
                            lbldes5.Text = descripcion
                            lbluds5.Text = cmbfpvu.SelectedItem.ToString
                            lblpre5.Text = precio
                            importe = xlWorkSheet.Cells(fil, 4).value * Val(cmbfpvu.Text)
                            lblimp5.Text = importe.ToString

                            lblcod5.Visible = True
                            lbldes5.Visible = True
                            lbluds5.Visible = True
                            lblpre5.Visible = True
                            lblimp5.Visible = True

                        End If

                    Catch ex As Exception

                    End Try

                    ' ******** FORRO TAPA OPERADOR *****************************************

                    Try

                        If (cmbftap.SelectedValue = codleido.ToString) Then

                            Dim importe As Double

                            precio = xlWorkSheet.Cells(fil, 4).value
                            descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                            lblcod6.Text = codleido.ToString
                            lbldes6.Text = descripcion
                            lbluds6.Text = cmbftapu.SelectedItem.ToString
                            lblpre6.Text = precio
                            importe = xlWorkSheet.Cells(fil, 4).value * Val(cmbftapu.Text)
                            lblimp6.Text = importe.ToString

                            lblcod6.Visible = True
                            lbldes6.Visible = True
                            lbluds6.Visible = True
                            lblpre6.Visible = True
                            lblimp6.Visible = True

                        End If

                    Catch ex As Exception

                    End Try

                    ' ******** FORRO PERFILERÍA *****************************************

                    Try

                        If (cmbfper.SelectedValue = codleido.ToString) Then

                            Dim importe As Double

                            If ((cmbfper.SelectedIndex = 5) Or (cmbfper.SelectedIndex = 6)) Then

                                precio = xlWorkSheet.Cells(fil, 4).value
                                descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                                lblcod7.Text = codleido.ToString
                                lbldes7.Text = descripcion
                                lbluds7.Text = tb_metros.Text
                                lblpre7.Text = precio
                                importe = xlWorkSheet.Cells(fil, 4).value * Val(tb_metros.Text)
                                lblimp7.Text = importe.ToString

                                lblcod7.Visible = True
                                lbldes7.Visible = True
                                lbluds7.Visible = True
                                lblpre7.Visible = True
                                lblimp7.Visible = True

                            ElseIf ((cmbfper.SelectedIndex <> 5) Or (cmbfper.SelectedIndex <> 6)) Then

                                precio = xlWorkSheet.Cells(fil, 4).value
                                descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                                lblcod7.Text = codleido.ToString
                                lbldes7.Text = descripcion
                                lbluds7.Text = cmbfperu.SelectedItem.ToString
                                lblpre7.Text = precio
                                importe = xlWorkSheet.Cells(fil, 4).value * Val(cmbfperu.Text)
                                lblimp7.Text = importe.ToString

                                lblcod7.Visible = True
                                lbldes7.Visible = True
                                lbluds7.Visible = True
                                lblpre7.Visible = True
                                lblimp7.Visible = True

                            End If

                        End If

                    Catch ex As Exception

                    End Try

                    ' ******** RADARES 1 *****************************************

                    Try

                        If (cmbrad1.SelectedValue = codleido.ToString) Then

                            Dim importe As Double

                            precio = xlWorkSheet.Cells(fil, 4).value
                            descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                            lblcod8.Text = codleido.ToString
                            lbldes8.Text = descripcion
                            lbluds8.Text = cmbradu1.SelectedItem.ToString
                            lblpre8.Text = precio
                            importe = xlWorkSheet.Cells(fil, 4).value * Val(cmbradu1.Text)
                            lblimp8.Text = importe.ToString

                            lblcod8.Visible = True
                            lbldes8.Visible = True
                            lbluds8.Visible = True
                            lblpre8.Visible = True
                            lblimp8.Visible = True

                        End If

                    Catch ex As Exception

                    End Try

                    ' ******** RADARES 2 *****************************************

                    Try

                        If (cmbrad2.SelectedValue = codleido.ToString) Then

                            Dim importe As Double

                            precio = xlWorkSheet.Cells(fil, 4).value
                            descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                            lblcod9.Text = codleido.ToString
                            lbldes9.Text = descripcion
                            lbluds9.Text = cmbradu2.SelectedItem.ToString
                            lblpre9.Text = precio
                            importe = xlWorkSheet.Cells(fil, 4).value * Val(cmbradu2.Text)
                            lblimp9.Text = importe.ToString

                            lblcod9.Visible = True
                            lbldes9.Visible = True
                            lbluds9.Visible = True
                            lblpre9.Visible = True
                            lblimp9.Visible = True

                        End If

                    Catch ex As Exception

                    End Try

                    ' ******** RADARES 3 *****************************************

                    Try

                        If (cmbrad3.SelectedValue = codleido.ToString) Then

                            Dim importe As Double

                            precio = xlWorkSheet.Cells(fil, 4).value
                            descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                            lblcod10.Text = codleido.ToString
                            lbldes10.Text = descripcion
                            lbluds10.Text = cmbradu3.SelectedItem.ToString
                            lblpre10.Text = precio
                            importe = xlWorkSheet.Cells(fil, 4).value * Val(cmbradu3.Text)
                            lblimp10.Text = importe.ToString

                            lblcod10.Visible = True
                            lbldes10.Visible = True
                            lbluds10.Visible = True
                            lblpre10.Visible = True
                            lblimp10.Visible = True

                        End If

                    Catch ex As Exception

                    End Try

                    ' ******** PULSADORES 1 *****************************************

                    Try

                        If (cmbpul1.SelectedValue = codleido.ToString) Then

                            Dim importe As Double

                            precio = xlWorkSheet.Cells(fil, 4).value
                            descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                            lblcod11.Text = codleido.ToString
                            lbldes11.Text = descripcion
                            lbluds11.Text = cmbpulu1.SelectedItem.ToString
                            lblpre11.Text = precio
                            importe = xlWorkSheet.Cells(fil, 4).value * Val(cmbpulu1.Text)
                            lblimp11.Text = importe.ToString

                            lblcod11.Visible = True
                            lbldes11.Visible = True
                            lbluds11.Visible = True
                            lblpre11.Visible = True
                            lblimp11.Visible = True

                        End If

                    Catch ex As Exception

                    End Try

                    ' ******** PULSADORES 2 *****************************************

                    Try

                        If (cmbpul2.SelectedValue = codleido.ToString) Then

                            Dim importe As Double

                            precio = xlWorkSheet.Cells(fil, 4).value
                            descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                            lblcod12.Text = codleido.ToString
                            lbldes12.Text = descripcion
                            lbluds12.Text = cmbpulu2.SelectedItem.ToString
                            lblpre12.Text = precio
                            importe = xlWorkSheet.Cells(fil, 4).value * Val(cmbpulu2.Text)
                            lblimp12.Text = importe.ToString

                            lblcod12.Visible = True
                            lbldes12.Visible = True
                            lbluds12.Visible = True
                            lblpre12.Visible = True
                            lblimp12.Visible = True

                        End If

                    Catch ex As Exception

                    End Try

                    ' ******** SELECTORES *****************************************

                    Try

                        If (cmbsel.SelectedValue = codleido.ToString) Then

                            Dim importe As Double

                            precio = xlWorkSheet.Cells(fil, 4).value
                            descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                            lblcod13.Text = codleido.ToString
                            lbldes13.Text = descripcion
                            lbluds13.Text = cmbselu.SelectedItem.ToString
                            lblpre13.Text = precio
                            importe = xlWorkSheet.Cells(fil, 4).value * Val(cmbselu.Text)
                            lblimp13.Text = importe.ToString

                            lblcod13.Visible = True
                            lbldes13.Visible = True
                            lbluds13.Visible = True
                            lblpre13.Visible = True
                            lblimp13.Visible = True

                        End If

                    Catch ex As Exception

                    End Try

                    ' ******** CERROJOS *****************************************

                    Try

                        If (cmbcerr.SelectedValue = codleido.ToString) Then

                            Dim importe As Double

                            precio = xlWorkSheet.Cells(fil, 4).value
                            descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                            lblcod14.Text = codleido.ToString
                            lbldes14.Text = descripcion
                            lbluds14.Text = cmbcerru.SelectedItem.ToString
                            lblpre14.Text = precio
                            importe = xlWorkSheet.Cells(fil, 4).value * Val(cmbcerru.Text)
                            lblimp14.Text = importe.ToString

                            lblcod14.Visible = True
                            lbldes14.Visible = True
                            lbluds14.Visible = True
                            lblpre14.Visible = True
                            lblimp14.Visible = True

                        End If

                    Catch ex As Exception

                    End Try

                    ' ******** OTROS 1 *****************************************

                    Try

                        If (cmbotr1.SelectedValue = codleido.ToString) Then

                            Dim importe As Double

                            If ((cmbotr1.SelectedIndex = 3) Or (cmbotr1.SelectedIndex = 7)) Then

                                precio = xlWorkSheet.Cells(fil, 4).value
                                descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                                lblcod15.Text = codleido.ToString
                                lbldes15.Text = descripcion
                                lbluds15.Text = tb_otr1.Text
                                lblpre15.Text = precio
                                importe = xlWorkSheet.Cells(fil, 4).value * Val(tb_otr1.Text)
                                lblimp15.Text = importe.ToString

                                lblcod15.Visible = True
                                lbldes15.Visible = True
                                lbluds15.Visible = True
                                lblpre15.Visible = True
                                lblimp15.Visible = True

                            ElseIf ((cmbotr1.SelectedIndex <> 3) Or (cmbotr1.SelectedIndex <> 7)) Then

                                precio = xlWorkSheet.Cells(fil, 4).value
                                descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                                lblcod15.Text = codleido.ToString
                                lbldes15.Text = descripcion
                                lbluds15.Text = cmbotru1.SelectedItem.ToString
                                lblpre15.Text = precio
                                importe = xlWorkSheet.Cells(fil, 4).value * Val(cmbotru1.Text)
                                lblimp15.Text = importe.ToString

                                lblcod15.Visible = True
                                lbldes15.Visible = True
                                lbluds15.Visible = True
                                lblpre15.Visible = True
                                lblimp15.Visible = True

                            End If

                        End If

                    Catch ex As Exception

                    End Try

                    ' ******** OTROS 2 *****************************************

                    Try

                        If (cmbotr2.SelectedValue = codleido.ToString) Then

                            Dim importe As Double

                            If ((cmbotr2.SelectedIndex = 3) Or (cmbotr2.SelectedIndex = 7)) Then

                                precio = xlWorkSheet.Cells(fil, 4).value
                                descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                                lblcod16.Text = codleido.ToString
                                lbldes16.Text = descripcion
                                lbluds16.Text = tb_otr1.Text
                                lblpre16.Text = precio
                                importe = xlWorkSheet.Cells(fil, 4).value * Val(tb_otr2.Text)
                                lblimp16.Text = importe.ToString

                                lblcod16.Visible = True
                                lbldes16.Visible = True
                                lbluds16.Visible = True
                                lblpre16.Visible = True
                                lblimp16.Visible = True

                            ElseIf ((cmbotr2.SelectedIndex <> 3) Or (cmbotr2.SelectedIndex <> 7)) Then

                                precio = xlWorkSheet.Cells(fil, 4).value
                                descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                                lblcod16.Text = codleido.ToString
                                lbldes16.Text = descripcion
                                lbluds16.Text = cmbotru2.SelectedItem.ToString
                                lblpre16.Text = precio
                                importe = xlWorkSheet.Cells(fil, 4).value * Val(cmbotru2.Text)
                                lblimp16.Text = importe.ToString

                                lblcod16.Visible = True
                                lbldes16.Visible = True
                                lbluds16.Visible = True
                                lblpre16.Visible = True
                                lblimp16.Visible = True

                            End If

                        End If

                    Catch ex As Exception

                    End Try

                    ' ******** ARTÍCULOS OPCIONALES PARA CARPINTERÍAS *****************************************

                    Try

                        If (cmbart1.SelectedValue = codleido.ToString) Then

                            Dim importe As Double

                            If ((cmbart1.SelectedIndex = cmbart1.FindStringExact("Perfil pilastra en aluminio BW52")) Or (cmbart1.SelectedIndex = cmbart1.FindStringExact("Tocho para pilastra y puerta batiente"))) Then

                                precio = xlWorkSheet.Cells(fil, 4).value
                                descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                                lblcod17.Text = codleido.ToString
                                lbldes17.Text = descripcion
                                lbluds17.Text = cmbartu1.SelectedItem.ToString
                                lblpre17.Text = precio
                                importe = xlWorkSheet.Cells(fil, 4).value * Val(cmbartu1.Text)
                                lblimp17.Text = importe.ToString

                                lblcod17.Visible = True
                                lbldes17.Visible = True
                                lbluds17.Visible = True
                                lblpre17.Visible = True
                                lblimp17.Visible = True

                            ElseIf ((cmbart1.SelectedIndex = cmbart1.FindStringExact("Fijo superior en perfilería BW52 en aluminio extrusionado, completamente acabado")) Or (cmbart1.SelectedIndex = cmbart1.FindStringExact("Perfil de cierre para MI"))) Then

                                precio = xlWorkSheet.Cells(fil, 4).value
                                descripcion = xlWorkSheet.Cells(fil, 1).value.ToString
                                lblcod17.Text = codleido.ToString
                                lbldes17.Text = descripcion
                                lbluds17.Text = tb_art1.Text
                                lblpre17.Text = precio
                                importe = xlWorkSheet.Cells(fil, 4).value * Val(tb_art1.Text)
                                lblimp17.Text = importe.ToString

                                lblcod17.Visible = True
                                lbldes17.Visible = True
                                lbluds17.Visible = True
                                lblpre17.Visible = True
                                lblimp17.Visible = True

                            End If

                        End If

                    Catch ex As Exception

                    End Try

                    ' ******** PRECIO SIN IVA E IMPORTE TOTAL *****************************************

                    Label160.Visible = True
                    Label161.Visible = True
                    Label162.Visible = True
                    Label163.Visible = True

                    Dim preciosiniva As Double
                    Dim precioconiva As Double

                    preciosiniva = Val(lblimp1.Text) + Val(lblimp2.Text) + Val(lblimp3.Text) + Val(lblimp4.Text) + Val(lblimp5.Text) + Val(lblimp6.Text) + Val(lblimp7.Text) + Val(lblimp8.Text) + Val(lblimp9.Text) + Val(lblimp10.Text) + Val(lblimp11.Text) + Val(lblimp12.Text) + Val(lblimp13.Text) + Val(lblimp14.Text) + Val(lblimp15.Text) + Val(lblimp16.Text) + Val(lblimp17.Text) + Val(lblimp18.Text) + Val(lblimp19.Text) + Val(lblimp20.Text) + Val(lblimp21.Text)
                    Label162.Text = preciosiniva.ToString
                    precioconiva = preciosiniva * 1.21
                    Label163.Text = precioconiva.ToString

        Next

        xlWorkBook.Close()
        xlApp.Quit()

        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)


        xlWorkBook2.Close()
        xlApp2.Quit()

        releaseObject(xlApp2)
        releaseObject(xlWorkBook2)
        releaseObject(xlWorkSheet2)

    End Sub

    Private Sub releaseObject(ByRef obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private WithEvents pd As New Printing.PrintDocument

    Private Sub imprimir()

        Try

            Dim PrintDialog1 As New PrintDialog()
            PrintDocument1.DefaultPageSettings.Margins.Left = 0
            PrintDocument1.DefaultPageSettings.Margins.Right = 0
            PrintDialog1.Document = PrintDocument1
            Dim result As DialogResult = PrintDialog1.ShowDialog()

            If (result = DialogResult.OK) Then
                PrintPreviewDialog1.Document = PrintDocument1
                PrintPreviewDialog1.Height = 1000
                PrintPreviewDialog1.Width = 1000
                PrintPreviewDialog1.ShowDialog()
            End If

        Catch ex As Exception

            MsgBox(ex.ToString, MsgBoxStyle.Critical, "Error al tratar de imprimir")

        End Try

    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage

        Using bmp As New Bitmap(Panel1.Width, Panel1.Height)
            Panel1.DrawToBitmap(bmp, New Rectangle(0, 0, Panel1.Width * 10, Panel1.Height * 2))
            e.Graphics.DrawImage(bmp, e.MarginBounds)
        End Using

    End Sub

End Class