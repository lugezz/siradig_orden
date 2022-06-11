VERSION 5.00
Begin VB.Form frmCarpeta 
   BackColor       =   &H00808080&
   Caption         =   "Orden Carpetas"
   ClientHeight    =   9480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10380
   Icon            =   "frmCarpeta.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9480
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbOrden 
      Height          =   315
      Left            =   4290
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1200
      Width           =   3975
   End
   Begin VB.CheckBox chkCambia 
      BackColor       =   &H00808080&
      Caption         =   "Cambia Nombre?"
      Height          =   255
      Left            =   4320
      TabIndex        =   7
      Top             =   6930
      Width           =   2175
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "Procesar"
      Height          =   430
      Left            =   4470
      TabIndex        =   2
      Top             =   7515
      Width           =   1300
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   430
      Left            =   6870
      TabIndex        =   3
      Top             =   7515
      Width           =   1300
   End
   Begin VB.ListBox lstArchivos 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6810
      Left            =   150
      TabIndex        =   5
      Top             =   1200
      Width           =   3100
   End
   Begin VB.CommandButton cmdSelArch 
      Caption         =   "Seleccione Carpeta"
      Height          =   400
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1725
   End
   Begin VB.Label lblCant 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Datos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4785
      Left            =   4230
      TabIndex        =   6
      Top             =   2070
      Width           =   3855
   End
   Begin VB.Label lblCarp 
      BackStyle       =   0  'Transparent
      Caption         =   "Datos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2070
      TabIndex        =   4
      Top             =   540
      Width           =   7275
   End
End
Attribute VB_Name = "frmCarpeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BD As New tbrBasedeDatos.clsDataBase
Dim FSO As New FileSystemObject
Dim Carpetas() As String, Canti() As Long

Private Sub cmdProcesar_Click()
    lblCant = ""
    
    ReDim Preserve Carpetas(0)
    ReDim Preserve Canti(0)
    
    ProcesarTXT
End Sub

Private Sub ProcesarTXT()
    Dim I As Long, j As Long
       
    For I = 0 To lstArchivos.ListCount - 1
        MoverArchivo lstArchivos.List(I), lblCarp.Caption, chkCambia.Value
    Next I
    
    lblCant = ""
    For j = 1 To UBound(Carpetas)
        If lblCant <> "" Then lblCant = lblCant + vbCrLf
        lblCant = lblCant + Carpetas(j) + ": " + CStr(Canti(j)) + " archivos"
    Next j
    
End Sub

Private Sub MoverArchivo(mNombre As String, mCarp As String, CbioNom As Boolean)
    ' Formato es como 20077004913_2014_presentacion_002.xml ó pdf, nro de presentacion puede venir con _b no lo considero
    ' Paso a Split CUIT_Año_presn_nro. Omito si hay algo más del sp (3)
    Dim SP() As String, Carp As String, NArchivo As String, tmP As String
    
    '1) Separo el split
    SP = Split(mNombre, "_")
    
    '2) Busco la carpeta correspondiente
    Carp = BD.GetValInRS(cmbOrden, "Criterio", "CUIL = '" + SP(0) + "'")
    
    If Carp = "" Then Carp = "Sin Asignar"
    
    '3) Asigno Nombre
    NArchivo = mNombre
    
    If CbioNom Then
        tmP = BD.GetValInRS(cmbOrden, "Leg", "CUIL = '" + SP(0) + "'")
        
        If tmP <> "" Then 'si no está no cambio
            If UBound(SP) < 4 Then SP(3) = Left(SP(3), Len(SP(3)) - 4)
            NArchivo = "L." + tmP + "." + BD.GetValInRS(cmbOrden, "Nombre", "CUIL = '" + SP(0) + "'")
            NArchivo = NArchivo + " (" + SP(1) + " Pr. " + SP(3) + ")." + Right(mNombre, 3)
        End If
    End If
  
    '4) Muevo
    If Right(mCarp, 1) <> "\" Then mCarp = mCarp + "\"
    
    If FSO.FolderExists(mCarp + Carp) = False Then
        FSO.CreateFolder mCarp + Carp
    End If
    
    SumoCarpeta Carp
    FSO.MoveFile mCarp + mNombre, mCarp + Carp + "\" + NArchivo
End Sub

Private Sub SumoCarpeta(mCcarp As String)
    Dim I As Long, j As Long, Estaa As Boolean
    
    Estaa = False
    I = UBound(Carpetas)
     
    For j = 0 To I
        If mCcarp = Carpetas(j) Then
            Canti(j) = Canti(j) + 1
            Estaa = True
            Exit For
        End If
    Next j
    
    If Estaa = False Then
        ReDim Preserve Carpetas(I + 1)
        Carpetas(I + 1) = mCcarp
        ReDim Preserve Canti(I + 1)
        Canti(I + 1) = 1
    End If
    
End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSelArch_Click()

    Set objShell = CreateObject("Shell.Application")
      
     'Abre el cuadro de diálogo para seleccionar
    
    'MsgBox AP
    Set objFolder = objShell.BrowseForFolder(0, "Seleccione carpetilla", 0, "")
    
    If objFolder Is Nothing Then Exit Sub
    lblCarp = objFolder.Self.Path
    CargarArch
End Sub

Private Sub Form_Load()

    lblCarp = ""
    Limpiar
    
    BD.cn_CONECTAR_MDB AP + "Orden Legajos.mdb"
    CargarListas
    
    ReDim Preserve Carpetas(0)
    ReDim Preserve Canti(0)
    
    BD.cn_CONECTAR_MDB AP + "Orden Legajos.mdb"
End Sub

Private Sub Limpiar()
    lstArchivos.Clear
    lblCant = ""
    
End Sub

Private Sub CargarListas()
    cmbOrden.Clear
    
    Dim Tabla As TableDef
    Dim base As Database
    Set base = OpenDatabase(AP + "Orden Legajos.mdb")
    
    For Each Tabla In base.TableDefs
        If Left(Tabla.Name, 2) <> "MS" Then cmbOrden.AddItem Tabla.Name
    Next Tabla
    
    BD.CN_CLOSE
    Set base = Nothing
    cmbOrden.ListIndex = 0
End Sub

Private Sub CargarArch()

    'Variable de tipo FILE y FOLDER para listar los archivos de un path
    Dim El_Archivo As File
    Dim El_Directorio As Folder

    'Si no hay items en el List sale
    If lblCarp = "" Then Exit Sub
    
    lstArchivos.Clear
    
    'Nuevo objeto FileSystemObject
    Set FSO = New FileSystemObject
    
    ' Obtiene el directorio
    Set El_Directorio = FSO.GetFolder(lblCarp)
        
    ' Lista los ficheros de esta carpeta
    For Each El_Archivo In El_Directorio.Files
       'Añade la ruta
       lstArchivos.AddItem El_Archivo.Name
        
    Next El_Archivo

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set FSO = Nothing
    BD.CN_CLOSE
    Set BD = Nothing
End Sub

