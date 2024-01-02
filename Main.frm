VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmKeithley 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Keithley2400 measure and logger by 9A4AM - Mario Anèiæ@2023"
   ClientHeight    =   13350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19335
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   13350
   ScaleWidth      =   19335
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "DMM Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   12000
      TabIndex        =   24
      Top             =   10080
      Width           =   3615
      Begin VB.CommandButton btnOutput 
         BackColor       =   &H008080FF&
         Caption         =   "Output ON"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txtVolt 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   720
         TabIndex        =   26
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton btnVolt 
         BackColor       =   &H000000FF&
         Caption         =   "S E T"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "SET Output"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1245
         TabIndex        =   29
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "SET Source Voltage "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   750
         TabIndex        =   27
         Top             =   360
         Width           =   2205
      End
   End
   Begin VB.CommandButton btnRead 
      BackColor       =   &H00C0C000&
      Caption         =   "READ DMM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   12360
      Width           =   3855
   End
   Begin VB.Frame frmActvalue 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Actual Value from DMM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   240
      TabIndex        =   11
      Top             =   10080
      Width           =   3855
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "mA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   360
         Left            =   3180
         TabIndex        =   17
         Top             =   480
         Width           =   495
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   3300
         TabIndex        =   16
         Top             =   1080
         Width           =   255
      End
      Begin VB.Label lblValueDMM_V 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "NO DATA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label lblValueDMM_A 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "NO DATA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   480
         Width           =   2895
      End
   End
   Begin VB.CommandButton btnSet 
      BackColor       =   &H000080FF&
      Caption         =   "              S E T                           AND                           I N I T             DMM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   10200
      Width           =   3375
   End
   Begin VB.CommandButton btnExit 
      BackColor       =   &H0000C000&
      Caption         =   "E X I T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12000
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   12360
      Width           =   3495
   End
   Begin VB.Frame Mode 
      BackColor       =   &H0080FFFF&
      Caption         =   "Select Mode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   8160
      TabIndex        =   7
      Top             =   12000
      Width           =   3495
      Begin VB.ComboBox cmbMode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         ItemData        =   "Main.frx":030A
         Left            =   120
         List            =   "Main.frx":0317
         TabIndex        =   8
         Top             =   480
         Width           =   3135
      End
   End
   Begin VB.Frame frmTime 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Sample Nr"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   15840
      TabIndex        =   5
      Top             =   10080
      Width           =   3375
      Begin VB.CommandButton btnSampleNr 
         BackColor       =   &H000080FF&
         Caption         =   "S E T"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtSampleNr 
         Alignment       =   2  'Center
         BackColor       =   &H000080FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   480
         TabIndex        =   21
         Top             =   2400
         Width           =   975
      End
      Begin VB.CommandButton btnSetInterval 
         BackColor       =   &H0000FFFF&
         Caption         =   "S E T"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txtSampleSet 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   480
         TabIndex        =   18
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "Number of sample"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   645
         TabIndex        =   22
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "Sample interval in ms"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   480
         TabIndex        =   19
         Top             =   1080
         Width           =   2265
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   360
         Left            =   1500
         TabIndex        =   6
         Top             =   360
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H000000FF&
      Caption         =   "STOP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   12360
      Width           =   3375
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H0000FF00&
      Caption         =   "START"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   12360
      Width           =   3375
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   10800
      Top             =   12600
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   9975
      Left            =   0
      OleObjectBlob   =   "Main.frx":0338
      TabIndex        =   2
      Top             =   0
      Width           =   19215
   End
   Begin VB.Frame Serial 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Select COM Port"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8160
      TabIndex        =   0
      Top             =   10800
      Width           =   3495
      Begin VB.ComboBox cmbPort 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3135
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   10080
      Top             =   12600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label lblDMM_online 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "DMM Offline"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   8160
      TabIndex        =   14
      Top             =   10200
      Width           =   3495
   End
End
Attribute VB_Name = "frmKeithley"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim FSO As New FileSystemObject
Dim Time_count As Integer
Dim TxtStr As TextStream
Dim DMM_Result As Integer
Dim DMM_Mode As String
Dim SampleNumber As Integer
Dim SampleInterval As Integer
Dim Data(0 To 300, 0 To 2) As Variant
    Dim I As Long
    Dim J As Long
    Dim K As Long
Dim Num_var As Single
Dim SN As Double
Dim SN_Txt As String
Dim Volt As Double
Dim Volt_MAX_Cfg As Variant
Dim Amp_MAX_Cfg As Variant
Dim Output_Cfg As Variant
Dim Sample_Cfg As Variant
Dim Output_Status As Boolean






Sub ListComPorts()
cmbPort.Clear

Dim Registry As Object, Names As Variant, Types As Variant
Set Registry = GetObject("winmgmts:\\.\root\default:StdRegProv")
If Registry.EnumValues(&H80000002, "HARDWARE\DEVICEMAP\SERIALCOMM", Names, Types) <> 0 Then Exit Sub

Dim I As Long
If IsArray(Names) Then
    For I = 0 To UBound(Names)
        Dim PortName As Variant
        Registry.GetStringValue &H80000002, "HARDWARE\DEVICEMAP\SERIALCOMM", Names(I), PortName
        cmbPort.AddItem Mid$(PortName, 4, 2) '& " - " & Names(I)
    Next
End If

End Sub

Private Sub btnExit_Click()
Dim Exit_var As Integer
Exit_var = MsgBox("          Exit from program!?", vbYesNo)
If Exit_var = vbYes Then


Unload frmKeithley

End If

End Sub

Private Sub btnOutput_Click()
If Output_Cfg = "Manual" Then
If MSComm1.PortOpen = False Then
MSComm1.PortOpen = True
End If
If Output_Status = False Then
MSComm1.Output = ":OUTP ON" & Chr(13)
btnOutput.Caption = "Output OFF"
Output_Status = True
Else
MSComm1.Output = ":OUTP OFF" & Chr(13)
btnOutput.Caption = "Output ON"
Output_Status = False
End If
Else
MsgBox "Output in AUTO MODE!!", vbExclamation
End If
End Sub

Private Sub btnRead_Click()


Measure_DMM
lblValueDMM_A.Caption = SN_Txt
lblValueDMM_V.Caption = Volt
End Sub

Private Sub btnSampleNr_Click()

If txtSampleNr.Text = "" Or txtSampleNr.Text = " " Then
MsgBox "Enter number!!", vbExclamation
Else
If IsNumeric(txtSampleNr.Text) Then
SampleNumber = Val(txtSampleNr.Text)
Else
MsgBox "Enter number!!", vbExclamation
End If
End If
If SampleNumber > Sample_Cfg Or SampleNumber = 0 Then
SampleNumber = Sample_Cfg
End If
MsgBox "Sample number is set: " & SampleNumber, vbInformation
txtSampleNr.Text = SampleNumber
End Sub

Private Sub btnSet_Click()
If cmbMode = "" Or cmbPort = "" Then
MsgBox "First select COM Port and Mode!!", vbExclamation
Else
'MsgBox Val(cmbPort.Text)
If MSComm1.PortOpen = False Then
MSComm1.CommPort = Val(cmbPort.Text)
End If
If cmbMode.Text = "Voltage" Then
DMM_Mode = "volt"
ElseIf cmbMode.Text = "Current" Then
DMM_Mode = "curr"
Else
DMM_Mode = "res"
End If
'MsgBox DMM_Mode
cmdStart.Enabled = True
btnRead.Enabled = True
Init_DMM
End If
J = 100
End Sub

Private Sub btnSetInterval_Click()

If txtSampleSet.Text = "" Or txtSampleSet.Text = " " Then
MsgBox "Enter number!!", vbExclamation
Else
If IsNumeric(txtSampleSet.Text) Then
SampleInterval = Val(txtSampleSet.Text)
MsgBox "Sample interval is set: " & SampleInterval & " msec", vbInformation
Else
MsgBox "Enter number!!", vbExclamation
End If
End If

txtSampleSet.Text = SampleInterval



End Sub

Private Sub btnVolt_Click()
If txtVolt.Text = "" Or txtVolt.Text = " " Then
MsgBox "Enter number!!", vbExclamation
Else
If IsNumeric(txtVolt.Text) Then
Volt = Val(txtVolt.Text)
MsgBox "Source voltage is set: " & Volt, vbInformation
Else
MsgBox "Enter number!!", vbExclamation
End If
End If

txtVolt.Text = Volt
End Sub

Private Sub cmdStart_Click()

If cmbMode = "" Or cmbPort = "" Then
MsgBox "First select COM Port and Mode!!", vbExclamation
Else
cmdStart.Visible = False
cmdStop.Visible = True
Timer1.Enabled = True
btnRead.Enabled = False
btnExit.Enabled = False
btnRead.Enabled = False
btnSet.Enabled = False
btnVolt.Enabled = False
btnSetInterval.Enabled = False
btnSampleNr.Enabled = False
btnOutput.Enabled = False

cmdStop.SetFocus
Timer1.Interval = SampleInterval

End If
Graph_Fill
End Sub

Private Sub cmdStop_Click()
cmdStop.Visible = False
cmdStart.Visible = True
Timer1.Enabled = False
Time_count = 0
J = 0
I = 0
K = 0
MSComm1.Output = ":OUTP OFF" & Chr(13)
Output_Status = False
btnRead.Enabled = True
btnExit.Enabled = True
btnRead.Enabled = True
btnSet.Enabled = True
btnVolt.Enabled = True
btnOutput.Enabled = True
btnSetInterval.Enabled = True
btnSampleNr.Enabled = True
cmdStart.SetFocus

End Sub

Private Sub Form_Activate()
cmdStop.Visible = False
cmdStart.Visible = True
cmdStart.Enabled = False
btnRead.Enabled = False
Graph_Fill
End Sub

Private Sub Form_Load()
Volt_MAX_Cfg = ReadIniFile(App.Path & "\config.ini", "Graph", "Volt")
Amp_MAX_Cfg = ReadIniFile(App.Path & "\config.ini", "Graph", "Amp")
Sample_Cfg = ReadIniFile(App.Path & "\config.ini", "Graph", "Sample")
Output_Cfg = ReadIniFile(App.Path & "\config.ini", "DMM", "Output")


Timer1.Enabled = False

Time_count = 0
SampleNumber = Sample_Cfg
SampleInterval = 1000
Volt = 1
Timer1.Interval = SampleInterval
txtSampleSet.Text = SampleInterval
txtSampleNr.Text = SampleNumber
txtVolt = Volt
ListComPorts
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then
    MSComm1.Output = "*RST" & Chr(13)
    MSComm1.Output = ":SYSTEM:LOCAL" & Chr(13)
    MSComm1.PortOpen = False
    End If
End Sub

Private Sub Timer1_Timer()
Time_count = Time_count + 1
Measure_DMM
lblTime = Time_count
J = Val(SN_Txt)
I = I + 1
K = Val(Volt)
Data(I, 1) = J
lblValueDMM_A.Caption = J
lblValueDMM_V = Volt
Graph_Fill
Log_Data
If Time_count = SampleNumber Then
Timer1.Enabled = False
cmdStop.Visible = False
cmdStart.Visible = True
Time_count = 0
I = 0
MSComm1.Output = ":OUTP OFF" & Chr(13)
Output_Status = False
btnRead.Enabled = True
btnExit.Enabled = True
btnRead.Enabled = True
btnSet.Enabled = True
btnVolt.Enabled = True
btnSetInterval.Enabled = True
btnSampleNr.Enabled = True
btnOutput.Enabled = True
cmdStart.SetFocus

End If
End Sub


Private Sub Graph_Fill()



     'For I = 0 To 300
        'Data(I, 0) = " " '& CStr(I) 'The leading space causes these to be axis labels
                                   'instead of series values.
        Data(I, 0) = K    'Voltage
        'Data(I, 2) = I * 5
        'Data(I, 2) = J    'Current
     'Next

    With MSChart1
        .chartType = VtChChartType2dLine
        .RandomFill = False
        .ShowLegend = True
        .ChartData = Data
        With .Plot
            With .Wall.Brush
                .Style = VtBrushStyleSolid
                .FillColor.Set 255, 255, 255
            End With
            With .Axis(VtChAxisIdX).CategoryScale
                .Auto = False
                .DivisionsPerLabel = 1
                .DivisionsPerTick = 1
                .LabelTick = False
            End With
            With .Axis(VtChAxisIdY)
                .AxisTitle = "Volt"
                With .ValueScale
                    .Auto = False
                    .Minimum = 0
                    .Maximum = Volt_MAX_Cfg
                    .MajorDivision = 0.1
                End With
            End With
            With .Axis(VtChAxisIdY2)
                .AxisTitle = "miliAmp"
                With .ValueScale
                    .Auto = False
                    .Minimum = 0
                    .Maximum = Amp_MAX_Cfg
                    .MajorDivision = 1
                End With
            End With
            With .SeriesCollection(1)
                .LegendText = "Volt"
                With .Pen
                    .Width = ScaleX(1, vbPixels, vbTwips)
                    .VtColor.Set 255, 0, 0
                End With
            End With
            With .SeriesCollection(2)
                .LegendText = "miliAmp"
                .SecondaryAxis = True
                With .Pen
                    .Width = ScaleX(1, vbPixels, vbTwips)
                    .VtColor.Set 0, 255, 0
                End With
            End With
        End With
    End With
End Sub

Private Sub Init_DMM()
If MSComm1.PortOpen = False Then
MSComm1.PortOpen = True
End If

    MSComm1.Output = "*RST" & Chr(13)
    MSComm1.Output = ":disp:text:data '  Program by 9A4AM'" & Chr(13)
    MSComm1.Output = ":disp:text:stat 1" & Chr(13)
    Sleep 500
    MSComm1.Output = ":SOUR:FUNC VOLT" & Chr(13)
    MSComm1.Output = ":SOUR:VOLT:MODE FIXED" & Chr(13)
    MSComm1.Output = ":SOUR:VOLT:RANGE 20" & Chr(13)
    MSComm1.Output = ":SOUR:VOLT:LEV " & Volt & Chr(13)
    MSComm1.Output = ":SENS:CURR:PROT 1" & Chr(13)
    MSComm1.Output = ":SENS:FUNC 'CURR'" & Chr(13)
    MSComm1.Output = ":SENS:CURR:RANGE 1" & Chr(13)
    MSComm1.Output = ":FORM:ELEM CURR" & Chr(13)
    
    lblDMM_online.ForeColor = vbGreen
    lblDMM_online.Caption = "DMM Online"
    'MSComm1.Output = ":output 0" & Chr(13)
    'MSComm1.Output = ":sour1:cle:auto on" & Chr(13)
    'MSComm1.Output = ":format:elem curr" & Chr(13)
   ' MSComm1.Output = ":disp:enab 0" & Chr(13)
    'MSComm1.Output = ":OUTP ON" & Chr(13)
    'MSComm1.Output = ":READ?" & Chr(13)
    'MSComm1.Output = ":OUTP OFF" & Chr(13)
   
   
    'MSComm1.PortOpen = False
    
End Sub
Private Sub Measure_DMM()
Dim strBuffer As String
' Measure Current sample
'*RST
': SOUR: FUNC VOLT
': SOUR: VOLT: Mode Fixed
': SOUR: VOLT: RANG 20
': SOUR: VOLT: LEV 3
': SENS: CURR: PROT 0.01
': SENS: FUNC "CURR"
': SENS: CURR: RANG 1
': Form: ELEM CURR
':OUTP ON
':READ?
': OUTP OFF
'
'
If MSComm1.PortOpen = False Then
MSComm1.PortOpen = True
End If
    'MSComm1.Output = ":meas?" & Chr(13)
    If Output_Status = False Then
    MSComm1.Output = ":OUTP ON" & Chr(13)
    Output_Status = True
    Sleep 2000
    End If
   
    MSComm1.Output = ":READ?" & Chr(13)
 
    'MSComm1.Output = ":OUTP OFF" & Chr(13)
    
    Sleep 100
    Do
            strBuffer = strBuffer & MSComm1.Input
        Loop Until Right$(strBuffer, 1) = Chr(13) Or Right$(strBuffer, 1) = ""   ' DMM output complete when a CR is received
       
        If Right$(strBuffer, 1) <> Chr(13) Then strBuffer = "Buffer empty!!!     Try again." 'or DMM output empty show invalid command
        'Text1.Text = strBuffer   ' Display strBuffer in DMM Output textbox
         'Protokol.Mjer1.Text = Mid$(strBuffer, 3, 12) & vbNewLine
        Num_var = CSng(Mid$(strBuffer, 3, 12)) / 1000000
        SN = Left$(Num_var, 12)
        SN_Txt = Format$((SN * 1000000000), "########")
        If SN_Txt = "" Then
        SN_Txt = "0"
        End If
        
        strBuffer = ""      ' Set strBuffer to a NULL string
    
    ' MSComm1.PortOpen = False

End Sub
Private Sub Log_Data()
Set TxtStr = FSO.OpenTextFile(App.Path & "/log.dat", ForAppending)
TxtStr.WriteLine Format(Now, " dd-mmm-yyyy HH:nn:ss ") & ";" & Time_count & ";" & J & ";" & K 'DMM_Result
TxtStr.Close

Set FSO = Nothing
Set TxtStr = Nothing
End Sub
Private Sub Graph_Filling()
With MSChart1
        .chartType = VtChChartType2dLine
        .RandomFill = False
        .ShowLegend = True
        .ChartData = Data
End With
End Sub
