VERSION 5.00
Begin VB.Form Calculator 
   BackColor       =   &H00C000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "the4 best calclutaotr 2"
   ClientHeight    =   5805
   ClientLeft      =   2580
   ClientTop       =   1485
   ClientWidth     =   7995
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "System"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "CALC.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5805
   ScaleWidth      =   7995
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton Number 
      BackColor       =   &H00004080&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Index           =   7
      Left            =   120
      MaskColor       =   &H00004080&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   7755
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00004080&
      Caption         =   "8"
      Height          =   480
      Index           =   8
      Left            =   720
      MaskColor       =   &H00004080&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   480
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00004080&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   9
      Left            =   1320
      MaskColor       =   &H00004080&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton Cancel 
      BackColor       =   &H00004080&
      Caption         =   "C"
      Height          =   480
      Left            =   2040
      MaskColor       =   &H00004080&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   600
      UseMaskColor    =   -1  'True
      Width           =   480
   End
   Begin VB.CommandButton CancelEntry 
      BackColor       =   &H00004080&
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2640
      MaskColor       =   &H00004080&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   600
      Width           =   480
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00004080&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   4
      Left            =   120
      MaskColor       =   &H00008080&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   480
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00004080&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   5
      Left            =   720
      MaskColor       =   &H00004080&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00004080&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   6
      Left            =   1320
      MaskColor       =   &H00004080&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1200
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      BackColor       =   &H00004080&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   1
      Left            =   2040
      MaskColor       =   &H00004080&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      BackColor       =   &H00004080&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   3
      Left            =   2640
      MaskColor       =   &H00004080&
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1200
      UseMaskColor    =   -1  'True
      Width           =   480
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00004080&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   1
      Left            =   120
      MaskColor       =   &H00004080&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00004080&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   2
      Left            =   720
      MaskColor       =   &H00004080&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00004080&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   3
      Left            =   1320
      MaskColor       =   &H00004080&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      BackColor       =   &H00004080&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   2
      Left            =   2040
      MaskColor       =   &H00004080&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      BackColor       =   &H00004080&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   0
      Left            =   2640
      MaskColor       =   &H00004080&
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1800
      UseMaskColor    =   -1  'True
      Width           =   480
   End
   Begin VB.CommandButton Number 
      BackColor       =   &H00004080&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   0
      Left            =   120
      MaskColor       =   &H00004080&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   1080
   End
   Begin VB.CommandButton Decimal 
      BackColor       =   &H00004080&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1320
      MaskColor       =   &H00004080&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   480
   End
   Begin VB.CommandButton Operator 
      BackColor       =   &H00004080&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Index           =   4
      Left            =   3240
      MaskColor       =   &H00004080&
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   120
      UseMaskColor    =   -1  'True
      Width           =   4680
   End
   Begin VB.CommandButton Percent 
      BackColor       =   &H00004080&
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2640
      MaskColor       =   &H00004080&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2400
      UseMaskColor    =   -1  'True
      Width           =   480
   End
   Begin VB.Label Readout 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   105
      Width           =   3000
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ------------------------------------------------------------------------
'               Copyright (C) 1994 Microsoft Corporation
'
' You have a royalty-free right to use, modify, reproduce and distribute
' the Sample Application Files (and/or any modified version) in any way
' you find useful, provided that you agree that Microsoft has no warranty,
' obligations or liability for any Sample Application Files.
' ------------------------------------------------------------------------
Option Explicit
Dim Op1, Op2                ' Previously input operand.
Dim DecimalFlag As Integer  ' Decimal point present yet?
Dim NumOps As Integer       ' Number of operands.
Dim LastInput               ' Indicate type of last keypress event.
Dim OpFlag                  ' Indicate pending operation.
Dim TempReadout

' Click event procedure for C (cancel) key.
' Reset the display and initializes variables.
Private Sub Cancel_Click()
    Readout = Format(0, "0.")
    Op1 = 0
    Op2 = 0
    Form_Load
End Sub

' Click event procedure for CE (cancel entry) key.
Private Sub CancelEntry_Click()
    Readout = Format(0, "0.")
    DecimalFlag = False
    LastInput = "CE"
End Sub

' Click event procedure for decimal point (.) key.
' If last keypress was an operator, initialize
' readout to "0." Otherwise, append a decimal
' point to the display.
Private Sub Decimal_Click()
    If LastInput = "NEG" Then
        Readout = Format(0, "-0.")
    ElseIf LastInput <> "NUMS" Then
        Readout = Format(0, "0.")
    End If
    DecimalFlag = True
    LastInput = "NUMS"
End Sub

' Initialization routine for the form.
' Set all variables to initial values.
Private Sub Form_Load()
    DecimalFlag = False
    NumOps = 0
    LastInput = "NONE"
    OpFlag = " "
    Readout = Format(0, "0.")
    'Decimal.Caption = Format(0, ".")
End Sub

' Click event procedure for number keys (0-9).
' Append new number to the number in the display.
Private Sub Number_Click(Index As Integer)
    If LastInput <> "NUMS" Then
        Readout = Format(0, ".")
        DecimalFlag = False
    End If
    If DecimalFlag Then
        Readout = Readout + Number(Index).Caption
    Else
        Readout = Left(Readout, InStr(Readout, Format(0, ".")) - 1) + Number(Index).Caption + Format(0, ".")
    End If
    If LastInput = "NEG" Then Readout = "-" & Readout
    LastInput = "NUMS"
End Sub

' Click event procedure for operator keys (+, -, x, /, =).
' If the immediately preceeding keypress was part of a
' number, increments NumOps. If one operand is present,
' set Op1. If two are present, set Op1 equal to the
' result of the operation on Op1 and the current
' input string, and display the result.
Private Sub Operator_Click(Index As Integer)
    TempReadout = Readout
    If LastInput = "NUMS" Then
        NumOps = NumOps + 1
    End If
    Select Case NumOps
        Case 0
        If Operator(Index).Caption = "-" And LastInput <> "NEG" Then
            Readout = "-" & Readout
            LastInput = "NEG"
        End If
        Case 1
        Op1 = Readout
        If Operator(Index).Caption = "-" And LastInput <> "NUMS" And OpFlag <> "=" Then
            Readout = "-"
            LastInput = "NEG"
        End If
        Case 2
        Op2 = TempReadout
        Select Case OpFlag
            Case "+"
                Op1 = CDbl(Op1) + CDbl(Op2)
            Case "-"
                Op1 = CDbl(Op1) - CDbl(Op2)
            Case "X"
                Op1 = CDbl(Op1) * CDbl(Op2)
            Case "/"
                   Op1 = CDbl(Op1) / CDbl(Op2)
            Case "="
                Op1 = CDbl(Op2)
            Case "%"
                Op1 = CDbl(Op1) * CDbl(Op2)
            End Select
        Readout = Op1
        NumOps = 1
    End Select
    If LastInput <> "NEG" Then
        LastInput = "OPS"
        OpFlag = Operator(Index).Caption
    End If
End Sub

' Click event procedure for percent key (%).
' Compute and display a percentage of the first operand.
Private Sub Percent_Click()
    Readout = Readout / 100
    LastInput = "Ops"
    OpFlag = "%"
    NumOps = NumOps + 1
    DecimalFlag = True
End Sub


