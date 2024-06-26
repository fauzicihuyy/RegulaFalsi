VERSION 5.00
Begin VB.Form Regula_Falsi 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FARID_AR_207011085_C"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   19140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox fxm 
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox xm 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox NilaiX2 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin VB.TextBox iterasi 
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton Hitung 
      Caption         =   "Hitung"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox NilaiX1 
      Height          =   375
      Left            =   1680
      TabIndex        =   0
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Metode Regula Falsi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   120
      TabIndex        =   6
      Top             =   240
      Width           =   19215
      Begin VB.ListBox List7 
         Height          =   5130
         Left            =   16680
         TabIndex        =   25
         Top             =   1320
         Width           =   2055
      End
      Begin VB.ListBox List6 
         Height          =   5130
         Left            =   14640
         TabIndex        =   24
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CommandButton Reset 
         Caption         =   "Reset"
         Height          =   495
         Left            =   360
         TabIndex        =   22
         Top             =   5280
         Width           =   975
      End
      Begin VB.ListBox List5 
         Height          =   5130
         Left            =   12600
         TabIndex        =   16
         Top             =   1320
         Width           =   1935
      End
      Begin VB.ListBox List4 
         Height          =   5130
         Left            =   10440
         TabIndex        =   15
         Top             =   1320
         Width           =   2055
      End
      Begin VB.ListBox List3 
         Height          =   5130
         Left            =   8280
         TabIndex        =   14
         Top             =   1320
         Width           =   2055
      End
      Begin VB.ListBox List2 
         Height          =   5130
         Left            =   6120
         TabIndex        =   13
         Top             =   1320
         Width           =   2055
      End
      Begin VB.ListBox List1 
         Height          =   5130
         Left            =   3960
         TabIndex        =   12
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "Xmid"
         Height          =   375
         Left            =   8280
         TabIndex        =   27
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label9 
         Caption         =   "X2"
         Height          =   375
         Left            =   6120
         TabIndex        =   26
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label La 
         Caption         =   "Persamaan   f(x)=x^3 - 7x + 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   4575
      End
      Begin VB.Label Label8 
         Caption         =   "Error"
         Height          =   375
         Left            =   16680
         TabIndex        =   21
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "F(Xmid)"
         Height          =   375
         Left            =   14760
         TabIndex        =   20
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "F(X2)"
         Height          =   375
         Left            =   12600
         TabIndex        =   19
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label5 
         Caption         =   "F(X1)"
         Height          =   375
         Left            =   10440
         TabIndex        =   18
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label4 
         Caption         =   "X1"
         Height          =   375
         Left            =   3960
         TabIndex        =   17
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "X1"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label label2 
         Caption         =   "X2"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label label3 
         Caption         =   "iterasi (n)"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label label3 
         Caption         =   "Xmid"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label label3 
         Caption         =   "F(xmid)"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   3720
         Width           =   975
      End
   End
End
Attribute VB_Name = "Regula_Falsi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Hitung_Click()
Dim x1, x2, xmid, fx1, fx2, fxmid, er As Double
Dim i, n As Integer

If NilaiX1.Text = "" Then
    MsgBox "Nilai X1 belum diisi !", vbInformation, "Konfirmasi"
    Cancel = True
    NilaiX1.SetFocus
    Exit Sub
Else
If NilaiX2.Text = "" Then
    MsgBox "Nilai X2 belum diisi !", vbInformation, "Konfirmasi"
    Cancel = True
    NilaiX2.SetFocus
    Exit Sub
Else
If iterasi.Text = "" Then
    MsgBox "Nilai iterasi belum diisi !", vbInformation, "Konfirmasi"
    Cancel = True
    iterasi.SetFocus
    Exit Sub
Else
    'mendefinisikan variable
    x1 = Val(NilaiX1.Text)
    x2 = Val(NilaiX2.Text)
    n = Val(iterasi.Text)

    'menghitung persamaan
    For i = 0 To n
        fx1 = (x1 ^ 3) - (7 * x1) + 1
        fx2 = (x2 ^ 3) - (7 * x2) + 1

        xmid = ((fx2 * x1) - (fx1 * x2)) / (fx2 - fx1)
        fxmid = (xmid ^ 3) - (7 * xmid) + 1

        xm.Text = xmid
        fxm.Text = fxmid
        'pERHITUNGAN ERROR
        er = ((xmid - x2) / xmid)
        If er < 0 Then er = er * -1


        If fx1 * fxmid < 0 Then
            x2 = xmid
            fx2 = fxmid
            
            'output
            List1.AddItem (x1)
            List2.AddItem (x2)
            List3.AddItem (xmid)
            List4.AddItem (fx1)
            List5.AddItem (fx2)
            List6.AddItem (fxmid)
            List7.AddItem (er)
        Else
            x1 = xmid
            fx1 = fxmid

            'output
            List1.AddItem (x1)
            List2.AddItem (x2)
            List3.AddItem (xmid)
            List4.AddItem (fx1)
            List5.AddItem (fx2)
            List6.AddItem (fxmid)
            List7.AddItem (er)
        End If
    Next i
End If
End If
End If
    
End Sub

Private Sub Reset_Click()
On Error Resume Next
NilaiX1.Text = ""
NilaiX2.Text = ""
iterasi.Text = ""
xm.Text = ""
fxm.Text = ""
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
List6.Clear
List7.Clear

NilaiX1.SetFocus

End Sub
