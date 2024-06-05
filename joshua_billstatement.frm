VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Billing Statement"
   ClientHeight    =   10665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7980
   FillColor       =   &H00FFFFC0&
   LinkTopic       =   "Form1"
   ScaleHeight     =   10665
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame tbl 
      Height          =   2895
      Left            =   4560
      TabIndex        =   60
      Top             =   4320
      Width           =   3135
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Add Previous Bill"
         Height          =   375
         Left            =   120
         TabIndex        =   72
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox txtpr4 
         Height          =   375
         Left            =   1560
         TabIndex        =   71
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   0
         TabIndex        =   70
         Top             =   1920
         Width           =   1575
      End
      Begin VB.TextBox txtpr3 
         Height          =   375
         Left            =   1560
         TabIndex        =   69
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   0
         TabIndex        =   68
         Top             =   1560
         Width           =   1575
      End
      Begin VB.TextBox txtpr2 
         Height          =   375
         Left            =   1560
         TabIndex        =   67
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   0
         TabIndex        =   66
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtpr1 
         Height          =   375
         Left            =   1560
         TabIndex        =   63
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   0
         TabIndex        =   62
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Water Bill"
         Height          =   375
         Left            =   1560
         TabIndex        =   65
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label32 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Date"
         Height          =   375
         Left            =   0
         TabIndex        =   64
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Caption         =   "Previous Bill"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "COMPUTE"
      Height          =   375
      Left            =   2880
      TabIndex        =   59
      Top             =   9000
      Width           =   1575
   End
   Begin VB.TextBox txtgrand 
      Height          =   375
      Left            =   5760
      TabIndex        =   58
      Top             =   8520
      Width           =   1575
   End
   Begin VB.TextBox Text26 
      Height          =   375
      Left            =   5760
      TabIndex        =   57
      Top             =   8160
      Width           =   1575
   End
   Begin VB.TextBox Text25 
      Height          =   375
      Left            =   1800
      TabIndex        =   54
      Top             =   9720
      Width           =   5655
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Check1"
      Height          =   195
      Left            =   6240
      TabIndex        =   51
      Top             =   3960
      Width           =   135
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   195
      Left            =   6240
      TabIndex        =   49
      Top             =   3600
      Width           =   135
   End
   Begin VB.TextBox Text24 
      Height          =   375
      Left            =   4440
      TabIndex        =   47
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox Text23 
      Height          =   375
      Left            =   1560
      TabIndex        =   44
      Top             =   4200
      Width           =   1335
   End
   Begin VB.TextBox Text22 
      Height          =   375
      Left            =   1560
      TabIndex        =   43
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txttotal 
      Height          =   375
      Left            =   2880
      TabIndex        =   42
      Top             =   8520
      Width           =   1575
   End
   Begin VB.TextBox txtothers 
      Height          =   375
      Left            =   2880
      TabIndex        =   41
      Top             =   8160
      Width           =   1575
   End
   Begin VB.TextBox txtmaterials 
      Height          =   375
      Left            =   2880
      TabIndex        =   40
      Top             =   7800
      Width           =   1575
   End
   Begin VB.TextBox txtsurcharge 
      Height          =   375
      Left            =   2880
      TabIndex        =   39
      Top             =   7440
      Width           =   1575
   End
   Begin VB.TextBox txtmortuary 
      Height          =   375
      Left            =   2880
      TabIndex        =   38
      Top             =   7080
      Width           =   1575
   End
   Begin VB.TextBox txtmedical 
      Height          =   375
      Left            =   2880
      TabIndex        =   37
      Top             =   6720
      Width           =   1575
   End
   Begin VB.TextBox txtidfee 
      Height          =   375
      Left            =   2880
      TabIndex        =   36
      Top             =   6360
      Width           =   1575
   End
   Begin VB.TextBox txtpca 
      Height          =   375
      Left            =   2880
      TabIndex        =   35
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox txtcpc 
      Height          =   375
      Left            =   2880
      TabIndex        =   34
      Top             =   5640
      Width           =   1575
   End
   Begin VB.TextBox txtcbu 
      Height          =   375
      Left            =   2880
      TabIndex        =   33
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox txtcurrent 
      Height          =   375
      Left            =   2880
      TabIndex        =   32
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox txtcmu 
      Height          =   375
      Left            =   2880
      TabIndex        =   31
      Top             =   4560
      Width           =   1575
   End
   Begin VB.TextBox txtprevious 
      Height          =   375
      Left            =   2880
      TabIndex        =   30
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox txtpresent 
      Height          =   375
      Left            =   2880
      TabIndex        =   15
      Top             =   3840
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   6000
      TabIndex        =   13
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   2880
      Width           =   5295
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   2520
      Width           =   5295
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   2160
      Width           =   5295
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   1800
      Width           =   5295
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label30 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Invoice No."
      Height          =   255
      Left            =   4560
      TabIndex        =   56
      Top             =   8280
      Width           =   1335
   End
   Begin VB.Label Label29 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Grand Total"
      Height          =   255
      Left            =   4560
      TabIndex        =   55
      Top             =   8640
      Width           =   1335
   End
   Begin VB.Label Label28 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Remarks"
      Height          =   255
      Left            =   840
      TabIndex        =   53
      Top             =   9840
      Width           =   1335
   End
   Begin VB.Label Label27 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Cancel"
      Height          =   255
      Left            =   6480
      TabIndex        =   52
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label Label26 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Paid"
      Height          =   255
      Left            =   6480
      TabIndex        =   50
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Due Date"
      Height          =   375
      Left            =   4440
      TabIndex        =   48
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Meter Reading"
      Height          =   375
      Left            =   2880
      TabIndex        =   46
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Service Period"
      Height          =   375
      Left            =   1560
      TabIndex        =   45
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label22 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Others"
      Height          =   255
      Left            =   720
      TabIndex        =   29
      Top             =   8280
      Width           =   2175
   End
   Begin VB.Label Label21 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Total Amount"
      Height          =   255
      Left            =   720
      TabIndex        =   28
      Top             =   8640
      Width           =   2175
   End
   Begin VB.Label Label20 
      BackColor       =   &H00C0FFFF&
      Caption         =   "PCA"
      Height          =   255
      Left            =   720
      TabIndex        =   27
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0FFFF&
      Caption         =   "ID Fee"
      Height          =   255
      Left            =   720
      TabIndex        =   26
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Label Label18 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Medical"
      Height          =   255
      Left            =   720
      TabIndex        =   25
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Label Label17 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Mortuary"
      Height          =   255
      Left            =   720
      TabIndex        =   24
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Label Label16 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Surcharge"
      Height          =   255
      Left            =   720
      TabIndex        =   23
      Top             =   7560
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Materials Fittings"
      Height          =   255
      Left            =   720
      TabIndex        =   22
      Top             =   7920
      Width           =   2175
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Present"
      Height          =   255
      Left            =   720
      TabIndex        =   21
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Previous"
      Height          =   255
      Left            =   720
      TabIndex        =   20
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cubic Meter Consumed"
      Height          =   255
      Left            =   720
      TabIndex        =   19
      Top             =   4680
      Width           =   2175
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Current Bill"
      Height          =   255
      Left            =   720
      TabIndex        =   18
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0FFFF&
      Caption         =   "CBU"
      Height          =   255
      Left            =   720
      TabIndex        =   17
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "CPC"
      Height          =   255
      Left            =   720
      TabIndex        =   16
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Billing Date:"
      Height          =   255
      Left            =   4920
      TabIndex        =   14
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Consumer Type"
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Barangay"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Street"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Name:"
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Consumer ID"
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Billing No."
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   7440
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Billing Statement"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    txtcmu.Text = SafeDouble(txtpresent.Text) - SafeDouble(txtprevious.Text)
    txtcurrent.Text = SafeDouble(txtcmu.Text) * 29.42
    
    Dim current As Double
    Dim cbu As Double
    Dim cpc As Double
    Dim pca As Double
    Dim idfee As Double
    Dim medical As Double
    Dim mortuary As Double
    Dim surcharge As Double
    Dim materials As Double
    Dim others As Double
    
    current = SafeDouble(txtcurrent.Text)
    cbu = SafeDouble(txtcbu.Text)
    cpc = SafeDouble(txtcpc.Text)
    pca = SafeDouble(txtpca.Text)
    idfee = SafeDouble(txtidfee.Text)
    medical = SafeDouble(txtmedical.Text)
    mortuary = SafeDouble(txtmortuary.Text)
    surcharge = SafeDouble(txtsurcharge.Text)
    materials = SafeDouble(txtmaterials.Text)
    others = SafeDouble(txtothers.Text)
    
    
    txttotal.Text = current + cbu + cpc + pca + idfee + medical + mortuary + surcharge + materials + others
    txtgrand.Text = SafeDouble(txttotal.Text)
    
    
    
End Sub

Function SafeDouble(ByVal txt As String) As Double
    If IsNumeric(txt) Then
        SafeDouble = CDbl(txt)
    Else
        SafeDouble = 0
    End If
End Function

Private Sub Command2_Click()
pr1 = SafeDouble(txtpr1.Text)
    pr2 = SafeDouble(txtpr2.Text)
    pr3 = SafeDouble(txtpr3.Text)
    pr4 = SafeDouble(txtpr4.Text)
    txtgrand.Text = SafeDouble(txttotal.Text) + pr1 + pr2 + pr3 + pr4
End Sub
