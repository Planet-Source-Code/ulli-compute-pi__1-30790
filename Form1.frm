VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Form1"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8355
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   8355
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox Picture2 
      Height          =   5940
      Left            =   180
      ScaleHeight     =   5880
      ScaleWidth      =   6600
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   165
      Width           =   6660
      Begin VB.VScrollBar VScroll1 
         Enabled         =   0   'False
         Height          =   5880
         LargeChange     =   100
         Left            =   6360
         Max             =   -24000
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   240
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   36000
         Left            =   0
         ScaleHeight     =   36000
         ScaleWidth      =   6345
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   6345
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Rechts
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6990
      MaxLength       =   4
      TabIndex        =   0
      Text            =   "5000"
      Top             =   225
      Width           =   525
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Compute !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7005
      TabIndex        =   1
      Top             =   675
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Digits approx"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7590
      TabIndex        =   2
      Top             =   165
      Width           =   525
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program to compute Pi
'Found the method on the web
'Uses the power series expansion of atn(x) = x - x ^ 3 / 3 + x ^ 5 / 5 - ...
'together with  pi = 16 * atn(1 / 5) - 4 * atn(1 / 239)
'This gives about 1.4 decimals per term.

Option Explicit
DefLng A-Z

Private i
Private StartTime As Single
Private PiDigits()
Private Terms()
Private NumDigits
Private NumWords
Private LowLim
Private HighLim
Private Dividend
Private Divisor
Private Quotient1
Private Quotient2
Private Const TenThousend As Long = 10000

Private Sub Atan239()

  Dim Remainder1
  Dim Remainder2
  Dim Remainder3
  Dim Remainder4
  Const xSq As Long = 239 ^ 2

    Remainder1 = Terms(LowLim)
    LowLim = LowLim + 1
    For i = LowLim To HighLim
        Dividend = Remainder1 * TenThousend + Terms(i)
        Quotient1 = Dividend \ xSq
        Remainder1 = Dividend - Quotient1 * xSq

        Dividend = Remainder2 * TenThousend + Quotient1
        Quotient2 = Dividend \ Divisor
        Remainder2 = Dividend - Quotient2 * Divisor
        PiDigits(i) = PiDigits(i) + Quotient2

        Dividend = Remainder3 * TenThousend + Quotient1
        Quotient1 = Dividend \ xSq
        Remainder3 = Dividend - Quotient1 * xSq

        Dividend = Remainder4 * TenThousend + Quotient1
        Quotient2 = Dividend \ (Divisor + 2)
        Remainder4 = Dividend - Quotient2 * (Divisor + 2)
        PiDigits(i) = PiDigits(i) - Quotient2
        Terms(i) = Quotient1
    Next i
    Do While Terms(LowLim) = 0
        LowLim = LowLim + 1
    Loop
    Divisor = Divisor + 4
    
End Sub

Private Sub Atan5()

  Dim Remainder1
  Dim Remainder2
  Const xSq As Long = 5 ^ 2

    For i = LowLim To HighLim + 1
        Dividend = Remainder1 * TenThousend + Terms(i)
        Quotient1 = Dividend \ xSq
        Remainder1 = Dividend - Quotient1 * xSq
        Terms(i) = Quotient1
        Dividend = Remainder2 * TenThousend + Quotient1
        Quotient1 = Dividend \ Divisor
        Remainder2 = Dividend - Quotient1 * Divisor
        PiDigits(i) = PiDigits(i) - Quotient1
    Next i

    For i = HighLim + 2 To NumWords
        Dividend = Remainder2 * TenThousend
        Quotient1 = Dividend \ Divisor
        Remainder2 = Dividend - Quotient1 * Divisor
        PiDigits(i) = PiDigits(i) - Quotient1
    Next i

    Do While Terms(LowLim) = 0
        LowLim = LowLim + 1
    Loop
    If Terms(HighLim + 1) > 0 And HighLim < NumWords Then
        HighLim = HighLim + 1
    End If

    Divisor = Divisor + 2
    Remainder1 = 0
    Remainder2 = 0

    For i = LowLim To HighLim + 1
        Dividend = Remainder1 * TenThousend + Terms(i)
        Quotient1 = Dividend \ xSq
        Remainder1 = Dividend - Quotient1 * xSq
        Terms(i) = Quotient1
        Dividend = Remainder2 * TenThousend + Quotient1
        Quotient1 = Dividend \ Divisor
        Remainder2 = Dividend - Quotient1 * Divisor
        PiDigits(i) = PiDigits(i) + Quotient1
    Next i

    For i = HighLim + 2 To NumWords
        Dividend = Remainder2 * TenThousend
        Quotient1 = Dividend \ Divisor
        Remainder2 = Dividend - Quotient1 * Divisor
        PiDigits(i) = PiDigits(i) + Quotient1
    Next i

    Do While Terms(LowLim) = 0
        LowLim = LowLim + 1
    Loop
    If Terms(HighLim + 1) > 0 And HighLim < NumWords Then
        HighLim = HighLim + 1
    End If
    Divisor = Divisor + 2
    
End Sub

Private Sub Command1_Click()

  'pi = 16 * Atn(1 / 5) - 4 * Atn(1 / 239)
   
  Dim Remainder
    
    Picture1.Cls
    VScroll1 = 0
    DoEvents
    Screen.MousePointer = vbHourglass
    Command1.Enabled = False
    VScroll1.Enabled = False
    Text1 = Abs(Text1)
    NumDigits = Val(Text1) + 8
    NumWords = NumDigits \ 4 + 1
    ReDim PiDigits(NumWords + 1), Terms(NumWords + 1)
    StartTime = Timer
                                        
    '16 * atn(1 / 5)
    LowLim = 1
    HighLim = 2
    '16/5 = 3.2 = first Terms of first series
    Terms(1) = 3
    Terms(2) = 2000
    
    PiDigits(1) = PiDigits(1) + Terms(1)
    PiDigits(2) = PiDigits(2) + Terms(2)
    
    Divisor = 3
    Do Until LowLim >= NumWords
        Atan5
    Loop

    '- 4 * atn(1 / 239)
    Remainder = 4
    '4 / 239 = 0,0167364... = first Terms of second series
    For i = 2 To NumWords
        Dividend = Remainder * TenThousend
        Terms(i) = Dividend \ 239
        Remainder = Dividend - Terms(i) * 239
        PiDigits(i) = PiDigits(i) - Terms(i)
    Next i
    
    LowLim = 2
    HighLim = NumWords
    
    Divisor = 3
    Do Until LowLim >= NumWords
        Atan239
    Loop
                                      
    'ripple carry / borrow
    For i = NumWords To 1 Step -1
        If PiDigits(i) < 0 Then
            Quotient1 = PiDigits(i) \ TenThousend
            PiDigits(i) = PiDigits(i) - (Quotient1 - 1) * TenThousend
            PiDigits(i - 1) = PiDigits(i - 1) + Quotient1 - 1
          ElseIf PiDigits(i) >= TenThousend Then 'NOT PIDIGITS(I)...
            Quotient1 = PiDigits(i) \ TenThousend
            PiDigits(i) = PiDigits(i) - Quotient1 * TenThousend
            PiDigits(i - 1) = PiDigits(i - 1) + Quotient1
        End If
    Next i
    
    PrintOut
    
    Picture1.Print " Computation time: "; Timer - StartTime; " seconds"
    VScroll1.SmallChange = Picture1.TextHeight("A")
    VScroll1.LargeChange = VScroll1.SmallChange * 5
    i = Picture2.ScaleHeight - Picture1.CurrentY
    If i < 0 Then
        VScroll1.Max = i
        VScroll1.Enabled = True
    End If
    Command1.Enabled = True
    Screen.MousePointer = vbDefault

End Sub

Private Sub PrintOut()
  
    If PiDigits(NumWords - 1) >= 5000 Then 'round
        PiDigits(NumWords - 2) = PiDigits(NumWords - 2) + 1
    End If
    Picture1.Print " pi ~ 3. ..."
    Picture1.Print " ";
    For i = 1 To NumWords \ 3 - 1
        Picture1.Print Format$(PiDigits(3 * (i - 1) + 2), "000\ 0");
        Picture1.Print Format$(PiDigits(3 * (i - 1) + 3), "00\ 00");
        Picture1.Print Format$(PiDigits(3 * (i - 1) + 4), "0\ 000\ ");
        If i Mod 5 = 0 Then
            Picture1.Print
            Picture1.Print " ";
        End If
    Next i
    Picture1.Print Format$(PiDigits(3 * (i - 1) + 2), "000\ 0");
    Picture1.Print
    Picture1.Print

End Sub

Private Sub VScroll1_Change()

    VScroll1_Scroll

End Sub

Private Sub VScroll1_Scroll()

    Picture1.Top = VScroll1
    Text1.SetFocus

End Sub

':) Ulli's VB Code Formatter V2.9.4 (14.01.2002 18:41:57) 22 + 216 = 238 Lines
