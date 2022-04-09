VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11505
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17625
   LinkTopic       =   "Form1"
   ScaleHeight     =   767
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1175
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3870
      Left            =   12720
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   11
      Top             =   4920
      Width           =   3870
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3870
      Left            =   8640
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   10
      Top             =   4920
      Width           =   3870
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3870
      Left            =   4560
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   9
      Top             =   4920
      Width           =   3870
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3870
      Left            =   480
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   8
      Top             =   4920
      Width           =   3870
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3870
      Left            =   12720
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   7
      Top             =   240
      Width           =   3870
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3870
      Left            =   8640
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   6
      Top             =   240
      Width           =   3870
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3870
      Left            =   4560
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   5
      Top             =   240
      Width           =   3870
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   16440
      Top             =   10200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "picture(*.jpg)|*.jpg|picture(*.bmp)|*.bmp|모든영상(*.*)|*.*"
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3870
      Left            =   480
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   0
      Top             =   240
      Width           =   3870
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   735
      Left            =   9600
      TabIndex        =   3
      Top             =   4320
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
      Height          =   615
      Left            =   13920
      TabIndex        =   15
      Top             =   8880
      Width           =   1695
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   375
      Left            =   9840
      TabIndex        =   14
      Top             =   8880
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   495
      Left            =   5280
      TabIndex        =   13
      Top             =   8880
      Width           =   2175
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   615
      Left            =   1560
      TabIndex        =   12
      Top             =   8880
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   855
      Left            =   13800
      TabIndex        =   4
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   615
      Left            =   5520
      TabIndex        =   2
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Menu mnu_File 
      Caption         =   "파일"
      Begin VB.Menu mnu_Clean 
         Caption         =   "그림지우기"
      End
      Begin VB.Menu mnu_LoadPicture 
         Caption         =   "그림불러오기"
      End
      Begin VB.Menu mnu_End 
         Caption         =   "종료"
      End
   End
   Begin VB.Menu mnu_Gray 
      Caption         =   "그레이화"
   End
   Begin VB.Menu mnu_chapOne 
      Caption         =   "1장"
      Begin VB.Menu mnu_Mirroring 
         Caption         =   "미러링"
      End
   End
   Begin VB.Menu mnu_chapTwo 
      Caption         =   "2장"
      Begin VB.Menu mnu_ArithmeticOperation 
         Caption         =   "산술연산"
         Begin VB.Menu mnu_Addition 
            Caption         =   "덧셈연산"
         End
         Begin VB.Menu mnu_Subtraction 
            Caption         =   "뺄셈연산"
         End
         Begin VB.Menu mnu_Multiplication 
            Caption         =   "곱셈연산"
         End
         Begin VB.Menu mnu_Division 
            Caption         =   "나눗셈연산"
         End
         Begin VB.Menu mnu_XOR 
            Caption         =   "XOR연산"
         End
      End
      Begin VB.Menu mnu_Histogram 
         Caption         =   "히스토그램"
         Begin VB.Menu mnu_BHistogram 
            Caption         =   "히스토그램(기본)"
         End
         Begin VB.Menu mnu_EHistogram 
            Caption         =   "히스토그램(평활화)"
         End
         Begin VB.Menu mnu_SHistogram 
            Caption         =   "히스토그램(명세화)"
         End
      End
      Begin VB.Menu mnu_Stretch 
         Caption         =   "명암대비스트레칭"
         Begin VB.Menu mnu_BasicStretch 
            Caption         =   "기본명암대비스트레칭"
         End
         Begin VB.Menu mnu_EndInStretch 
            Caption         =   "앤드인명암대비스트레칭"
         End
      End
      Begin VB.Menu mnu_LightAndShade 
         Caption         =   "명암변화"
         Begin VB.Menu mnu_Reverse 
            Caption         =   "역변환"
         End
         Begin VB.Menu mnu_Gamma 
            Caption         =   "감마상관변환"
         End
         Begin VB.Menu mnu_Bitclimpping 
            Caption         =   "비트클림핑"
         End
         Begin VB.Menu mnu_Posterizing 
            Caption         =   "포스트라이징"
         End
         Begin VB.Menu mnu_Solarizing 
            Caption         =   "솔라이징"
         End
         Begin VB.Menu mnu_Parabola 
            Caption         =   "파라볼라"
         End
         Begin VB.Menu mnu_ContrastTransform 
            Caption         =   "명암대비변환"
         End
         Begin VB.Menu mnu_ContrastCompress 
            Caption         =   "명암압축변환"
         End
      End
   End
   Begin VB.Menu mnu_ChapThree 
      Caption         =   "3장"
      Begin VB.Menu mnu_Embosing 
         Caption         =   "엠보싱"
      End
      Begin VB.Menu mnu_Blurring 
         Caption         =   "블러링"
      End
      Begin VB.Menu mnu_Sharpning 
         Caption         =   "샤프닝"
         Begin VB.Menu mnu_MaskSharpning 
            Caption         =   "마스크샤프닝"
         End
         Begin VB.Menu mnu_HighPass 
            Caption         =   "고주파통과필터링1"
         End
         Begin VB.Menu mnu_HighPass2 
            Caption         =   "고주파통과필터링2"
         End
         Begin VB.Menu mnu_HighPass3 
            Caption         =   "고주파통과필터링3"
         End
         Begin VB.Menu mnu_UnSharpning 
            Caption         =   "언샤프마스킹"
         End
         Begin VB.Menu mnu_HighBoost 
            Caption         =   "고주파지원필터링"
         End
      End
      Begin VB.Menu mnu_Edge 
         Caption         =   "에지검출"
         Begin VB.Menu mnu_SimilarEdge 
            Caption         =   "유사연산자"
         End
         Begin VB.Menu mnu_SubtractionEdge 
            Caption         =   "차연산자"
         End
         Begin VB.Menu mnu_ThresholdEdge 
            Caption         =   "임계값이용"
         End
         Begin VB.Menu mnu_MaxMinThresholdEdge 
            Caption         =   "임계값(최대/최소)"
         End
         Begin VB.Menu mnu_FDerivation 
            Caption         =   "1차미분"
            Begin VB.Menu mnu_SobelFD 
               Caption         =   "소벨엣지"
            End
         End
         Begin VB.Menu mnu_Sderivation 
            Caption         =   "2차미분"
            Begin VB.Menu mnu_LaplacianSD 
               Caption         =   "라플라시안"
            End
         End
         Begin VB.Menu mnu_Compass 
            Caption         =   "컴파스"
         End
         Begin VB.Menu mnu_ColorEdge 
            Caption         =   "컬러엣지"
         End
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim picHeight, picWidth As Integer
Dim r, g, b As Long
Dim i, j As Long
Dim cc, mm, yy As Long
Dim pixel As Double
Dim Gray As Long
Dim grayarr(256, 256) As Long
Dim value As Single
Dim result As Long
Dim r_result, g_result, b_result As Single
Dim gamma As Single
Dim nbit As Long, modnum As Long
Dim modr As Long, modg As Long, modb As Long
Dim modi As Long
Dim threshold As Long
Dim r1 As Long, g1 As Long, b1 As Long
Dim r2 As Long, g2 As Long, b2 As Long
Dim Low, High As Long
Dim desthisto(256) As Long
Dim num As Long
Dim min_num, max_num As Long
Dim histo(256) As Long
Dim normailzedsum As Long
Dim histoequal(256) As Long, sumarr(256) As Long
Dim sum As Long
Dim minus As Long
Dim min As Long
Dim equalArr(256, 256) As Long
Dim histospeci(256) As Long
Dim histooutput(256) As Long
Dim mask(8) As Single, blur(256, 256) As Long
Dim max As Long, n As Long
Dim X As Long, y As Long
Dim y0(8) As Long, x0(8) As Long, k As Long '행/열 검출마스크(3*3)





Private Sub Form_Load()
picHeight = Picture1.ScaleHeight - 1
picWidth = Picture1.ScaleWidth - 1


End Sub

Private Sub mnu_Addition_Click()
f_Addition
End Sub
Private Sub f_Addition()
Picture3.Cls
value = InputBox("값입력?")
f_Gray
For i = 0 To picHeight
For j = 0 To picWidth
pixel = grayarr(j, i)
result = pixel + value
If result > 255 Then result = 255
Picture3.PSet (j, i), RGB(result, result, result)
Next j
Next i
Label3.Caption = "덧셈연산"
End Sub

Private Sub mnu_BasicStretch_Click()
f_BasicStretch
End Sub
Private Sub f_BasicStretch()
Picture5.Cls: Picture6.Cls
f_BHistogram

For i = 0 To 255
desthisto(i) = 0
Next i
For i = 0 To 255
If (histo(i) > 0) Then
Low = i
Exit For
End If
Next i
For i = 255 To 0 Step -1
If (histo(i) > 0) Then
High = i
Exit For
End If
Next i
For i = 0 To picHeight
For j = 0 To picWidth
value = ((grayarr(j, i) - Low) / (High - Low)) * 255
If (value > 255) Then value = 0
desthisto(value) = desthisto(value) + 1
Picture5.PSet (j, i), RGB(value, value, value)
Next j
Next i
For i = 0 To 255
Picture6.Line (i, picHeight)-(i, picHeight - (desthisto(i) \ 5))
Next i
Label5.Caption = "기본명암대비영상"
Label6.Caption = "기본명암대비스트레칭"
End Sub

Private Sub mnu_BHistogram_Click()
f_BHistogram
End Sub
Private Sub f_BHistogram()
Picture3.Cls
f_Gray
For i = 0 To 255
histo(i) = 0
Next i
For i = 0 To picHeight
For j = 0 To picWidth
Gray = grayarr(j, i)
histo(Gray) = histo(Gray) + 1
Next j
Next i
For i = 0 To 255
Picture3.Line (i, picHeight)-(i, picHeight - (histo(i) \ 5))
Next i
Label3.Caption = "원영상 히스토그램"
End Sub
Private Sub mnu_Bitclimpping_Click()
f_Bitclimpping
End Sub
Private Sub f_Bitclimpping()
Picture5.Cls
nbit = InputBox("몇비트로 할건가요?")
modnum = 256 \ (2 ^ nbit)
For i = 0 To picHeight
For j = 0 To picWidth
pixel = Picture1.Point(j, i)
r = pixel And &HFF
g = (pixel And &HFF00&) \ &H100
b = (pixel And &HFF0000) \ &H10000
modr = r Mod modnum
modg = g Mod modnum
modb = b Mod modnum
Picture5.PSet (j, i), RGB(modr, modg, modb)
Next j
Next i
Label5.Caption = nbit & "컬러비트클림핑"
End Sub

Private Sub mnu_Blurring_Click()
f_Blurring
End Sub
Private Sub f_Blurring()
Picture4.Cls
f_Gray
mask(0) = 1 / 9: mask(1) = 1 / 9: mask(2) = 1 / 9
mask(3) = 1 / 9: mask(4) = 1 / 9: mask(5) = 1 / 9
mask(6) = 1 / 9: mask(7) = 1 / 9: mask(8) = 1 / 9
For i = 1 To picHeight - 1
For j = 1 To picWidth - 1
value = 0: n = 0
For y = i - 1 To i + 1
For X = j - 1 To j + 1
value = value + (grayarr(X, y) * mask(n))
n = n + 1
Next X
Next y
If value > 255 Then value = 255 '클램핑 (오버플로우 방지)
If value < 0 Then value = 0 '클램핑 (언더플로우 방지)
blur(j, i) = value
Picture4.PSet (j, i), RGB(value, value, value)
Next j
Next i
Label4.Caption = "블러링"

End Sub

Private Sub mnu_Clean_Click()
f_Clean
End Sub
Private Sub f_Clean()
Picture2.Cls
Picture3.Cls
Picture4.Cls
Picture5.Cls
Picture6.Cls
Picture7.Cls
Picture8.Cls

End Sub

Private Sub mnu_Compass_Click()
f_Compass
End Sub
Private Sub f_Compass()
'Form2.Picture1(0).Cls: Form2.Picture1(1).Cls: Form2.Picture1(2).Cls
'Form2.Picture1(3).Cls: Form2.Picture1(4).Cls: Form2.Picture1(5).Cls
'Form2.Picture1(6).Cls: Form2.Picture1(7).Cls: Form2.Picture1(8).Cls
For t = 0 To 8
Form2.Picture1(t).Cls:
Next t


Dim H(8, 8) As Long
Dim x1 As Long, y1 As Long, t As Long
f_Gray
Form2.Show
'H1
H(0, 0) = 1: H(0, 1) = 1: H(0, 2) = -1
H(0, 3) = 1: H(0, 4) = -2: H(0, 5) = -1
H(0, 6) = 1: H(0, 7) = 1: H(0, 8) = -1

'H2
H(1, 0) = 1: H(1, 1) = -1: H(1, 2) = -1
H(1, 3) = 1: H(1, 4) = -2: H(1, 5) = -1
H(1, 6) = 1: H(1, 7) = 1: H(1, 8) = 1

'H3
H(2, 0) = -1: H(2, 1) = -1: H(2, 2) = -1
H(2, 3) = 1: H(2, 4) = -2: H(2, 5) = 1
H(2, 6) = 1: H(2, 7) = 1: H(2, 8) = 1

'H4
H(3, 0) = -1: H(3, 1) = -1: H(3, 2) = 1
H(3, 3) = -1: H(3, 4) = -2: H(3, 5) = 1
H(3, 6) = 1: H(3, 7) = 1: H(3, 8) = 1

'H5
H(4, 0) = -1: H(4, 1) = 1: H(4, 2) = 1
H(4, 3) = -1: H(4, 4) = -2: H(4, 5) = 1
H(4, 6) = -1: H(4, 7) = 1: H(4, 8) = 1

'H6
H(5, 0) = 1: H(5, 1) = 1: H(5, 2) = 1
H(5, 3) = -1: H(5, 4) = -2: H(5, 5) = 1
H(5, 6) = -1: H(5, 7) = -1: H(5, 8) = 1

'H7
H(6, 0) = 1: H(6, 1) = 1: H(6, 2) = 1
H(6, 3) = 1: H(6, 4) = -2: H(6, 5) = 1
H(6, 6) = -1: H(6, 7) = -1: H(6, 8) = -1

'H8
H(7, 0) = 1: H(7, 1) = 1: H(7, 2) = 1
H(7, 3) = 1: H(7, 4) = -2: H(7, 5) = -1
H(7, 6) = 1: H(7, 7) = -1: H(7, 8) = -1

For t = 0 To 7
For i = 1 To picHeight - 1
For j = 1 To picWidth - 1
value = 0
n = 0: x1 = 0
For y = i - 1 To i + 1
For X = j - 1 To j + 1
x1 = x1 + (H(t, n) * grayarr(X, y))
n = n + 1
Next X
Next y
If x1 < 0 Then x1 = 0
If x1 > 255 Then x1 = 255
If value < x1 Then value = x1
Form2.Picture1(t).PSet (j, i), RGB(value, value, value)
Next j
Next i
Next t

For i = 1 To picHeight - 1
For j = 1 To picWidth - 1
value = 0
For t = 0 To 7
n = 0: x1 = 0
For y = i - 1 To i + 1
For X = j - 1 To j + 1
x1 = x1 + (H(t, n) * grayarr(X, y))
n = n + 1
Next X
Next y
If x1 < 0 Then x1 = 0: If x1 > 255 Then x1 = 255
If value < x1 Then value = x1
Next t
Form2.Picture1(8).PSet (j, i), RGB(value, value, value)
Next j
Next i







End Sub
Private Sub mnu_ContrastCompress_Click()
f_ContrastCompress

End Sub
Private Sub f_ContrastCompress()
Picture4.Cls
Low = InputBox("Low")
High = InputBox("High")
f_Gray
For i = 0 To picHeight
For j = 0 To picWidth
value = (grayarr(j, i) + Low) * (High - Low) / 255
Picture4.PSet (j, i), RGB(value, value, value)
Next j
Next i
Label4.Caption = "명암대비압축변환"

End Sub
Private Sub mnu_ContrastTransform_Click()
f_ContrastTransform

End Sub
Private Sub f_ContrastTransform()
Picture3.Cls
Low = InputBox("Low")
High = InputBox("High")
f_Gray
For i = 0 To picHeight
For j = 0 To picWidth
If grayarr(j, i) <= Low Then
value = 0
ElseIf grayarr(j, i) >= High Then
value = 255
Else
value = (grayarr(j, i) - Low) * 255 / (High - Low)
End If
Picture3.PSet (j, i), RGB(value, value, value)
Next j
Next i
Label3.Caption = "명암대비변환"



End Sub
Private Sub mnu_Division_Click()
f_Division
End Sub
Private Sub f_Division()
Picture6.Cls
value = InputBox("값입력?")
f_Gray
For i = 0 To picHeight
For j = 0 To picWidth
pixel = grayarr(j, i)
result = pixel / value
If result > 255 Then result = 255
If result < 0 Then result = 0
Picture6.PSet (j, i), RGB(result, result, result)
Next j
Next i
Label6.Caption = "나눗셈연산"


End Sub

Private Sub mnu_EHistogram_Click()
f_EHistogram
End Sub
Private Sub f_EHistogram()
Picture5.Cls: Picture6.Cls
f_BHistogram
sum = 0
For i = 0 To picHeight
For j = 0 To picWidth
equalArr(j, i) = 0
Next j
Next i
For i = 0 To 255
histoequal(i) = 0
sum = sum + histo(i)
normalizedsum = sum * 255 / (picWidth * picHeight)
If normalizedsum > 255 Then normalizedsum = 255
sumarr(i) = normalizedsum
Next i
For i = o To picHeight
For j = 0 To picWidth
pixel = grayarr(j, i)
result = sumarr(pixel)
equalArr(j, i) = result
histoequal(result) = histoequal(result) + 1
Picture5.PSet (j, i), RGB(result, result, result)

Next j
Next i
Label5.Caption = "평활화된 영상"
For i = 0 To 255
Picture6.Line (i, picHeight)-(i, picHeight - (histoequal(i) \ 5))
Next i
Label6.Caption = "평활화된 히스토그램"
End Sub

Private Sub mnu_Embosing_Click()
f_Embosing

End Sub
Private Sub f_Embosing()
Picture3.Cls
f_Gray
mask(0) = -1: mask(1) = 0: mask(2) = 0
mask(3) = 0: mask(4) = 0: mask(5) = 0
mask(6) = 0: mask(7) = 0: mask(8) = 1
For i = 1 To picHeight - 1
For j = 1 To picWidth - 1
n = 0: value = 0
For y = i - 1 To i + 1
For X = j - 1 To j + 1
value = value + (grayarr(X, y) * mask(n))
n = n + 1
Next X
Next y
value = value + 128
If value > 255 Then value = 255
If value < 0 Then value = 0
Picture3.PSet (j, i), RGB(value, value, value)
Next j
Next i
Label3.Caption = "엠보싱"
End Sub
Private Sub mnu_END_Click()
End
End Sub

Private Sub mnu_EndInStretch_Click()
f_EndInStretch
End Sub
Private Sub f_EndInStretch()
Picture7.Cls: Picture8.Cls
f_BHistogram

For i = 0 To 255
desthisto(i) = 0
Next i
min_num = InputBox("최소픽셀의 값?")
For i = 0 To 255
If (histo(i) > min_num) Then
Low = i
Exit For
End If
Next i
For i = 255 To 0 Step -1
If (histo(i) > min_num) Then
High = i
Exit For
End If
Next i
For i = 0 To picHeight
For j = 0 To picWidth
value = ((grayarr(j, i) - Low) / (High - Low)) * 255
If (value > 255) Then value = 255
If (value < 0) Then value = 0
desthisto(value) = desthisto(value) + 1
Picture7.PSet (j, i), RGB(value, value, value)
Next j
Next i
For i = 0 To 255
Picture8.Line (i, picHeight)-(i, picHeight - (desthisto(i) \ 5))
Next i
Label7.Caption = "앤드인 명암대비 영상"
Label8.Caption = "앤드인 명암대비 히스토그램"

End Sub

Private Sub mnu_Gamma_Click()
f_Gamma
End Sub
Private Sub f_Gamma()
gamma = InputBox("감마값?")
If gamma > 1 Then
Picture3.Cls
Else
Picture4.Cls
End If
For i = 0 To picHeight
For j = 0 To picWidth
pixel = Picture1.Point(j, i)
r = pixel And &HFF
g = (pixel And &HFF00&) \ &H100
b = (pixel And &HFF0000) \ &H10000
r = ((r / 255) ^ (1 / gamma)) * 255
g = ((g / 255) ^ (1 / gamma)) * 255
b = ((b / 255) ^ (1 / gamma)) * 255
If r > 255 Then r = 255
If g > 255 Then g = 255
If b > 255 Then b = 255
If gamma > 1 Then
Picture3.PSet (j, i), RGB(r, g, b)
Else
Picture4.PSet (j, i), RGB(r, g, b)
End If
Next j
Next i
If gamma > 1 Then
Label3.Caption = "gamma=" & gamma & "인 감마상관변환"
Else
Label4.Caption = "gamma=" & gamma & "인 감마상관변환"
End If
End Sub
Private Sub mnu_Gray_Click()
f_Gray
End Sub
Private Sub f_Gray()
Picture2.Cls
For i = 0 To picHeight
For j = 0 To picWidth
pixel = Picture1.Point(j, i)
r = pixel And &HFF
g = (pixel And &HFF00&) / &H100
b = (pixel And &HFF0000) / &H10000
Gray = (r + g + b) \ 3
grayarr(j, i) = Gray
Picture2.PSet (j, i), RGB(Gray, Gray, Gray)
Next j
Next i
Label2.Caption = "그레이화"
End Sub


Private Sub mnu_HighBoost_Click()
f_HighBoost
End Sub
Private Sub f_HighBoost()
Picture8.Cls
Dim w As Long, alpa As Single
alpa = InputBox("얼마만큼 지원?")
w = 9 * alpa - 1
f_Gray
mask(0) = -1 / 9: mask(1) = -1 / 9: mask(2) = -1 / 9
mask(3) = -1 / 9: mask(4) = w / 9: mask(5) = -1 / 9
mask(6) = -1 / 9: mask(7) = -1 / 9: mask(8) = -1 / 9
For i = 1 To picHeight - 1
For j = 1 To picHeight - 1
n = 0: value = 0
For y = i - 1 To i + 1
For X = j - 1 To j + 1
value = value + (grayarr(X, y) * mask(n))
n = n + 1
Next X
Next y
If value > 255 Then value = 255
If value < 0 Then value = 0
Picture8.PSet (j, i), RGB(value, value, value)
Next j
Next i
Label8.Caption = "고주파지원필터"



End Sub

Private Sub mnu_HighPass_Click()
f_HighPass
End Sub
Private Sub f_HighPass()
Picture6.Cls
f_Gray
mask(0) = -1 / 9: mask(1) = -1 / 9: mask(2) = -1 / 9
mask(3) = -1 / 9: mask(4) = 8 / 9: mask(5) = -1 / 9
mask(6) = -1 / 9: mask(7) = -1 / 9: mask(8) = -1 / 9
For i = 1 To picHeight - 1
For j = 1 To picWidth - 1
n = 0: value = 0
For y = i - 1 To i + 1
For X = j - 1 To j + 1
value = value + (grayarr(X, y) * mask(n))
n = n + 1
Next X
Next y
If value > 255 Then value = 255
If value < 0 Then value = 0
Picture6.PSet (j, i), RGB(value, value, value)
Next j
Next i
Label6.Caption = "고주파통과필터1"
End Sub


Private Sub mnu_HighPass2_Click()
f_HighPass2
End Sub
Private Sub f_HighPass2()
Picture7.Cls
f_Gray
mask(0) = 0: mask(1) = -1: mask(2) = 0
mask(3) = -1: mask(4) = 5: mask(5) = -1
mask(6) = 0: mask(7) = -1: mask(8) = 0
For i = 1 To picHeight - 1
For j = 1 To picWidth - 1
n = 0: value = 0
For y = i - 1 To i + 1
For X = j - 1 To j + 1
value = value + (grayarr(X, y) * mask(n))
n = n + 1
Next X
Next y
If value > 255 Then value = 255
If value < 0 Then value = 0
Picture7.PSet (j, i), RGB(value, value, value)
Next j
Next i
Label7.Caption = "고주파통과필터2"
End Sub

Private Sub mnu_HighPass3_Click()
f_HighPass3
End Sub
Private Sub f_HighPass3()

Picture8.Cls
f_Gray
mask(0) = 1: mask(1) = -2: mask(2) = 1
mask(3) = -2: mask(4) = 5: mask(5) = -2
mask(6) = 1: mask(7) = -2: mask(8) = 1
For i = 1 To picHeight - 1
For j = 1 To picWidth - 1
n = 0: value = 0
For y = i - 1 To i + 1
For X = j - 1 To j + 1
value = value + (grayarr(X, y) * mask(n))
n = n + 1
Next X
Next y
If value > 255 Then value = 255
If value < 0 Then value = 0
Picture8.PSet (j, i), RGB(value, value, value)
Next j
Next i
Label8.Caption = "고주파통과필터"
End Sub




Private Sub mnu_LaplacianSD_Click()
f_LaplacianSD

End Sub
Private Sub f_LaplacianSD()
Picture8.Cls
f_Gray
mask(0) = 0: mask(1) = -1: mask(2) = 0
mask(3) = -1: mask(4) = 4: mask(5) = -1
mask(6) = 0: mask(7) = -1: mask(8) = 0

For i = 1 To picHeight - 1
For j = 1 To picWidth - 1
n = 0: value = 0
For y = i - 1 To i + 1
For X = j - 1 To j + 1
value = value + (grayarr(X, y) * mask(n))
n = n + 1
Next X
Next y
If value > 255 Then value = 255
If value < 0 Then value = 0
Picture8.PSet (j, i), RGB(value, value, value)
Next j
Next i
Label8.Caption = "라플라시안"

End Sub
Private Sub mnu_LoadPicture_Click()
CDialog.ShowOpen
Picture1.Picture = LoadPicture(CDialog.FileName)
Label1.Caption = "원래영상"

End Sub

Private Sub mnu_MaskSharpning_Click()
f_MaskSharpning
End Sub
Private Sub f_MaskSharpning()
Picture5.Cls
f_Gray
mask(0) = -1: mask(1) = -1: mask(2) = -1
mask(3) = -1: mask(4) = 9: mask(5) = -1
mask(6) = -1: mask(7) = -1: mask(8) = -1
For i = 1 To picHeight - 1
For j = 1 To picWidth - 1
n = 0: value = 0
For y = i - 1 To i + 1
For X = j - 1 To j + 1
value = value + (grayarr(X, y) * mask(n))
n = n + 1
Next X
Next y
If value > 255 Then value = 255
If value < 0 Then value = 0
Picture5.PSet (j, i), RGB(value, value, value)
Next j
Next i
Label5.Caption = "마스크이용 샤프닝"


End Sub

Private Sub mnu_MaxMinThresholdEdge_Click()
f_MaxMinThresholdEdge
End Sub
Private Sub f_MaxMinThresholdEdge()
Dim threHigh, threLow As Long
Picture6.Cls
threHigh = InputBox("최대 임계값?")
threLow = InputBox("최소 임계값?")
f_SimilarEdge
For i = 1 To picHeight - 1
For j = 1 To picWidth - 1
pixel = Picture3.Point(j, i)

r = pixel And &HFF
g = (pixel And &HFF00&) \ &H100
b = (pixel And &HFF0000) \ &H10000

Gray = (r + g + b) \ 3
If Gray > threHigh Then
value = 255
ElseIf Gray < threLow Then
value = 0
Else
value = Gray
End If
Picture6.PSet (j, i), RGB(value, value, value)
Next j
Next i
Label6.Caption = "최대/최소임계값이용"




End Sub

Private Sub mnu_Mirroring_Click()
f_Mirroring
End Sub

Private Sub mnu_Multiplication_Click()
f_Multiplication
End Sub
Private Sub f_Multiplication()
Picture5.Cls
value = InputBox("값입력?")
f_Gray
For i = 0 To picHeight
For j = 0 To picWidth
pixel = grayarr(j, i)
result = pixel * value
If result > 255 Then result = 255
If result < 0 Then result = 0
Picture5.PSet (j, i), RGB(result, result, result)
Next j
Next i
Label5.Caption = "곱셈연산"

End Sub

Private Sub mnu_Parabola_Click()
f_Parabola
End Sub
Private Sub f_Parabola()
Picture3.Cls: Picture4.Cls
Dim lut1(256), lut2(256) As Double
For i = 0 To 255
lut1(i) = 255 - 255 * (i / 128 - 1) ^ 2
lut2(i) = 255 * (i / 128 - 1) ^ 2
Next i
For i = 0 To picHeight
For j = 0 To picWidth
pixel = Picture1.Point(j, i)
r = pixel And &HFF
g = (pixel And &HFF00&) \ &H100
b = (pixel And &HFF0000) \ &H10000
r1 = lut1(r)
g1 = lut1(g)
b1 = lut1(b)
Picture3.PSet (j, i), RGB(r1, g1, b1)
r2 = lut2(r)
g2 = lut2(g)
b2 = lut2(b)
Picture4.PSet (j, i), RGB(r2, g2, b2)
Next j
Next i
Label3.Caption = "파라볼라변환(1)"
Label4.Caption = "파라볼라변환(2)"
End Sub
Private Sub mnu_Posterizing_Click()
f_Posterizing
End Sub
Private Sub f_Posterizing()
Picture6.Cls
For i = 0 To picHeight
For j = 0 To picWidth
pixel = Picture1.Point(j, i)
r = pixel And &HFF
g = (pixel And &HFF00&) \ &H100
b = (pixel And &HFF0000) \ &H10000
Picture6.PSet (j, i), RGB((r \ 32) * 32, (g \ 32) * 32, (b \ 32) * 32)
Next j
Next i
Label6.Caption = "8단계 포스트라이징"
End Sub
Private Sub mnu_Reverse_Click()
f_Reverse
End Sub
Private Sub f_Reverse()
Picture2.Cls
For i = 0 To picHeight
For j = 0 To picWidth
pixel = Picture1.Point(j, i)
r = (pixel And &HFF)
g = (pixel And &HFF00&) / &H100
b = (pixel And &HFF0000) / &H10000
r_result = -r + 255
g_result = -g + 255
b_result = -b + 255
Picture2.PSet (j, i), RGB(r_result, g_result, b_result)
Next j
Next i
Label2.Caption = "역변환"
End Sub

Private Sub mnu_SimilarEdge_Click()
f_SimilarEdge
End Sub
Private Sub f_SimilarEdge()
Picture3.Cls
f_Gray
For i = 1 To picHeight - 1
For j = 1 To picWidth - 1
max = 0
For y = i - 1 To i + 1
For X = j - 1 To j + 1
value = Abs(grayarr(j, i) - grayarr(X, y))
If max < value Then max = value
Next X
Next y
Picture3.PSet (j, i), RGB(max, max, max)
Next j
Next i
Label3.Caption = "유사연산자"


End Sub

Private Sub mnu_SobelFD_Click()
f_SobelFD
End Sub
Private Sub f_SobelFD()
Picture5.Cls: Picture6.Cls: Picture7.Cls
Dim x1 As Long, y1 As Long
f_Gray
x0(0) = -1: x0(1) = -2: x0(2) = -1
x0(3) = 0: x0(4) = 0: x0(5) = 0
x0(6) = 1: x0(7) = 2: x0(8) = 1

y0(0) = 1: y0(1) = 0: y0(2) = -1
y0(3) = 2: y0(4) = 0: y0(5) = -2
y0(6) = 1: y0(7) = 0: y0(8) = -1
For i = 1 To picHeight - 1
For j = 1 To picWidth - 1
n = 0: x1 = 0: y1 = 0

For y = i - 1 To i + 1
For X = j - 1 To j + 1
x1 = x1 + (x0(n) * grayarr(X, y))
y1 = x1 + (y0(n) * grayarr(X, y))
n = n + 1
Next X
Next y
If x1 < 0 Then x1 = 0
Picture5.PSet (j, i), RGB(x1, x1, x1)
If y1 < 0 Then y1 = 0
Picture6.PSet (j, i), RGB(y1, y1, y1)
value = Sqr(x1 ^ 2 + y1 ^ 2)
Picture7.PSet (j, i), RGB(value, value, value)
Next j
Next i
Label5.Caption = "수평소벨"
Label6.Caption = "수직소벨"
Label7.Caption = "수평수직소벨"


End Sub

Private Sub mnu_Solarizing_Click()
f_Solarizing
End Sub
Private Sub f_Solarizing()
Picture7.Cls
threshold = InputBox("임계값을 넣으시오")
For i = 0 To picHeight
For j = 0 To picWidth
pixel = Picture1.Point(j, i)
r = pixel And &HFF
g = (pixel And &HFF00&) \ &H100
b = (pixel And &HFF0000) \ &H10000
If r > threshold Then r = 255 - r: If r < 0 Then r = 0
If g > threshold Then g = 255 - g: If g < 0 Then g = 0
If b > threshold Then b = 255 - b: If b < 0 Then b = 0
Picture7.PSet (j, i), RGB(r, g, b)
Next j
Next i
Label7.Caption = "임계값" & threshold & "의 솔라이징"

End Sub

Private Sub mnu_Subtraction_Click()
f_Subtraction
End Sub
Private Sub f_Subtraction()
Picture4.Cls
value = InputBox("값입력?")
f_Gray
For i = 0 To picHeight
For j = 0 To picWidth
pixel = grayarr(j, i)
result = pixel - value
If result < 0 Then result = 0
Picture4.PSet (j, i), RGB(result, result, result)
Next j
Next i
Label4.Caption = "뺄셈연산"
End Sub

Private Sub mnu_SubtractionEdge_Click()
f_SubtractionEdge
End Sub
Private Sub f_SubtractionEdge()
Picture4.Cls
f_Gray
For i = 1 To picHeight - 1
For j = 1 To picWidth - 1
max = 0
value = Abs(grayarr(j - 1, i - 1) - grayarr(j + 1, i + 1))
If value > max Then max = value
value = Abs(grayarr(j, i - 1) - grayarr(j, i + 1))
If value > max Then max = value
value = Abs(grayarr(j + 1, i - 1) - grayarr(j - 1, i + 1))
If value > max Then max = value
value = Abs(grayarr(j + 1, i) - grayarr(j - 1, i))
If value > max Then max = value
Picture4.PSet (j, i), RGB(max, max, max)
Next j
Next i
Label4.Caption = "차연산자"
End Sub

Private Sub mnu_ThresholdEdge_Click()
f_ThresholdEdge
End Sub
Private Sub f_ThresholdEdge()
Dim thre As Long
Picture5.Cls
thre = InputBox("임계값?")
f_SimilarEdge

For i = 1 To picHeight - 1
For j = 1 To picWidth - 1
pixel = Picture3.Point(j, i)
r = (pixel And &HFF)
g = (pixel And &HFF00&) / &H100
b = (pixel And &HFF0000) / &H10000

Gray = (r + g + b) \ 3

If Gray > thre Then
value = 255
Else
value = 0
End If
Picture5.PSet (j, i), RGB(value, value, value)
Next j
Next i
Label5.Caption = "임계값 이용"


End Sub

Private Sub mnu_UnSharpning_Click()
f_UnSharpning
End Sub
Private Sub f_UnSharpning()
Picture7.Cls
f_Blurring
For i = 1 To picHeight - 1
For j = 1 To picWidth - 1
value = grayarr(j, i) - blur(j, i)
If value > 255 Then value = 255
If value < 0 Then value = 0
Picture7.PSet (j, i), RGB(value, value, value)
Next j
Next i
Label7.Caption = "UnSharpning Masking"
End Sub

Private Sub mnu_XOR_Click()
f_XOR
End Sub

Private Sub f_XOR()
Picture7.Cls
value = InputBox("값입력(0~255)?")
f_Gray
For i = 0 To picHeight
For j = 0 To picWidth
pixel = grayarr(j, i)
result = pixel Xor value
Picture7.PSet (j, i), RGB(result, result, result)
Next j
Next i
Label7.Caption = value & "로 XOR연산"

End Sub

Private Sub f_Mirroring()
 Picture2.Cls: Picture3.Cls: Picture4.Cls:
    For i = 0 To picHeight
        For j = 0 To picWidth
        pixel = Picture1.Point(j, i)
        r = (pixel And &HFF)
        g = (pixel And &HFF00&) / &H100
        b = (pixel And &HFF0000) / &H10000
        Picture2.PSet (picWidth - j, i), RGB(r, g, b)
        Picture3.PSet (j, picHeight - i), RGB(r, g, b)
        Picture4.PSet (picWidth - j, picHeight - i), RGB(r, g, b)
        Next j
        Next i
        Label2.Caption = "수직"
        Label3.Caption = "수평"
        Label4.Caption = "수직/수평"

End Sub

