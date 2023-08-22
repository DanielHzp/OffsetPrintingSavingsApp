VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "Calibrate Ink Input and Printing Coverage"
   ClientHeight    =   7032
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12120
   OleObjectBlob   =   "InkCalibrationForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Calculate optimal amounts of ink for a production batch
Private Sub CommandButton1_Click()



Dim cantImp As Integer, i As Integer


'Select the starting cell to read input printing parameters from the worksheet

For i = 1 To 10000
   If Range("L1:L" & 10000).Cells(i, 1).Text = "Tamaño de las muestras de inspección de calidad" Then
    Range("L" & i).Select
    Exit For
    End If
    Next i
    
    'Input data will be read moving from left to right on the worksheet
    'Each column from L to A represent a user input
    'First value to fetch is the amount of sheets to print per unit of time
    For i = 1 To 10000
    If ActiveCell.Offset(i, 0).Value = "" Then
    
    cantImp = ActiveCell.Offset(i - 1, 0).Value
    Exit For
    End If
    Next i
    
    'Calculate optimal or suggested ink amounts in grams
    'This logic is standard for the press printing machine used
    If TextBox39.Value > 0.64 And TextBox39.Value < 0.8 Then
    TextBox9.Text = "9.5 a 13.5 gr"
    Else
    
    TextBox9.Value = TextBox39.Value * 2 * 0.28 * 0.23 * cantImp * 0.16666
    
    End If


    If TextBox38.Value > 0.4 And TextBox38.Value < 0.6 Then
    TextBox10.Text = "17 a 36.5 gr"
    Else
    
    TextBox10.Value = TextBox38.Value * 2 * 0.28 * 0.23 * cantImp * 0.833333
    
    End If


    If TextBox37.Value > 0.3 And TextBox37.Value < 0.5 Then
    TextBox11.Text = "45 a 90 gr"
    Else
    
    TextBox11.Value = TextBox37.Value * 2 * 0.28 * 0.23 * cantImp * 1
    
    End If
    


    If TextBox36.Value > 0.34 And TextBox36.Value < 0.5 Then

    TextBox12.Text = "45 a 73 gr"
    Else
    
    TextBox12.Value = TextBox36.Value * 2 * 0.28 * 0.23 * cantImp * 0.5
    
    End If



End Sub



'Calculate optimal coverage in percentage amounts
Private Sub CommandButton2_Click()


Dim i As Integer, muestraCalidad As Integer


For i = 1 To 10000
    
    
    'Input data will be read moving from left to right on the worksheet
    'Each column from L to A represent a user input
    'First value to fetch is the amount of sheets to print per unit of time
   If Range("L1:L" & 10000).Cells(i, 1).Text = "Tamaño de las muestras de inspección de calidad" Then
    Range("L" & i).Select
    Exit For
    End If
    Next i
    
    'To estimate coverage the batch size is required as a user input
    For i = 1 To 10000
    
    If ActiveCell.Offset(i, 0).Value = "" Then
    
    muestraCalidad = ActiveCell.Offset(i - 1, 0).Value
    
    Exit For
    
    End If
    Next i
    
    
If muestraCalidad = 250 Then
    TextBox28.Text = "30% a 50%"
    TextBox27.Text = "35% a 50%"
    TextBox30.Text = "<<60%"
    TextBox29.Text = "<<35%"
End If
If muestraCalidad = 500 Then
    TextBox28.Text = ">>50%"
    TextBox27.Text = ">>50%"
    TextBox30.Text = "<<60%"
    TextBox29.Text = "40% a 60%"
End If
If muestraCalidad = 1000 Then
    TextBox28.Text = ">>60%"
    TextBox27.Text = ">>60%"
    TextBox30.Text = "65% a 80%"
    TextBox29.Text = ">>60%"
End If



End Sub



'The code below executes the screen calibration logic
Private Sub CommandButton3_Click()
Dim i As Integer, alto As Double, ancho As Double, rend As Double, numimp As Integer


    For i = 1 To 10000
    
    'Input data will be read moving from left to right on the worksheet
    'Each column from L to A represents a user input
    'Worksheet columns are iterated from right to left
   If Range("C1:C" & 10000).Cells(i, 1).Text = "Cantidad de impresiones" Then
    Range("C" & i).Select
    Exit For
    End If
    Next i


'In order to calibrate the screens and confirm the printing parameters are optimal, it is required to read the printing dimensions and batch size
For i = 1 To 10000

    If ActiveCell.Offset(i, 0).Value = "" Then
    
    ActiveCell.Offset(i - 1, 0).Select
    
    'Input parameters are fixed columns in the worksheet
    numimp = ActiveCell.Value
    ancho = ActiveCell.Offset(0, 2).Value
    alto = ActiveCell.Offset(0, 3).Value
    rend = ActiveCell.Offset(0, 6).Value
    Exit For
    End If
    Next i
    
    
    'IDEALLY STANDARD AND CONSTANT VALIDATION VALUES MUST BE PARAMETRIZED
    
  'CYAN CALIBRATION LOGIC AND VALIDATION
    TextBox35.Value = (TextBox16.Value) / (TextBox39.Value * numimp * ancho * alto * rend)
    
    If TextBox35.Value > 0.186 Then
    TextBox26.Text = "Tonalidad muy oscura"
    ElseIf TextBox35.Value < 0.146 Then
    TextBox26.Text = "Tonalidad muy clara"
    Else
    TextBox26.Text = "OK, COMENZAR IMPRESIÓN"
    End If
    
    
    
    'MAGENTA CALIBRATION LOGIC AND VALIDATIONS
    TextBox34.Value = (TextBox15.Value) / (TextBox38.Value * numimp * ancho * alto * rend)
    
    If TextBox34.Value > 0.8533 Then
    TextBox25.Text = "Tonalidad muy oscura"
    ElseIf TextBox34.Value < 0.8133 Then
    TextBox25.Text = "Tonalidad muy clara"
    Else
    TextBox25.Text = "OK, COMENZAR IMPRESIÓN"
    End If
    
    
    
   'YELLOW CALIBRATION LOGIC AND VALIDATIONS
     TextBox33.Value = (TextBox14.Value) / (TextBox37.Value * numimp * ancho * alto * rend)
     
    If TextBox33.Value > 1.2 Then
    TextBox24.Text = "Tonalidad muy oscura"
    ElseIf TextBox33.Value < 0.98 Then
    TextBox24.Text = "Tonalidad muy clara"
    Else
    TextBox24.Text = "OK, COMENZAR IMPRESIÓN"
    End If
    
    
    
    'BLACK CALIBRATION LOGIC AND VALIDATIONS
     TextBox32.Value = (TextBox13.Value) / (TextBox36.Value * numimp * ancho * alto * rend)
     
    If TextBox32.Value > 0.52 Then
    TextBox31.Text = "Tonalidad muy oscura"
    ElseIf TextBox32.Value < 0.48 Then
    TextBox31.Text = "Tonalidad muy clara"
    Else
    TextBox31.Text = "OK, COMENZAR IMPRESIÓN"
    End If



End Sub



'Export production cycle data output on the second worksheet
Private Sub CommandButton4_Click()
Dim i As Integer
Dim tintacyan As Double, cobertcyan As Double, tramacyan As Double, tintamag As Integer, cobertmag As Double, tintaamar As Double, cobertamar As Double, tramaamar As Double, tintanegro As Double, cobertnegro As Double, tramanegro As Double


    For i = 1 To 10000
    
   If Worksheets("Sheet2").Range("A1:A" & 10000).Cells(i, 1).Text = "Gr tinta" Then
   
    Worksheets("Sheet2").Range("A" & i).Select
    
    Exit For
    End If
    Next i
    
    
    
    For i = 1 To 10000
 
    If ActiveCell.Offset(i, 0).Value = "" Then
    
    ActiveCell.Offset(i, 0).Select
    ActiveCell.Value = TextBox16.Value
    ActiveCell.Offset(0, 1) = TextBox20.Value
    ActiveCell.Offset(0, 2) = TextBox1.Value
    
    
    Exit For
    End If
    Next i
    
    
    
    
End Sub


'Open step-by-step instructions to begin production cycle
Private Sub CommandButton5_Click()

UserForm3.Show

End Sub

Private Sub TextBox33_Change()

End Sub
