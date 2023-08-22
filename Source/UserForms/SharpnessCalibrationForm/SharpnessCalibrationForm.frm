VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm5 
   Caption         =   "CALIBRATE LUMINOUS INTENSITY OF PRINTING COLORS"
   ClientHeight    =   9624.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12384
   OleObjectBlob   =   "SharpnessCalibrationForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Calculate standard optimal values for the color intensities of the production batch
'Certain range values must be avoided for the type of industrial printer used
Private Sub CommandButton1_Click()


'Fixed value for standard printing surface type
If Range("G12").Text = "Propalcote" Then
    TextBox3.Text = "5%"
    TextBox12.Text = "10%"
    TextBox14.Text = ">>15%"
    
    TextBox4.Text = "10%"
    TextBox13.Text = "15%"
    TextBox17.Text = "1%,5%"
    
    TextBox1.Text = "15%"
    TextBox11.Text = "20%"
    TextBox16.Text = "1%"
    
    TextBox2.Text = "20%"
    TextBox10.Text = "10%"
    TextBox15.Text = "1%,5%"
End If
    

End Sub


'Verifies if the sharpness result is optimal for the production cycle
Private Sub CommandButton2_Click()


Dim promTotal As Double
Dim contraste As Double



'CDBL function converts values to doubles
'Calculate sharpness parameters with the following equations
TextBox18.Value = (WorksheetFunction.Log10(1 / TextBox9.Value) + WorksheetFunction.Log10(1 / TextBox6.Value) + WorksheetFunction.Log10(1 / TextBox7.Value) + WorksheetFunction.Log10(1 / TextBox8.Value)) / 4

contraste = (CDbl(TextBox9.Value) + CDbl(TextBox6.Value) + CDbl(TextBox7.Value) + CDbl(TextBox8.Value)) / 4
TextBox19.Value = contraste

TextBox21.Value = (WorksheetFunction.Log10(1 / TextBox9.Value) + WorksheetFunction.Log10(1 / TextBox6.Value)) / (WorksheetFunction.Log10(1 / TextBox9.Value) + WorksheetFunction.Log10(1 / TextBox6.Value) + WorksheetFunction.Log10(1 / TextBox7.Value) + WorksheetFunction.Log10(1 / TextBox8.Value))

TextBox20.Value = (CDbl(TextBox9.Value) + CDbl(TextBox6.Value)) / (CDbl(TextBox9.Value) + CDbl(TextBox6.Value) + CDbl(TextBox7.Value) + CDbl(TextBox8.Value))

TextBox22.Value = (CDbl(TextBox9.Value) + CDbl(TextBox6.Value)) / (CDbl(TextBox9.Value) + CDbl(TextBox6.Value) + CDbl(TextBox7.Value))

'Sharpness average output
promTotal = (CDbl(TextBox18.Value) + CDbl(TextBox19.Value) + CDbl(TextBox20.Value) + CDbl(TextBox21.Value) + CDbl(TextBox22.Value)) / 5


'Sharpness level output
If promTotal < 0.2 Then
    TextBox24.Text = "BAJA NITIDEZ, AUMENTAR LUMINOSIDADES"
ElseIf promTotal > 0.6 Then
    TextBox24.Text = "NITIDEZ MUY SATURADA, DISMINUIR LUMINOSIDADES"
    Else
    TextBox23.Text = "OK, IMPRESIÓN OFFSET"
    End If
    

End Sub

Private Sub Image2_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

'call Macro1

End Sub








'Inactive form controls
Private Sub Label13_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label9_Click()

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub TextBox6_Change()

End Sub

Private Sub TextBox9_Change()

End Sub
