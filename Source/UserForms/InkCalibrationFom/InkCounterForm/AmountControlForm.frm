VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "INK DOSAGE COUNTER"
   ClientHeight    =   11388
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   12408
   OleObjectBlob   =   "AmountControlForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CheckBox1_Click()
'call llamartramas
End Sub

Private Sub CheckBox3_Click()

'call llamartramas
End Sub

Private Sub CheckBox39_Click()

'call llamartramas
End Sub

Private Sub CheckBox82_Click()
'call llamartramas

End Sub


'Expand counter
Private Sub CommandButton1_Click()

MsgBox ("Esta agregando un rango de tinta magenta que puede causar sobreestimación, evite cantidades mayores a 40 gramos y consulte al diseñador")
InputBox ("Indique cuantas dosis de tinta magenta le faltan para completar la cantidad sugerida")



End Sub


'Expand counter
Private Sub CommandButton2_Click()
MsgBox ("Esta agregando un rango de tinta amarilla que puede causar sobreestimación, evite cantidades mayores a 90 gramos y consulte al diseñador")
InputBox ("Indique cuantas dosis de tinta amarilla le faltan para completar la cantidad sugerida")
End Sub



'Open Ink Dosage form
Private Sub CommandButton3_Click()

UserForm4.Show


End Sub


'Suggest scoop amounts
Private Sub CommandButton4_Click()

TextBox5.Value = TextBox1.Value / 2
TextBox6.Value = TextBox2.Value / 2
TextBox7.Value = TextBox3.Value / 2
TextBox8.Value = TextBox4.Value / 2



End Sub

Private Sub Label1_Click()

'call Macro1

End Sub

Private Sub Label10_Click()

'call Macro1

End Sub

Private Sub Label11_Click()

'call Macro1

End Sub

Private Sub Label12_Click()

'call Macro1

End Sub

Private Sub Label4_Click()

'call Macro1

End Sub


