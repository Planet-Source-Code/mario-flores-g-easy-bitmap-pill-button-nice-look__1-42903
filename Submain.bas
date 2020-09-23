Attribute VB_Name = "Module1"
Option Explicit
Dim i

Sub Main()
    
    For i = 1 To Forms.Count
        Forms(i).ScaleMode = vbPixels     'Api's Trabajan solo con Pixeles
    Next i

    Form1.Show 'Iniciamos El programa
End Sub
