Attribute VB_Name = "ModuleStart"

Public Sub StartProcessFile()
Dim Totale As String
Dim s        As Long
Dim Corrente As Long
Dim Nome2    As String
Dim Num      As String
Dim Ext()    As String

On Error GoTo 100 'Gestione errori

Progress.Show

Totale = 0
Totale = Val(Form1.Number.Caption) + Form1.Numero.Text
'Print Len(Totale)
'Print Totale
If Len(Totale) <= Form1.InsertZero.Text Then GoTo 50
If Len(Totale) > Form1.InsertZero.Text Then Errore.Show 1: GoTo 500


50
Form1.Command1.Enabled = False
Form1.Command2.Enabled = False
Form1.Command3.Enabled = False
Form1.Command4.Enabled = False
Progress.Barra.Width = 15
J = -1
b = 0
b = Progress.Bordo.Width \ (Form1.Grid.Rows - 1)

Corrente = Val(Form1.Numero.Text)


For s = 1 To Val(Form1.Number.Caption)
    Num = Corrente
    Nome2 = Left(String(Val(Form1.InsertZero.Text), "0"), Len(String(Val(Form1.InsertZero.Text), "0")) - Len(Num)) & Num
    
    'cambia il nome
    Ext = Split(Form1.Grid.TextMatrix(s, 1), ".", -1, vbTextCompare)
    If Form1.Check1.Value = 1 Then
        If Form1.Option1.Value = True Then Nome2 = Nome2 & Form1.Matrix.Text & "." & Ext(UBound(Ext))
        If Form1.Option2.Value = True Then Nome2 = Form1.Matrix.Text & Nome2 & "." & Ext(UBound(Ext))
    Else
    'non cambia ma aggiunge
        Nome2 = Nome2 & Form1.Matrix.Text & Form1.Grid.TextMatrix(s, 1)
    End If
    'Aggiunge 1
    Form1.Numero.Text = Form1.Numero.Text + 1
    'Copia dei file
    Dim fin As String
    fin = Form1.Grid.TextMatrix(s, 0) + Form1.Grid.TextMatrix(s, 1)
    FileCopy fin, Form1.DestDir.Text & Nome2
    Corrente = Corrente + 1
    'Scorrimento barra
    Progress.curNumProg.Caption = s & " \ " & Form1.Number.Caption
    Progress.curFileProg.Caption = Form1.Grid.TextMatrix(s, 1) & " --> " & Nome2
    Progress.Barra.Width = Progress.Barra.Width + b
    If Progress.Barra.Width > Progress.Bordo.Width Then Progress.Barra.Width = Progress.Bordo.Width
    If Progress.Barra.Width = Progress.Bordo.Width Then Progress.Barra.Visible = False: Progress.Bordo.Visible = False
Next s

Unload Progress

'Cancellazione file (se abilitata)
If Form1.Check7.Value = 1 Then Delete.Show 1

GoTo 200


100 'Generato un errore
MsgBox "Copy file error! Please check the new position of the files and try again.", vbApplicationModal, "COPY ERROR"


200
Form1.Command1.Enabled = True
Form1.Command2.Enabled = True
Form1.Command3.Enabled = True
Form1.Command4.Enabled = True
Progress.Barra.Width = 15
Form1.Dir1.Refresh
Form1.Dir2.Refresh
Form1.File1.Refresh
Form1.File2.Refresh

500

Unload Progress

End Sub


