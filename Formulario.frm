VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Bienvenido"
   ClientHeight    =   7512
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7344
   OleObjectBlob   =   "Formulario.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    'Boton registrar'
    'se desactivan los botones de modificar, borrar y buscar'
    Frame1.Visible = True
    CommandButton2.Enabled = False
    CommandButton3.Enabled = False
    CommandButton4.Enabled = False
    
End Sub

Private Sub CommandButton10_Click()
    CommandButton11.Enabled = False
    CommandButton9.Enabled = True
    
    Frame3.Visible = False
    CommandButton1.Enabled = True
    CommandButton2.Enabled = True
    CommandButton4.Enabled = True
    
    TextBox10.Text = ""
    TextBox11.Text = ""
    TextBox12.Text = ""
    TextBox13.Text = ""
    TextBox14.Text = ""
End Sub

Private Sub CommandButton11_Click()
    'Modificar datos'
    'btn modificar'
    Dim numfil As Integer
    Dim opcion As Integer
    opcion = MsgBox("¿Esta seguro que quiere modificar los datos?", vbYesNo, "Modificar usuario")
    'MsgBox opcion'
    
    If opcion = 6 Then
        CommandButton11.Enabled = False
        CommandButton9.Enabled = True
        numfil = Val(TextBox10.Text)
        Hoja1.Cells(numfil + 2, 2) = TextBox11.Value
        Hoja1.Cells(numfil + 2, 3) = TextBox12.Value
        Hoja1.Cells(numfil + 2, 4) = TextBox13.Value
        Hoja1.Cells(numfil + 2, 5) = TextBox14.Value
        
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox13.Text = ""
        TextBox14.Text = ""
    End If
End Sub

Private Sub CommandButton2_Click()
    'boton buscar'
    'Se desactivan btnes registrar,modificar y borrar  al pulsar buscar usuario'
    Frame2.Visible = True
    CommandButton1.Enabled = False
    CommandButton3.Enabled = False
    CommandButton4.Enabled = False
    
End Sub

Private Sub CommandButton3_Click()
    'boton modificar'
    'Se desactivan botones registrar, buscar, eliminar'
    Frame3.Visible = True
    CommandButton1.Enabled = False
    CommandButton2.Enabled = False
    CommandButton4.Enabled = False
End Sub

Private Sub CommandButton5_Click()
    'Registrar usuarios'
    'se activan los botones de modificar, borrar y buscar'
    Frame1.Visible = False
    CommandButton2.Enabled = True
    CommandButton3.Enabled = True
    CommandButton4.Enabled = True
    
    TextBox1.Text = ""
    TextBox2.Text = ""
    TextBox3.Text = ""
    TextBox4.Text = ""
End Sub

Private Sub CommandButton6_Click()
    'Registrar usuario'
    'btn aceptar'
    Dim nombre As String
    Dim apellido As String
    Dim telefono As String
    Dim email As String
    Dim id As Integer
    Dim fila As Integer
    
    nombre = TextBox1.Value
    apellido = TextBox2.Value
    telefono = TextBox3.Value
    email = TextBox4.Value
    
    'Linea para que devuelva la ultima fila'
    fila = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row
    'MsgBox fila'
    'Al id le resto -1 porque me devuelve el numero de la celda vacia y para ir sumando '
    'como id se resta -1'
    id = fila - 1
    'MsgBox id'
    Hoja1.Cells(fila + 1, 1) = id
    Hoja1.Cells(fila + 1, 2) = nombre
    Hoja1.Cells(fila + 1, 3) = apellido
    Hoja1.Cells(fila + 1, 4) = telefono
    Hoja1.Cells(fila + 1, 5) = email
    
    'Con esto se borra los cuadros de texto cada vez que da click'
    TextBox1.Text = ""
    TextBox2.Text = ""
    TextBox3.Text = ""
    TextBox4.Text = ""
    
End Sub

Private Sub CommandButton7_Click()
    'buscar usuario'
    'btn buscar'
    Dim numfil As Integer
    'LLamamos a la funcion para obtener el numero de filas que no estan vacias'
    numfil = numfila()
    'MsgBox numfil'
    'val()-> cambia de tipo del contenido del textbox5'
    If Val(TextBox5.Text) >= 1 And Val(TextBox5.Text) <= numfil Then
        TextBox6.Text = Hoja1.Cells(Val(TextBox5.Text) + 2, 2)
        TextBox7.Text = Hoja1.Cells(Val(TextBox5.Text) + 2, 3)
        TextBox8.Text = Hoja1.Cells(Val(TextBox5.Text) + 2, 4)
        TextBox9.Text = Hoja1.Cells(Val(TextBox5.Text) + 2, 5)
     Else
        MsgBox "No se encuentran los datos", vbExclamation, "Error"
    End If
End Sub

Private Sub CommandButton8_Click()
    'Buscar usuario'
    ' btn volver, se activan bnt principales'
    Frame2.Visible = False
    CommandButton1.Enabled = True
    CommandButton3.Enabled = True
    CommandButton4.Enabled = True
    
    'lineas para dejar en blanco los textbox cuando se da a volver'
    TextBox5.Text = ""
    TextBox6.Text = ""
    TextBox7.Text = ""
    TextBox8.Text = ""
    TextBox9.Text = ""
End Sub

Private Function numfila() As Integer
    'Funcion que cuenta las celdas que estan ocupadas en fila es decir cuantas filas hay'
    Dim i As Integer
    i = 3
    
    Do While Hoja1.Cells(i, 2) <> ""
        i = i + 1
    Loop
    
    numfila = i - 3
End Function

Private Sub CommandButton9_Click()
    'Modificar datos'
    Dim numfil As Integer
    CommandButton11.Enabled = True
    CommandButton9.Enabled = False

    'LLamamos a la funcion para obtener el numero de filas que no estan vacias'
    numfil = numfila()
    'val()-> cambia de tipo del contenido del textbox5'
    If Val(TextBox10.Text) >= 1 And Val(TextBox10.Text) <= numfil Then
        TextBox11.Text = Hoja1.Cells(Val(TextBox10.Text) + 2, 2)
        TextBox12.Text = Hoja1.Cells(Val(TextBox10.Text) + 2, 3)
        TextBox13.Text = Hoja1.Cells(Val(TextBox10.Text) + 2, 4)
        TextBox14.Text = Hoja1.Cells(Val(TextBox10.Text) + 2, 5)
     Else
        MsgBox "No se encuentran los datos", vbExclamation, "Error"
        CommandButton9.Enabled = True
        CommandButton11.Enabled = False
    End If
End Sub
