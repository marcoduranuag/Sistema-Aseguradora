# Sistema-Aseguradora
Private Sub Validar_Click()
'Validar Usuarios
If login <> "" And Password <> "" Then
Hoja2.Range("D2") = login
If Val(Password) = Hoja2.Range("E2") Then
   MsgBox "Bienvenido " & login
   Ingresar.Hide
   opciones.Show
Else
    MsgBox "Usuario o contraseña incorrectos"
   End If
Else
  MsgBox "Faltan Datos"
  End If
  

End Sub
