# Sistema-Aseguradora
'Aqui es donde se lleva acabo el login del usuario mediante su contraseña y su usuario registrado en la base de datos'

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

'Aqui es donde se lleva a cabo el formulario que recaba los datos del usuario para poder darte la cotizacion del seguro y todos los datos que ingresaste en una tabla.

Private Sub Guardar_Click()
Range("A2:N2").Insert
Range("A2") = Val(Poliza)
Range("B2") = NombreA
Range("C2") = Val(EdadA)
Range("D2") = SeguS
Range("E2") = "=vlookup(D2,DM,2,0)"

'Aqui es la parte donde se registran los empleados y se calcula la nomina

Private Sub calcular_Click()
If puesto <> "" And Resgitro <> "" And nombre <> "" Then

    Hoja2.Range("j2") = puesto
    sueldo_q = Hoja2.Range("k2")
    imss = sueldo_q * 0.05
    afore = sueldo_q * 0.85
    infona = sueldo_q * 0.03
    sueldo = sueldo_q - imss - afore - infona
   
Range("a2:I2").Insert
Range("a2") = Registro
Range("b2") = nombre
Range("c2") = puesto
Range("d2") = departamento
Range("e2") = salario
Range("f2") = imss
Range("g2") = afore
Range("h2") = infona
Range("i2") = sueldo
 End If
 
 'este es el codigo que te muestra la ventana de empleados y la de venta, tambien se regresan a la anterior de ser requeridos'
 
Private Sub Trabajadores_Click()
MsgBox "Trabajadores"
   opciones.Hide
   Empleados.Show 
   
   Private Sub seguro_Click()
MsgBox "Seguros"
   opciones.Hide
   Venta.Show
End Sub
 
 'Salir 
'Mensaje por si el usuario no quiere salir
s = MsgBox("¿Seguro que quieres salir?", vbOKCancel, "Salir del programa")
If s = vbox Then
'Almacenado Documento
ThisWorkbook.Save
'Saliendo Excel
Application.Quit
 
 
 
 
 
 
 
 
 
 
