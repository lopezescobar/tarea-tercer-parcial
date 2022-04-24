Module Module1
    Public cliente(7) As Integer
    Public placa(7) As Integer
    Public carros(7) As Integer
    Public valor(7) As Integer
    Public total(7) As Byte
    Public tipo(7) As Integer
    Public kilometrajeinicial As Integer
    Public kilometrajefinal As Integer
    Public cobrobase As Integer

    Public totalcarro As Integer = 0
    Public Const tipo1 = 500
    Public Const tipo2 = 400
    Public Const tipo3 = 300
    Public Const tipo4 = 200
    Public fila As Byte = 0
    Public I As Byte

    Sub calcular()
        If (totalcarro < 7) Then
            cliente(totalcarro) = Form1.TextBox1.Text
            placa(totalcarro) = Form1.TextBox2.Text
        End If

        If tipo(totalcarro) < 12 Then
            total(totalcarro) = (valor(totalcarro) * 3) + valor(totalcarro)
            tipo = tipo1 + ((valor(total) * 3) + valor(totalcarro))



        ElseIf tipo(totalcarro) >= 12 And tipo(totalcarro) <= 24 Then
            total(totalcarro) = (valor(totalcarro) * 4) + valor(totalcarro)
            tipo = tipo + 1
            tipo2 = tipo2 + ((valor(total) * 4) + valor(totalcarro))

        ElseIf tipo(totalcarro) >= 24 And tipo(totalcarro) <= 36 Then
            total(totalcarro) = (valor(totalcarro) * 5) + valor(totalcarro)
            tipo3 = tipo3 + ((valor(totalcarro) * 5) + valor(totalcarro))

        ElseIf valor(totalcarro) > 8000 And valor(totalcarro) = 12 Then
            total(totalcarro) = (valor(totalcarro) * (4 - 3)) + valor(totalcarro) Or total(totalcarro) = (valor(totalcarro) * (5 - 3)) + valor(totalcarro) Or
            tipo3 = tipo3 + ((valor(totalcarro) * (4 - 3)) + valor(totalcarro) Or total(totalcarro) = (valor(totalcarro) * (5 - 3)) + valor(totalcarro) Or
                total(totalcarro))

            MsgBox("NO SE PUEDE ALQUILAR AUTOS")
        End If
    End Sub
    Sub salir()

        If MsgBox("¿Desea Salir?", vbQuestion + vbYesNo, "Mensaje Salida") = vbYes Then
            Form1.Close()
        Else
            limpiar_entradas()

        End If
    End Sub
    Sub limpiar_entradas()
        Form1.TextBox1.Clear()
        Form1.TextBox2.Clear()
        Form1.RadioButton1.Checked = False
        Form1.RadioButton2.Checked = False
        Form1.RadioButton3.Checked = False
        Form1.RadioButton4.Checked = False
        Form1.RadioButton5.Checked = False
        Form1.RadioButton6.Checked = False
        Form1.RadioButton7.Checked = False
        Form1.TextBox1.Focus()
    End Sub
    Sub limpiar_vectores()
        Dim J As Byte

        Form1.DataGridView1.Rows.Clear()

        fila = 0

        I = 0
    End Sub
    Sub Mostrar_vectores()
        Form1.DataGridView1.Rows.Clear()

        I = 0
        While (I <= 4)
            '
            If (tipo(7) <> Nothing) Then
                Form1.DataGridView1.Rows.Add(cliente, placa, tipo, cobrobase, kilometrajeinicial, kilometrajefinal, total)
            Else
                'Si están vacíos, se realiza una salida forzada del ciclo while con la siguiente instrucción
                Exit While
            End If
            'se incrementa el valor de la variable I para controlar las vueltas del ciclo
            I = I + 1
        End While
    End Sub


    Sub buscar_con_For()
        Dim g As Byte
        For g = 0 To 4

            If (tipo(7) <> Nothing) Then
                If (Val(tipo(7)) = Val(Form1.TextBox2.Text)) Then
                    MsgBox("registro encontrado")
                    Exit For
                End If
            Else
                Exit For
            End If
        Next g
        If (g = 5) Then
            MsgBox("carro no encontrado")
        Else
            MsgBox("Registro Encontrado exitosamente")

            Form1.TextBox1.Text = tipo(7)
            Form1.TextBox2.Text = tipo(7)



            If (tipo(7) = "tipo1") Then
                'SI el curso es Matemática almacenado en la columna 5, entonces se activa el radiobutton1
                'y así se hace con los demás cursos y radiobutton1
                Form1.RadioButton1.Checked = True
            ElseIf (tipo(7) = "tipo2") Then
                Form1.RadioButton2.Checked = True
            ElseIf (tipo(7) = "tipo3") Then
                Form1.RadioButton3.Checked = True
            ElseIf (tipo(7) = "tipo4") Then
                Form1.RadioButton4.Checked = True
            Else : Form1.RadioButton5.Checked = True
            End If


            Form1.DataGridView1.Rows.Clear()
            Form1.DataGridView1.Rows.Add(cliente, placa, tipo, cobrobase, kilometrajeinicial, kilometrajefinal, total)

            fila = I
        End If
    End Sub




End Module
