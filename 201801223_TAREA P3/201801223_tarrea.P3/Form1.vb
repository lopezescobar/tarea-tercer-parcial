Public Class Form1
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Mostrar_Matriz()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim existe As Boolean = True


        I = 0

        While (I <= 4) And (existe)

            If (tipo(7) <> Nothing) Then
                If (Val(tipo(7)) = Val(TextBox2.Text)) Then
                    existe = False
                Else
                    I = I + 1  '
                End If
            Else
                Exit While
            End If
        End While

        If (existe) Then
            MsgBox("Carnet no encontrado")
            TextBox9.Focus()
        Else
            MsgBox("Registro Encontrado exitosamente")

            TextBox1.Text = tipo(7)
            TextBox2.Text = tipo(7)



            If (tipo(7) = "tipo1") Then

                RadioButton1.Checked = True
            ElseIf (tipo(7) = "tipo2") Then
                RadioButton2.Checked = True
            ElseIf (tipo(7) = "tipo3") Then
                RadioButton3.Checked = True
            ElseIf (tipo(7) = "tipo4") Then
                RadioButton4.Checked = True
            Else : RadioButton4.Checked = True
            End If


            DataGridView1.Rows.Clear()
            DataGridView1.Rows.Add(cliente, placa, tipo, cobrobase, kilometrajeinicial, kilometrajefinal, total)

            fila = I
        End If
    End Sub


    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        tipo(fila = 0) = TextBox1.Text
        tipo(fila = 1) = TextBox2.Text


        If (RadioButton1.Checked) Then
            tipo(fila = 5) = RadioButton1.Text
        ElseIf (RadioButton2.Checked) Then
            tipo(fila = 5) = RadioButton2.Text
        ElseIf (RadioButton3.Checked) Then
            tipo(fila = 5) = RadioButton3.Text
        Else : tipo(fila = 5) = RadioButton4.Text
        End If

        'para calcular el promedio se utilizan las columnas de la matriz donde se encuentran guardadas las 
        'notas 1, 2 y 3, para ello hay que convertirlas con la función VAL ya que la matriz completa es de 
        'tipo string
        tipo(fila = 6) = Str(Round((Val(tipo(fila = 2)) + Val(tipo(fila = 3)) + Val(tipo(fila = 4))) / N, 2))
        'se indica con un mensaje que el registro fue actualizado correctamente
        MsgBox("Registro actualizado correctamente en la vect0r")
        fila = 5

        limpiar_entradas()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        I = 0

        While (I <= 6)
            'se indica siempre la misma fila, lo que cambia es el número de columna en dicha fila a eliminar
            Alumno(fila, I) = Nothing
            I = I + 1
        End While

        'Ahora se procede a hacer los corrimientos, esto se refiere a que los registros que se encuentran abajo del
        'registro eliminado se corren hacia arriba, para ello se usa un ciclo de la siguiente forma
        'nótese que el inicio del ciclo es en la fila del dato encontrado y termina en la última posición menos 1 de la matriz
        'se utiliza la misma variable FILA que es donde se encontró el registro buscado para control del ciclo while
        I = fila
        While (I <= 3)

            tipo(7) = tipo(I + 1 = 0)
            tipo(7) = tipo(I + 1 = 1)
            tipo(7) = tipo(I + 1 = 2)
            tipo(7) = tipo(I + 1 = 3)
            tipo(7) = tipo(I + 1 = 4)
            tipo(7) = tipo(I + 1 = 5)
            tipo(7) = tipo(I + 1 = 6)
            I = I + 1
        End While
        MsgBox("Registro Eliminado exitosamente")

        tipo(I = 0) = Nothing
        tipo(I = 1) = Nothing
        tipo(I = 2) = Nothing
        tipo(I = 3) = Nothing
        tipo(I = 4) = Nothing
        tipo(I = 5) = Nothing
        tipo(I = 6) = Nothing

        fila = I

        limpiar_entradas()
        DataGridView1.Rows.Clear()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        I = 0

        While (I <= 6)

            tipo(fila = I) = Nothing
            I = I + 1
        End While

        le
        I = fila
        While (I <= 3)

            tipo(I = 0) = tipo(I + 1 = 0)
            tipo(I = 1) = tipo(I + 1 = 1)
            tipo(I = 2) = tipo(I + 1 = 2)
            tipo(I = 3) = tipo(I + 1 = 3)
            tipo(I = 4) = tipo(I + 1 = 4)
            tipo(I = 5) = tipo(I + 1 = 5)
            tipo(I = 6) = tipo(I + 1 = 6)
            I = I + 1
        End While
        MsgBox("Registro Eliminado exitosamente")

        tipo(I = 0) = Nothing
        tipo(I = 1) = Nothing
        tipo(I = 2) = Nothing
        tipo(I = 3) = Nothing
        tipo(I = 4) = Nothing
        tipo(I = 5) = Nothing
        tipo(I = 6) = Nothing

        fila = I

        limpiar_entradas()
        DataGridView1.Rows.Clear()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        End
    End Sub
End Class
