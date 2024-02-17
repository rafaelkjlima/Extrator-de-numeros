Function ExtNum (celula As String)
Dim caracteres As long
Dim result As String
Dim x  As long
Dim texto As String

caracteres = Len(celula)

for x = 1 to caracteres

    texto = mid(celula,x,1)

    if isnumeric(texto) then
        result = result & texto
    End if

next 

ExtNum=result
End Function