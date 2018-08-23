
@Code
    Layout = Nothing
End Code

<!DOCTYPE html>

<html>
<head>
    <meta name="viewport" content="width=device-width" />
    <title>Success</title>
</head>
<body>
    <h3>Lista Productos</h3>

    <table cellpadding="2" cellspacing="2" border="1">
        <tr>
            <th>Area</th>
            <th>Elemento</th>
            <th>Cantidad</th>
            <th>Lago</th>
            <th>Ancho</th>
            <th>Alto</th>
            <th>Lados</th>
        </tr>
        @code
            For Each p In ViewBag.Lista
            @<tr>
                <td>@p.area</td>
                <td>@p.elemento</td>
                <td>@p.cantidad</td>
                <td>@p.largo</td>
                <td>@p.ancho</td>
                <td>@p.alto</td>
                <td>@p.lado</td>

            </tr>
            Next
        End Code

    </table>

</body>
</html>
