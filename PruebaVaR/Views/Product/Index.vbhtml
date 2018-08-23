
@Code
    Layout = Nothing
End Code

<!DOCTYPE html>

<html>
<head>
    <meta name="viewport" content="width=device-width" />
    <title>Index</title>
</head>
<body>



    @Using Html.BeginForm("Import", "Product", Nothing, FormMethod.Post, New With {.enctype = "multipart/form-data"})
        @Html.Raw(ViewBag.Error)

        @<div >
            <span>Excel File </span> <input type="file" name="excelfile" />
            <br />
            <input type="submit" value="Import" />
            End Using
        </div>
    End Using

    </body>
</html>
