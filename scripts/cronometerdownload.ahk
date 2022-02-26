#Include <Default_Settings>
#Include <Chrome>

googlePath := A_ProgramFiles . "\Google\Chrome\Application\chrome.exe --remote-debugging-port=9222"

; Navigate to cronometer web page
Run, % googlePath
Sleep, 3000
Page := Chrome.GetPage()
Page.Call("Page.navigate", {"url": "https://cronometer.com/#settings"})
Page.WaitForLoad()

Sleep, 3000

; Click the adds GUI box in case it appears
Try,
    Page.Evaluate("document.getElementsByClassName('titlebar')[0].querySelector('img').click()")
Catch, err
    Sleep, 500

; Wait until the button of "Soporte" appears on the web page
CoordMode, Pixel, Screen
Loop, {
    PixelGetColor, green, 80, 988, RGB
    Sleep, 250

} Until (green == 0x27AE60)

; Click the buttons to download the csv file
Try, {
    Page.Evaluate("document.getElementsByClassName('tab-section')[4].querySelector('button').click()")
    Sleep, 500

    ; Set the value of the 3 last months
    Page.Evaluate("document.getElementsByClassName('gwt-ListBox')[39].value='Last 3 months'")

    ; Click to download data
    Page.Evaluate("document.getElementsByClassName('prettydialog')[0].querySelector('button').click()")
}

; In case the DOM is changed catch the error
Catch, e {
    MsgBox, 48, Warning, The querySelector was changed., 3
    WinClose Google Chrome
    ExitApp
}

; In the download folder wait until the csv is downloaded
fileCronometer := pathDirectories(1) . "\dailySummary.csv"

WaitForFileInDrive(fileCronometer)

; Get the last date to retrieve all rows from which are not in macronutrients table
dbPath := A_ScriptDir . "\..\data\external\Diet_be.accdb"
ExistingFile(dbPath)
queryStr := "SELECT Format(DateAdd(""d"", 1, MAX(Fecha)), ""yyyy-mm-dd"") AS last_date FROM macronutrients;"
lastDate := GetLastDateFromDB(dbPath, queryStr)

WinClose Google Chrome

; Create a ADODB connection and insert the row into the database.
InsertIntoAccDB(fileCronometer, lastDate, dbPath)

; Erase all csv files used
DeletingCSVFiles(fileCronometer)

return


ADOConn(dbPath) {
/*
Stablishes the ODBC connection with the database.

This functions creates and open a ODBC connection with the access.

Parameters:
    - dbPath (string) = Path directory to the database.

Returns:
    - ODBC connection open.
*/
    connectionStr := "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" . dbPath . ";Persist Security Info=False;"
    adoConn := ComObjCreate("ADODB.Connection")
    adoConn.Open(connectionStr)

    return adoConn
}

RecordSetOpen(dbConn, query) {
/*
Extract information of SQL recordset.

This functions gets the first value of the extracted value from SQL.

Parameters:
    - dbConn (ADODB connection) = ODBC connection opened object.
    - query (string) = SQL command to exexute.

Returns:
    - ODBC connection open
*/
    return dbConn.Execute(query)
}

WaitForFileInDrive(filePath) {
/*
Waits the given file to be in harddrive.

This functions waits until the given file is in the harddrive disk.

Parameters:
    - filePath (string) = Path directory where is the file.

Returns:
    - None.
*/
    Loop, {
        Sleep, 500
        fileAttribute := FileExist(filePath)
    } Until (fileAttribute != "")

        return
}

; ! Deprecated
CleanCSVForLoad(filePath) {
/*
Executes a shell Script

This functions executes a shell command and waits until it finishes.

Parameters:
    - filePath (string) = Path directory where is the Script.

Returns:
    - None.
*/
    ExistingFile(filePath)

    oWsh := ComObjCreate("WScript.Shell")
    oWsh.Run(filePath, 0, true)

    return
}

; Deprecated
GetHeadersCSV(filePath) {
/*
Get the headers file of a csv file.

This functions gets the headers values from a csv file.

Parameters:
    - filePath (string) = Path directory where is the file.

Returns:
    - A string with the headers.
*/

    FileReadLine, firstLine, % filePath, 1

    return firstLine
}

GetLastDateFromDB(dbPath, query) {
/*
Retreives the date to get iterate through csv file.

This functions retreived the date to start the iteration in the csv file.

Parameters:
    - dbPath (string) = Path directory to the database.
    - query (string) = SQL command to exexute.

Returns:
    - A date value.
*/
    ; Get ODBC connection
    dbConn := ADOConn(dbPath)
    ; Get the recordset query
    rs := recordSetOpen(dbConn, query)
    lastDate := rs.Fields.Item(0).Value
    rs.Close()
    dbConn.Close()

    return lastDate
}

InsertIntoAccDB(filePath, lastDate, dbPath) {
/*
Creates the cleaned csv file.

This functions forms the csv file with the correct headers and data filtered.

Parameters:
    - filePath (string) = The csv where is the raw data.
    - lastDate (date) = Date value.
    - dbPath (string) = Path directory of the database.

Returns:
    - None.
*/
    dbConn := ADOConn(dbPath)
    cmd := ComObjCreate("ADODB.Command")
    cmd.ActiveConnection := dbConn
    i := 0

    ; Loop thorough all rows in the file and just extract the rows that match
    Loop, read, % filePath
    {
        ; Extract all above the last date to the end of the file
        If (RegExMatch(A_LoopReadLine, lastDate) > 0) or (i > 0 and InStr(A_LoopReadLine, "true", false) > 0) {
            ; Parse the date to have the correct format for Access database
            line := ParseRowToSQL(A_LoopReadLine)
            queryInsert := "INSERT INTO macronutrients (" . HeadersFieldsToAccDB() . ") VALUES (" . line . ");"
            cmd.CommandText := queryInsert
            ; Execute the insertion
            Try,
                cmd.Execute
            Catch, e {
                MsgBox, 48, Warning, The ODBC connection failed for INSERT INTO statement., 3
                dbConn.Close()
                ExitApp
            }
            ++i
        }
    }

    dbConn.Close()

    return
}

DeletingCSVFiles(fileDirectory) {
/*
Deletes a file in the directory.

This functions erases a file file or many files in the directory.

Parameters:
    - fileDirectory (string) = Path directory to the file to erase.

Returns:
    - None.
*/
    FileDelete, % fileDirectory
    MsgBox, 64, Script finished, The data was loaded inside database., 3

    return
}

; Deprecated
DatabaseTransferData(dbPath, fileToImport) {
/*
Import the csv file to the database.

This functions imports the csv file cleaned data to database.

Parameters:
    - dbPath (string) = Path to the database.
    - fileToImport (string) = Path directory of the csv cleaned file.

Returns:
    - None.
*/
        ExistingFile(dbPath)

        oAcc := ComObjCreate("Access.Application")
        oAcc.OpenCurrentDataBase(dbPath)
        ;oAcc.Run("main")

        ; Variables para agregar un registro en la tabla con fecha del día anterior para el comando SQL
        currentDate := A_Now
        EnvAdd, currentDate, -1, Days
        FormatTime, formatedDate, % currentDate, yyyy/MM/dd
        insertQuery := "INSERT INTO downloads (start_date) VALUES (#" . formatedDate . "#);"

        oAcc.DoCmd.SetWarnings(false)
        ; Agrego la siguiente fecha de ejecución
        oAcc.DoCmd.RunSql(insertQuery)

        ;Transferencia de los datos del archivo CSV
        oAcc.DoCmd.TransferText(0,, "macronutrients", fileToImport, true)

        ; Reactivo el mensaje de advertencia
        oAcc.DoCmd.SetWarnings(true)

        ; Cierro la installed de Access
        oAcc.CloseCurrentDatabase
        oAcc.Quit

        return
}

HeadersFieldsToAccDB() {
/*
Headers of the fields of the database.

This functions returns the corresponding headers to match the fields in the database.

Parameters:
    - None.

Returns:
    - The headers file string.
*/
    return "Fecha,energy,alcohol,caffeine,water,vitamin_b1,vitamin_b2,vitamin_b3,vitamin_b5,vitamin_b6,vitamin_b12,folate,vitamin_a,vitamin_c,vitamin_d,vitamin_e,vitamin_k,calcium,copper,iron,magnesium,manganese,phosphorus,potassium,selenium,sodium,Zinc,carbohydrates,fiber,starch,simple_sugars,total_carbohydrates,fats,cholesterol,monounsaturated_fats,polyunsaturated_fats,saturated_fats,trans_fats,Omega_3,Omega_6,cysteine,histidine,isoleucine,leucine,lysine,methionine,phenylalanine,proteins,treonine,triptophane,tyrosine,valine,completed"
}

ParseRowToSQL(row) {
/*
Parses the row of csv file.

This function parsers the date of each row to the correct form for date field in Access.

Parameters:
    - row (string) = A row of the csv file.

Returns:
    - A string with row parsed.
*/
    RegExMatch(row, "\d{4,4}-\d{2,2}-\d{2,2}", dateString)

    return StrReplace(row, dateString, "#" . dateString . "#")
}