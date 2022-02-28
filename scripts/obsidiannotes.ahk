#Include <Default_Settings>

Loop files, % pathDirectories(2) . "\*.md"
    Main(A_LoopFileFullPath)

return


Main(filePath) {
/*
    This functions inserts the data from the Obsidian notes to the Access database

    Parameters:
        - filePath (string): Path of the Obsidian directory.

    Returns:
        - None.
*/

    ; Make the ODBC connection to insert into the database
    dbConn := ADOConn(A_ScriptDir . "\..\data\external\Diet_be.accdb")

    cmd := ComObjCreate("ADODB.Command")
    cmd.ActiveConnection := dbConn

    ; Get data from file to make the SQL statement
    FileRead, textData, % filePath
    cmd.CommandText := GetQueryString(textData)

    ; Insert data into database
    Try,
        cmd.Execute
    Catch, er {
        MsgBox, 48, SQL insertiion was not possible., 3
        dbConn.Close()
        ExitApp
    }

    dbConn.Close()

    ; Delete the file to avoid repetition
    FileDelete, % filePath
}

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

    Try,
        adoConn.Open(connectionStr)
    Catch, e {
        MsgBox, 48, ODBC connection was not possible., 3
        ExitApp
    }

    return adoConn
}

GetQueryString(textData) {
/*
    Forms the query string for insertion of Access database.

    This functions gets the text and data from the Obsidian notes and transforms them to SQL.

    Parameters:
        - textData (string): Data from the text file.
    
    Return:
        - The string of SQL to make the insertion.
*/

    ; I wanto the extrac the text of the note and the creation date
    regexArr := ["O)___\n(.*)", "\d{4,4}-\d{2,2}-\d{2,2}\s\d{2,2}:\d{2,2}", "O)#\s([^`n]+)"]
    RegExMatch(textData, regexArr[1], journalText)
    RegExMatch(textData, regexArr[2], dateData)
    RegExMatch(textData, regexArr[3], titleText)

    return "INSERT INTO emotional_journal (paragraph, title, journal_created_at) VALUES ('" . journalText.Value(1) . "', '" . titleText.Value(1) . "', #" . dateData . "#);"

}