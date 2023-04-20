Imports System.Collections.Immutable
Imports System.Diagnostics.Eventing
Imports System.IO
Imports System.Text.RegularExpressions
Imports ABI
Imports System.Xml
Imports ICSharpCode.SharpZipLib.GZip
Imports ICSharpCode.SharpZipLib.Tar
Imports Microsoft.Data.Sqlite
Imports Microsoft.EntityFrameworkCore.Sqlite.Update.Internal
Imports Microsoft.EntityFrameworkCore.ValueGeneration.Internal

Public Class Form1
    ' Support functions for VKCL
    Private Const CheckerDB = "data Source=Checker.db3"   ' name of the checker file
    Private Const NameFile = "data Source=Name.db3"   ' name of the names file
    Dim updated As Integer       ' number of  records updated
    Dim added As Integer       ' number of new records added
    ReadOnly StopWatch As New Stopwatch  ' timing device
    ReadOnly count As Integer
    Private Sub ExtractNamesFromLogsToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ExtractNamesFromLogsToolStripMenuItem1.Click
        ' Extract names data for the VKCL names file from existing contest logs

        Dim records As Integer

        Dim sql As SqliteCommand, sqldr As SqliteDataReader
        Dim sqlNames As SqliteCommand, sqlNamesdr As SqliteDataReader

        Using names As New SqliteConnection(NameFile), connect As New SqliteConnection(CheckerDB)
            updated = 0
            added = 0
            ' Get the number of records in the name file
            names.Open()
            connect.Open()
            sql = connect.CreateCommand
            sqlNames = names.CreateCommand
            sqlNames.CommandText = "Select COUNT(*) as Count from nameTbl"
            sqlNamesdr = sqlNames.ExecuteReader()
            sqlNamesdr.Read()
            records = sqlNamesdr("Count")
            sqlNamesdr.Close()
            TextBox2.AppendText($"The names file contains {records} records{vbCrLf}")
            sql.CommandText = "SELECT * FROM Contests JOIN Stations AS S USING (contestID) WHERE S.station IS NOT NULL ORDER BY start"
            sqldr = sql.ExecuteReader()
            While sqldr.Read
                Dim location = If(IsDBNull(sqldr("location")), "", sqldr("location"))
                Dim gridsquare = If(IsDBNull(sqldr("gridsquare")), "", sqldr("gridsquare"))
                If sqldr("name") <> "" And sqldr("station") <> "" Then AddToNames(names, sqldr("name"), sqldr("station"), location, gridsquare)
            End While
            sqldr.Close()
            sqlNames.CommandText = "Select COUNT(*) as Count from nameTbl"
            sqlNamesdr = sqlNames.ExecuteReader()
            sqlNamesdr.Read()
            records = sqlNamesdr("Count")
            sqlNamesdr.Close()
            TextBox2.AppendText($"name file now has {records} records; {added} entries added, {updated} entries updated{vbCrLf}")
        End Using
    End Sub
    Sub AddToNames(ByRef connect As SqliteConnection, Name As String, Callsign As String, Location As String, GridLocator As String)
        ' Add the extracted data to the names database
        Dim sql As SqliteCommand, sqldr As SqliteDataReader
        Dim name_rowId As Integer, words() As String

        Trace.Assert(Name <> "", $"Name cannot be blank{Environment.StackTrace}{vbCrLf}")
        Trace.Assert(Callsign <> "", $"Callsign cannot be blank{Environment.StackTrace}{vbCrLf}")

        Callsign = UCase(Callsign)
        GridLocator = LCase(GridLocator)
        Dim regex As New Regex("^[a-r][a-r][0-9][0-9]([a-x][a-x]([0-9][0-9])?)?$")     ' 4 or 6 or 8 character locator
        Dim match As Match = regex.Match(GridLocator)
        If Not match.Success Then GridLocator = ""      ' remove invalid locator
        Name = Replace(Name, "'", "''")     ' escape single quotes
        If IsDBNull(Location) Then Location = ""
        Location = Replace(Location, "'", "''")     ' escape single quotes
        ' if name contains 2 words, assume it is a normal personal name (e.g. Marc Hillman) and record only the first.
        ' If not, record full name
        words = Split(Name, " ")
        If words.Length = 2 Then Name = words(0)     ' Just the first word

        ' name_rowId is the primary key, not callsign, so we have to do this the hard way
        ' Search for callsign record. If found update, else insert
        sql = connect.CreateCommand
        sql.CommandText = $"SELECT * FROM nameTbl WHERE name_Callsign='{Callsign}'"
        Try
            sqldr = sql.ExecuteReader
            If sqldr.HasRows Then
                sqldr.Read()
                name_rowId = sqldr("name_rowId")    ' use primary key for speed
            Else
                name_rowId = 0      ' no record exists
            End If
            sqldr.Close()
            If name_rowId <= 0 Then
                ' New - do insert
                sql.CommandText = $"INSERT INTO nameTbl (name_Callsign,name_OpName,name_QTH,name_MdnLctr) VALUES ('{Callsign}','{Name}','{Location}','{GridLocator}')"
                sql.ExecuteNonQuery()
                added += 1
            Else
                ' Exists - do update
                sql.CommandText = $"UPDATE nameTbl SET name_OpName='{Name}',name_QTH='{Location}',name_MdnLctr='{GridLocator}' WHERE name_rowId={name_rowId}"
                sql.ExecuteNonQuery()
                updated += 1
            End If
        Catch ex As SqliteException
            MsgBox($"{ex.Message}{vbCrLf}SQL={sql.CommandText}", vbCritical + vbOK, "SQLite error")
        End Try
    End Sub

    Private Sub RemoveDuplicateLogsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RemoveDuplicateLogsToolStripMenuItem.Click
        ' It is common that submitted log folders contain multple versions of the same file. Versions are indicated by (n) in the filename.
        ' This can happen because people forgot they had already submitted a log, or they made an adjustment to an existing log.
        ' In either case, the file with the highest version is used, and all others are moved to a \duplicates folder.
        Dim contestID As Integer, fileList As New List(Of FileDetails)
        Dim sql As SqliteCommand, sqldr As SqliteDataReader, count As Integer = 0
        If dlgContest.ShowDialog = DialogResult.OK Then
            contestID = dlgContest.DataGridView1.SelectedRows(0).Cells("id").Value
            Using connect As New SqliteConnection(CheckerDB)
                Dim SolutionDirectory = IO.Directory.GetParent(IO.Directory.GetParent(IO.Directory.GetParent(My.Application.Info.DirectoryPath).ToString).ToString)
                connect.Open()
                ' get the contest details
                sql = connect.CreateCommand
                sql.CommandText = $"SELECT * FROM `Contests` WHERE `ContestID`={contestID}"
                sqldr = sql.ExecuteReader()
                sqldr.Read()
                Dim di As New DirectoryInfo($"{SolutionDirectory}{sqldr("path")}")     ' path to logs
                Dim fiArr As FileInfo() = di.GetFiles       ' get list of all files in contest folder
                For Each fi In fiArr
                    fileList.Add(New FileDetails(fi.FullName))  ' capture file details, and parse to filename & version
                Next
                ' Find duplicate files, i.e. filenames where the filenames match, but there exists a lower version
                Dim duplicates = fileList.Where(Function(a) fileList.Exists(Function(b) a.filename = b.filename And a.version < b.version))
                ' move all duplicate files to a \duplicates folder
                For Each duplicate In duplicates
                    Dim fi = My.Computer.FileSystem.GetFileInfo(duplicate.path)
                    Dim MovedFile = $"{fi.Directory}\duplicates\{fi.Name}"     ' add \duplicate to path
                    My.Computer.FileSystem.MoveFile(duplicate.path, MovedFile)  ' move duplicate file to \duplicates folder
                Next
                TextBox2.Text = $"{fileList.Count} files in {sqldr("path")}. {duplicates.Count} removed.{vbCrLf}"
            End Using
        End If
    End Sub
    Class FileDetails
        Property path As String     ' fully qualified name of file
        Property filename As String ' filename with version removed.
        Property version As Integer     ' version. 0 if no version
        Sub New(f As String)
            ' split f into name and optional version
            Dim fi As FileInfo
            _path = f
            fi = FileIO.FileSystem.GetFileInfo(f)
            Dim name As String = fi.Name      ' name, with version number
            Dim matches As MatchCollection = Regex.Matches(name, "^(.*)(\((\d)\))(.*)$")
            If matches.Count = 0 Then
                _filename = name
                _version = 0
            Else
                With matches(0)
                    _filename = $"{ .Groups(1).Value}{ .Groups(4).Value}"
                    _version = CInt(.Groups(3).Value)
                End With
            End If
        End Sub
    End Class
    ' =====================================================
    ' Lots of lists for lookup and validation
    ' =====================================================
    ' table of commonly used band labels, and actual frequency ranges.
    ' Sometimes the Cabrillo file contains the frequency as an integer
    Const MHz = 10 ^ 6, GHz = 10 ^ 9
    ReadOnly bandRange As New Dictionary(Of String, (low As Long, high As Long)) From {
        {"50", (50 * MHz, 54 * MHz)},
        {"144", (144 * MHz, 148 * MHz)},
        {"432", (420 * MHz, 450 * MHz)},
        {"1.2G", (1.24 * GHz, 1.3 * GHz)},
        {"2.4G", (2.4 * GHz, 2.45 * GHz)},
        {"3.3G", (3.3 * GHz, 3.6 * GHz)},
        {"5.6G", (5.65 * GHz, 5.85 * GHz)},
        {"10G", (10 * GHz, 10.5 * GHz)},
        {"24G", (24 * GHz, 24.25 * GHz)},
        {"47G", (47 * GHz, 47.2 * GHz)},
        {"78G", (76 * GHz, 81 * GHz)},
        {"122G", (122.25 * GHz, 123 * GHz)},
        {"134G", (134 * GHz, 141 * GHz)},
        {"241G", (241 * GHz, 250 * GHz)}
        }
    ' Score calculation table
    ReadOnly ScoreTable As New Dictionary(Of String, (multiplier As Single, threshold1 As Integer, mult1 As Integer, threshold2 As Integer, mult2 As Integer)) From {
            {"50", (1.7, 700, 1, 100, 1)},
            {"144", (1.0, 700, 1, 100, 1)},
            {"432", (2.7, 700, 1, 100, 1)},
            {"1.2G", (3.7, 0, 0, 1, 1)},
            {"2.4G", (4.4, 0, 0, 1, 1)},
            {"3.3G", (5.4, 0, 0, 1, 1)},
            {"5.6G", (6.4, 0, 0, 1, 1)},
            {"10G", (7.4, 0, 0, 1, 1)},
            {"24G", (10, 0, 0, 1, 1)},
            {"47G", (10, 0, 0, 1, 1)},
            {"78G", (10, 0, 0, 1, 1)},
            {"120G", (10, 0, 0, 1, 1)},
            {"122G", (10, 0, 0, 1, 1)},
            {"134G", (10, 0, 0, 1, 1)},
            {"241G", (10, 0, 0, 1, 1)}
            }
    ReadOnly Sections As New Dictionary(Of String, String) From {
                    {"A1", "Portable, Single Op 24 Hours"},
                    {"A2", "Portable, Single Op 8 Hours"},
                    {"B1", "Portable, Multi Op 24 Hours"},
                    {"B2", "Portable, Multi Op 8 Hours"},
                    {"C1", "Home Station, 24 Hours"},
                    {"C2", "Home Station, 8 Hours"},
                    {"D1", "Rover, 24 Hours"},
                    {"D2", "Rover, 8 Hours"}
                }
    ReadOnly SubSections As New Dictionary(Of String, String) From {
                    {"a", "Single band"},
                    {"b", "Four bands"},
                    {"c", "All bands"}
                }
    ' arrays of allowable values for validation
    ReadOnly CategoryStationValidation() As String = {"HOME", "DISTRIBUTED", "FIXED", "MOBILE", "PORTABLE", "ROVER", "ROVER-LIMITED", "ROVER-UNLIMITED", "EXPEDITION", "HQ", "SCHOOL", "EXPLORER"}
    ReadOnly CategoryOperatorValidation() As String = {"SINGLE-OP", "MULTI-OP", "CHECKLOG"}
    ReadOnly CategoryBandValidation() As String = {"ALL", "FOUR", "SINGLE", "6M", "2M", "432", "1.2G", "2.3G", "3.4G", "5.7G", "10G", "24G", "47G", "78G", "122G", "134G", "241G", "Light"}
    ReadOnly CategoryTimeValidation() As String = {"6-HOURS", "8-HOURS", "12-HOURS", "24-HOURS"}
    ReadOnly modeValidation() As String = {"PH", "CW", "FM", "DG"}
    Private Sub IngestLogsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IngestLogsToolStripMenuItem.Click
        ' load logs into database and perform checks
        Dim contestID As Integer, fileCount As Integer = 0, errors As Integer = 0

        Dim sql As SqliteCommand, sql1 As SqliteCommand, sqldr As SqliteDataReader, count As Integer = 0
        Dim QSOsql As SqliteCommand, Stationssql As SqliteCommand
        ' list of commonly used mis-spellings of cabrillo template names
        Dim templateMapping As New Dictionary(Of String, String) From {
            {"WIA-FIELDDAY", "WIA-VHF/UHF-FD"},
            {"VHFUHFFD", "WIA-VHF/UHF-FD"},
            {"WIA-SPRING-FIELDDAY-2022", "WIA-VHF/UHF-FD"},
            {"WIA Spring -VHF/UHF-FD 2022", "WIA-VHF/UHF-FD"},
            {"VHF-UHF-FIELD-DAY", "WIA-VHF/UHF-FD"}
            }

        If dlgContest.ShowDialog = DialogResult.OK Then
            contestID = dlgContest.DataGridView1.SelectedRows(0).Cells("id").Value
            Using connect As New SqliteConnection(CheckerDB)
                connect.Open()
                Dim tr As SqliteTransaction = connect.BeginTransaction
                Using tr
                    ' get the contest details
                    sql = connect.CreateCommand
                    sql1 = connect.CreateCommand
                    sql1.Transaction = tr
                    QSOsql = connect.CreateCommand
                    QSOsql.Transaction = tr
                    Stationssql = connect.CreateCommand
                    Stationssql.Transaction = tr
                    sql.CommandText = $"SELECT * FROM `Contests` WHERE `ContestID`={contestID}"
                    sqldr = sql.ExecuteReader()
                    sqldr.Read()

                    ' process the uploads folder
                    Dim di As New DirectoryInfo($"F:\Users\Marc\Documents\Visual Studio 2022\Projects\VKCL support\VKCL support{sqldr("path")}")     ' path to logs
                    sqldr.Close()
                    Dim fiArr As FileInfo() = di.GetFiles
                    'Array.Sort(fiArr, Function(fi1, fi2) String.Compare(fi1.Name, fi2.Name))    ' sort array by filename
                    ' Delete any existing data
                    Stationssql.CommandText = $"DELETE FROM `Stations` WHERE contestID={contestID}"
                    Stationssql.ExecuteNonQuery()
                    QSOsql.CommandText = $"DELETE FROM `QSO` WHERE contestID={contestID}"
                    QSOsql.ExecuteNonQuery()
                    ' Create a prepare statements
                    QSOsql.CommandText = $"INSERT INTO `QSO`
                        (contestID, section, date, band, mode, sent_call, sent_rst, sent_exch, sent_grid, rcvd_call, rcvd_rst, rcvd_exch, rcvd_grid, flags) VALUES
                        (@contestID,@section, @date,@band,@mode,@sent_call,@sent_rst,@sent_exch,@sent_grid,@rcvd_call,@rcvd_rst,@rcvd_exch,@rcvd_grid,@flags)"
                    QSOsql.Transaction = tr
                    QSOsql.Prepare()
                    Stationssql.CommandText = $"REPLACE INTO `Stations` 
                        (contestID, filename, station, CategoryStation, CategoryOperator, CategoryBand, CategoryTime, operators, section, subsection, gridsquare, name, location, email, soapbox, ClaimedQSO, ActualQSO, ClaimedScore, CreatedBy, result) VALUES 
                        (@contestID, @filename, @station, @CategoryStation, @CategoryOperator, @CategoryBand, @CategoryTime, @operators, @section, @subsection, @gridsquare, @name, @location, @email, @soapbox, @ClaimedQSO, @ActualQSO, @ClaimedScore, @CreatedBy,@result)"
                    Stationssql.Transaction = tr
                    Stationssql.Prepare()

                    ' process all files
                    TextBox2.Text = $"Processing folder {di.FullName}{vbCrLf}"
                    For Each fri In fiArr
                        fileCount += 1
                        TextBox2.AppendText($"Processing {fri.Name} - ")
                        Application.DoEvents()
                        ' initialise the values we are looking for
                        Dim Contest As String = ""
                        Dim ContestDate As String = ""
                        Dim Callsign As String = ""
                        Dim Name As String = ""
                        Dim Email As String = ""
                        Dim Location As String = ""
                        Dim GridLocator As String = ""
                        Dim ValidFile As Boolean = False
                        Dim result As String = ""
                        Dim soapbox As String = ""
                        Dim lastTime As String = ""
                        Dim time As String
                        Dim QSOcount As Integer = 0
                        Dim QSO As Dictionary(Of String, String)
                        Dim dt As DateTime
                        Dim ClaimedQSO As Integer = 0
                        Dim ClaimedScore As Integer = 0
                        Dim CreatedBy As String = ""
                        Dim CategoryStation As String = ""
                        Dim CategoryOperator As String = ""
                        Dim CategoryBand As String = ""
                        Dim CategoryTime As String = ""
                        Dim CategoryPower As String = ""
                        Dim Operators As String = ""
                        Dim section = ""
                        Dim subsection = ""

                        Select Case LCase(fri.Extension)
                            Case ".log", ".txt"
                                ' Cabrillo file
                                Dim fileReader As System.IO.StreamReader = My.Computer.FileSystem.OpenTextFileReader(fri.FullName)
                                While Not fileReader.EndOfStream
                                    Dim raw As String = fileReader.ReadLine
                                    raw = LTrim(raw)      ' remove leading spaces
                                    Dim line As String = Trim(Regex.Replace(raw, "\s+", " "))      ' remove multiple spaces
                                    Dim words = Split(line, " ")     ' split line into words
                                    Select Case words(0).ToUpper
                                        Case "START-OF-LOG:"
                                            ValidFile = True
                                        Case "CONTEST:"
                                            Contest = Join(words.Skip(1).ToArray, " ").ToUpper
                                            If templateMapping.ContainsKey(Contest) Then
                                                result &= $"Contest template {Contest} remapped to {templateMapping(Contest)}{vbCrLf}"
                                                Contest = templateMapping(Contest)
                                            End If
                                        Case "DATE-OF-CONTEST:"
                                            If ValidFile And words.Length >= 2 Then ContestDate = words(1)
                                        Case "CATEGORY:"
                                            ' Cabrillo V2
                                            If ValidFile And words.Length >= 4 Then
                                                CategoryOperator = words(1).ToUpper
                                                CategoryBand = words(2).ToUpper
                                                CategoryPower = words(3).ToUpper
                                            End If
                                        Case "CATEGORY-STATION:"
                                            If ValidFile And words.Length >= 2 Then CategoryStation = words(1).ToUpper
                                            CategoryStation = ConvertCabrillo(CategoryStation)
                                            If Not CategoryStationValidation.Contains(CategoryStation) Then result &= $"Invalid CATEGORY-STATION value of {CategoryStation}{vbCrLf}"
                                        Case "SECTION:"
                                            If ValidFile And words.Length >= 2 Then
                                                ' Non standard Cabrillo. Convert to standard
                                                Dim TheSection = Join(words.Skip(1).ToArray, " ").ToUpper.Replace(" ", "-")
                                                TheSection = ConvertCabrillo(TheSection)
                                                Select Case TheSection
                                                    Case "HOME"
                                                        CategoryStation = "HOME"
                                                        CategoryOperator = "SINGLE-OP"
                                                    Case "PORTABLE-SINGLE", "PORTABLE"
                                                        CategoryStation = "PORTABLE"
                                                        CategoryOperator = "SINGLE-OP"
                                                    Case "PORTABLE-MULTI"
                                                        CategoryStation = "PORTABLE"
                                                        CategoryOperator = "MULTI-OP"
                                                    Case Else
                                                        MsgBox($"Unsupported SECTION {TheSection} in {fri.Name}", vbCritical + vbOKOnly, "Unsupported SECTION")
                                                End Select
                                            End If
                                        Case "CATEGORY-OPERATOR:"
                                            If ValidFile And words.Length >= 2 Then CategoryOperator = words(1).ToUpper
                                            If Not CategoryOperatorValidation.Contains(CategoryOperator) Then result &= $"Invalid CATEGORY-OPERATOR value of {CategoryOperator}{vbCrLf}"
                                        Case "CATEGORY-BAND:", "SUB-SECTION:"
                                            If ValidFile And words.Length >= 2 Then CategoryBand = Join(words.Skip(1).ToArray, " ").ToUpper.Replace(" ", "-")
                                            ' Non standard Cabrillo - convert to standard
                                            CategoryBand = ConvertCabrillo(CategoryBand)
                                            If Not CategoryBandValidation.Contains(CategoryBand) Then result &= $"Invalid CATEGORY-BAND value of {CategoryBand}{vbCrLf}"
                                        Case "CATEGORY-TIME:", "DURATION:"
                                            If ValidFile And words.Length >= 2 Then CategoryTime = Join(words.Skip(1).ToArray, " ").ToUpper.Replace(" ", "-")
                                            If Not CategoryTimeValidation.Contains(CategoryTime) Then result &= $"Invalid CATEGORY-TIME value of {CategoryTime}{vbCrLf}"
                                        Case "DIVISION:"
                                            ' Non standard Cabrillo. Convert to standard
                                        Case "OPERATORS:"
                                            If ValidFile And words.Length >= 2 Then
                                                If Operators <> "" Then Operators &= ", "
                                                Operators &= Join(words.Skip(1).ToArray, " ")
                                            End If
                                        Case "CALLSIGN:"
                                            If ValidFile And words.Length >= 2 Then Callsign = basecall(words(1))   ' Callsign may be compound. Split out the base callsign
                                        Case "CLAIMED-CONTACTS:"
                                            If ValidFile And words.Length >= 2 Then ClaimedQSO = CInt(words(1))
                                        Case "CLAIMED-SCORE:"
                                            If ValidFile And words.Length >= 2 Then ClaimedScore = CInt(words(1))
                                        Case "CREATED-BY:"
                                            If ValidFile And words.Length >= 2 Then CreatedBy = Join(words.Skip(1).ToArray, " ")
                                        Case "NAME:"
                                            If ValidFile And words.Length >= 2 Then Name = Join(words.Skip(1).ToArray, " ")      ' join all words
                                        Case "EMAIL:"
                                            If ValidFile And words.Length >= 2 Then Email = Join(words.Skip(1).ToArray, " ")      ' join all words
                                        Case "LOCATION:"
                                            If ValidFile And words.Length >= 2 Then Location = Join(words.Skip(1).ToArray, " ")      ' join all words
                                        Case "GRID-LOCATOR:"
                                            If ValidFile And words.Length >= 2 Then GridLocator = words(1)
                                        Case "SOAPBOX:", "COMMENTS:"
                                            If ValidFile And words.Length >= 2 Then
                                                If soapbox <> "" Then soapbox &= $"{vbCrLf}"
                                                soapbox &= Join(words.Skip(1).ToArray, " ")      ' join all words
                                            End If
                                        Case "QSO:"
                                            If section = "" Then
                                                ' Work out the Section
                                                Select Case True
                                                    Case CategoryStation = "PORTABLE" And CategoryOperator = "SINGLE-OP" And CategoryTime = "24-HOURS"
                                                        section = "A1"
                                                    Case CategoryStation = "PORTABLE" And CategoryOperator = "SINGLE-OP" And CategoryTime = "8-HOURS"
                                                        section = "A2"
                                                    Case CategoryStation = "PORTABLE" And CategoryOperator = "MULTI-OP" And CategoryTime = "24-HOURS"
                                                        section = "B1"
                                                    Case CategoryStation = "PORTABLE" And CategoryOperator = "MULTI-OP" And CategoryTime = "8-HOURS"
                                                        section = "B2"
                                                    Case (CategoryStation = "HOME") And CategoryTime = "24-HOURS"
                                                        section = "C1"
                                                    Case (CategoryStation = "HOME") And CategoryTime = "8-HOURS"
                                                        section = "C2"
                                                    Case CategoryStation = "ROVER" And CategoryTime = "24-HOURS"
                                                        section = "D1"
                                                    Case CategoryStation = "ROVER" And CategoryTime = "8-HOURS"
                                                        section = "D2"
                                                End Select
                                                Select Case CategoryBand
                                                    Case "SINGLE"
                                                        subsection = "a"
                                                    Case "FOUR"
                                                        subsection = "b"
                                                    Case "ALL"
                                                        subsection = "c"
                                                End Select
                                                If section = "" Then result &= $"Unable to determine Section{vbCrLf}"
                                                If subsection = "" Then result &= $"Unable to determine SubSection{vbCrLf}"
                                            End If
                                            QSO = GetQSO(Contest, raw)      ' decode QSO
                                            If QSO.Count = 0 Then
                                                result &= $"QSO bad format '{raw}'{vbCrLf}"
                                            Else
                                                time = $"{QSO("time").Substring(0, 2)}:{QSO("time").Substring(2, 2)}"
                                                If Not QSO.ContainsKey("date") Then
                                                    ' Prior to VKCL V4.8, there was no date field. Need to reconstruct this.
                                                    ' Using DATE-OF-CONTEST, and looking for time wrap-around, determine the date
                                                    If time < lastTime Then
                                                        ' wrap-around - increment contest date
                                                        Dim d As Date = Date.Parse(ContestDate)
                                                        d = d.AddDays(1)
                                                        ContestDate = d.ToString("yyyy-MM-dd")
                                                    End If
                                                    QSO.Add("date", ContestDate)
                                                End If
                                                ' try to extract datetime
                                                Try
                                                    dt = DateTime.Parse($"{QSO("date")} {time}")
                                                Catch ex As Exception
                                                    TextBox2.AppendText($"Invalid date {QSO("date")} {time}")
                                                    Exit While
                                                End Try
                                                QSO.Add("datetime", dt.ToString("yyyy-MM-dd HH:mm"))
                                                If Not QSO.ContainsKey("sent_call") Then QSO.Add("sent_call", Callsign)
                                                QSOFixup(QSO)
                                                QSOcount += 1
                                                With QSOsql.Parameters
                                                    .Clear()
                                                    .AddWithValue("contestID", contestID)
                                                    .AddWithValue("section", section)
                                                    .AddWithValue("date", QSO("datetime"))
                                                    .AddWithValue("band", QSO("freq"))
                                                    .AddWithValue("mode", QSO("mode").ToUpper)
                                                    .AddWithValue("sent_call", QSO("sent_call").ToUpper)
                                                    .AddWithValue("sent_rst", QSO("sent_rst"))
                                                    .AddWithValue("sent_exch", QSO("sent_exch"))
                                                    .AddWithValue("sent_grid", QSO("sent_grid").ToUpper)
                                                    .AddWithValue("rcvd_call", QSO("rcvd_call").ToUpper)
                                                    .AddWithValue("rcvd_rst", QSO("rcvd_rst"))
                                                    .AddWithValue("rcvd_exch", QSO("rcvd_exch"))
                                                    .AddWithValue("rcvd_grid", QSO("rcvd_grid").ToUpper)
                                                    .AddWithValue("flags", 0)
                                                End With
                                                Try
                                                    QSOsql.ExecuteNonQuery()
                                                Catch ex As SqliteException
                                                    If ex.SqliteErrorCode <> 19 Then
                                                        ' Error 19 occurs when attempting to insert a duplicate entry. This occurs when there are duplicate logs.
                                                        ' In this case we just ignore the duplicate
                                                        MsgBox($"SQLite error - {ex.Message}", vbCritical + vbOKOnly, "SQLite error")
                                                    End If
                                                End Try
                                            End If
                                    End Select
                                End While
                                If Not ValidFile Then
                                    result &= $"Not a valid Cabrillo file"
                                Else
                                    result &= $"{QSOcount} QSO extracted"
                                End If
                                fileReader.Close()
                                lastTime = time
                                ' update Stations table
                                ' add data to database
                                With Stationssql.Parameters
                                    .Clear()
                                    .AddWithValue("contestID", contestID)
                                    .AddWithValue("filename", fri.FullName)
                                    .AddWithValue("station", Callsign)
                                    .AddWithValue("operators", Operators)
                                    .AddWithValue("CategoryStation", CategoryStation)
                                    .AddWithValue("CategoryOperator", CategoryOperator)
                                    .AddWithValue("CategoryBand", CategoryBand)
                                    .AddWithValue("CategoryTime", CategoryTime)
                                    .AddWithValue("section", section)
                                    .AddWithValue("subsection", subsection)
                                    .AddWithValue("gridsquare", GridLocator)
                                    .AddWithValue("name", Name)
                                    .AddWithValue("location", Location)
                                    .AddWithValue("email", Email)
                                    .AddWithValue("soapbox", soapbox)
                                    .AddWithValue("ClaimedQSO", ClaimedQSO)
                                    .AddWithValue("ActualQSO", QSOcount)
                                    .AddWithValue("ClaimedScore", ClaimedScore)
                                    .AddWithValue("CreatedBy", CreatedBy)
                                    .AddWithValue("result", result)
                                End With
                                Stationssql.ExecuteNonQuery()
                            Case ".db3"
                                ' An Sqlite database
                                Dim sqldb3 As SqliteCommand, sqldrdb3 As SqliteDataReader
                                Try
                                    Using connectdb3 As New SqliteConnection($"data Source={fri.FullName}")
                                        connectdb3.Open()
                                        sqldb3 = connectdb3.CreateCommand
                                        ' Have to get the callsign to fill in QSO field
                                        sqldb3.CommandText = "SELECT * FROM `Contest`"
                                        sqldrdb3 = sqldb3.ExecuteReader
                                        If sqldrdb3.Read() Then
                                            Callsign = sqldrdb3("cont_CallSign")
                                        End If
                                        sqldrdb3.Close()

                                        ' Now copy QSO
                                        sqldb3.CommandText = "SELECT * FROM `contLog`"
                                        sqldrdb3 = sqldb3.ExecuteReader
                                        While sqldrdb3.Read
                                            With QSO
                                                .Clear()
                                                .Add("datetime", sqldrdb3("sql_zDTime"))
                                                .Add("freq", sqldrdb3("sql_iBand"))
                                                .Add("mode", Trim(sqldrdb3("sql_Mode").toupper))
                                                .Add("sent_call", Callsign)
                                                .Add("sent_rst", sqldrdb3("sql_rptSent"))
                                                .Add("sent_exch", sqldrdb3("sql_nmbSent"))
                                                .Add("sent_grid", IIf(IsDBNull(sqldrdb3("sql_ALctr")), "", sqldrdb3("sql_ALctr")).toupper)
                                                .Add("rcvd_call", sqldrdb3("sql_callSign").toupper)
                                                .Add("rcvd_rst", sqldrdb3("sql_rptRcvd"))
                                                .Add("rcvd_exch", sqldrdb3("sql_nmbRcvd"))
                                                .Add("rcvd_grid", IIf(IsDBNull(sqldrdb3("sql_Name")), "", sqldrdb3("sql_Name")).toupper)
                                                QSOFixup(QSO)
                                            End With

                                            With QSOsql.Parameters
                                                .Clear()
                                                .AddWithValue("contestID", contestID)
                                                .AddWithValue("section", section)
                                                .AddWithValue("date", QSO("datetime"))
                                                .AddWithValue("band", QSO("freq"))
                                                .AddWithValue("mode", QSO("mode"))
                                                .AddWithValue("sent_call", Callsign)
                                                .AddWithValue("sent_rst", QSO("sent_rst"))
                                                .AddWithValue("sent_exch", QSO("sent_exch"))
                                                .AddWithValue("sent_grid", QSO("sent_grid"))
                                                .AddWithValue("rcvd_call", QSO("rcvd_call"))
                                                .AddWithValue("rcvd_rst", QSO("rcvd_rst"))
                                                .AddWithValue("rcvd_exch", QSO("rcvd_exch"))
                                                .AddWithValue("rcvd_grid", QSO("rcvd_grid"))
                                                .AddWithValue("flags", 0)
                                            End With

                                            QSOcount += 1
                                            Try
                                                QSOsql.ExecuteNonQuery()
                                            Catch ex As SqliteException
                                                If ex.SqliteErrorCode <> 19 Then
                                                    ' Error 19 occurs when attempting to insert a duplicate entry. This occurs when there are duplicate logs.
                                                    ' In this case we just ignore the duplicate
                                                    MsgBox($"SQLite error - {ex.Message}", vbCritical + vbOKOnly, "SQLite error")
                                                End If
                                            End Try
                                        End While
                                        sqldrdb3.Close()

                                        ' update Stations table
                                        sqldb3.CommandText = "SELECT * FROM `Contest`"
                                        sqldrdb3 = sqldb3.ExecuteReader
                                        If sqldrdb3.Read() Then
                                            'Location = sqldrdb3("cont_Location")

                                            Callsign = IIf(IsDBNull(sqldrdb3("cont_Callsign")), "", sqldrdb3("cont_Callsign").toupper)
                                            result &= $"{QSOcount} QSO extracted"
                                            With Stationssql.Parameters
                                                .Clear()
                                                .AddWithValue("contestID", contestID)
                                                .AddWithValue("filename", fri.FullName)
                                                .AddWithValue("station", Callsign)
                                                .AddWithValue("operators", "")   ' TODO
                                                .AddWithValue("gridsquare", IIf(IsDBNull(sqldrdb3("cont_ActvLctr")), "", sqldrdb3("cont_ActvLctr")))
                                                .AddWithValue("name", IIf(IsDBNull(sqldrdb3("cont_OpName")), "", sqldrdb3("cont_OpName")))
                                                .AddWithValue("location", IIf(IsDBNull(sqldrdb3("cont_Location")), "", sqldrdb3("cont_Location")))
                                                .AddWithValue("email", IIf(IsDBNull(sqldrdb3("cont_eMail")), "", sqldrdb3("cont_eMail")))
                                                .AddWithValue("soapbox", IIf(IsDBNull(sqldrdb3("cont_SoapBox")), "", sqldrdb3("cont_SoapBox")))
                                                .AddWithValue("CategoryStation", "")    ' TODO
                                                .AddWithValue("CategoryOperator", "")   ' TODO
                                                .AddWithValue("CategoryBand", "")   ' TODO
                                                .AddWithValue("CategoryTime", "")   ' TODO
                                                .AddWithValue("section", "")   ' TODO
                                                .AddWithValue("subsection", "")   ' TODO
                                                .AddWithValue("ClaimedQSO", IIf(IsDBNull(sqldrdb3("cont_NumCont")), "", sqldrdb3("cont_NumCont")))
                                                .AddWithValue("ActualQSO", QSOcount)
                                                .AddWithValue("ClaimedScore", IIf(IsDBNull(sqldrdb3("cont_ClmdScore")), "", sqldrdb3("cont_ClmdScore")))
                                                .AddWithValue("CreatedBy", IIf(IsDBNull(sqldrdb3("cont_VKCLver")), "", sqldrdb3("cont_VKCLver")))
                                                .AddWithValue("result", result)
                                            End With
                                            Stationssql.ExecuteNonQuery()
                                        Else
                                            Throw New System.Exception("Failed to find any record in the Contest table")
                                        End If
                                        sqldrdb3.Close()
                                    End Using
                                Catch ex As SqliteException
                                    MsgBox($"{ex.Message}{vbCrLf}SQL={sql.CommandText}", vbCritical + vbOK, "SQLite error")
                                End Try

                            Case Else
                                result = $"Unsupported file type{vbCrLf}"
                                sql1.CommandText = $"REPLACE INTO `Stations` (contestID,filename,result) VALUES ({contestID},'{fri.FullName}','{result}')"
                                sql1.ExecuteNonQuery()
                        End Select
                        count += 1
                        TextBox2.AppendText($"QSO count {QSOcount}{vbCrLf}")
                    Next
                    ' do some post ingest checks
                    TextBox2.AppendText($"POST INGEST CHECKS{vbCrLf}")
                    ' You cannot have an entry in a 24-HOURS category and an 8-HOURS category
                    sql.CommandText = $"SELECT `station`
                                        FROM   `Stations` AS A
                                        WHERE  `contestID`={contestID}
                                        AND    `CategoryTime`='8-HOURS'
                                        AND    EXISTS
                                               (
                                                      SELECT *
                                                      FROM   `Stations`
                                                      WHERE  `contestID`={contestID}
                                                      AND    `station`=A.station
                                                      AND    `CategoryTime`='24-HOURS')"
                    sqldr = sql.ExecuteReader
                    If sqldr.HasRows Then
                        While sqldr.Read
                            TextBox2.AppendText($"{sqldr("station")} has an entry in both 8 and 24 hour categories. The 8 hour entry should be disqualified.{vbCrLf}")
                        End While
                    End If
                    tr.Commit()
                    connect.Close()
                End Using
            End Using
            TextBox2.AppendText($"Processing of {fileCount} files complete, with {errors} in error")
        End If
    End Sub
    ReadOnly Cabrillo As New Dictionary(Of String, String) From {
            {"ALL-BAND", "ALL"},
            {"ONE-BAND", "SINGLE"},
            {"ONE", "SINGLE"},
            {"FOUR-BAND", "FOUR"},
            {"FOUR-BANDS", "FOUR"},
            {"FIXED", "HOME"}
            }
    Private Shared Function ConvertCabrillo(st As String) As String
        ' Convert alternative Cabrillo clause to standard
        If Form1.Cabrillo.ContainsKey(st) Then Return Form1.Cabrillo(st) Else Return st
    End Function
    Private Shared Function GetQSO(template As String, line As String) As Dictionary(Of String, String)
        ' decode a QSO line in accordance with a template
        ' templates is a series of named templates. The list contains the QSO field names. Fields are separated by multiple spaces

        Dim templates As New Dictionary(Of String, List(Of String)) From {
            {"WIA-VHF/UHF-FD", New List(Of String) From {"freq", "mode", "date", "time", "sent_call", "sent_rst", "sent_exch", "sent_grid", "rcvd_call", "rcvd_rst", "rcvd_exch", "rcvd_grid"}},
            {"VHF-UHF FIELD DAY", New List(Of String) From {"time", "freq", "mode", "rcvd_call", "sent_rst_exch", "rcvd_rst_exch", "rcvd_grid", "sent_grid"}}
            }
        Dim result As New Dictionary(Of String, String)
        result.Clear()
        If Not templates.ContainsKey(template) Then
            Form1.TextBox2.AppendText($"Template {template} is not defined{vbCrLf}")
        Else
            line = Trim(Regex.Replace(line, "\s+", " "))      ' remove multiple spaces
            Dim words = Split(line, " ")     ' split line into words
            If words.Length - 1 < templates(template).Count Then
                Form1.TextBox2.AppendText($"There are too few fields in the QSO")
                Return result
            End If
            For f = 0 To templates(template).Count - 1
                result.Add(templates(template)(f), words(f + 1))
            Next
            ' Sometimes fields are combined, i.e. no space separating them
            ' In this case we extract a combined field, and then split it.
            If result.ContainsKey("sent_rst_exch") Then
                result.Add("sent_rst", result("sent_rst_exch").Substring(0, 2))
                result.Add("sent_exch", result("sent_rst_exch").Substring(2, 3))
                result.Remove("sent_rst_exch")      ' remove combined field
            End If
            If result.ContainsKey("rcvd_rst_exch") Then
                result.Add("rcvd_rst", result("rcvd_rst_exch").Substring(0, 2))
                result.Add("rcvd_exch", result("rcvd_rst_exch").Substring(2, 3))
                result.Remove("rcvd_rst_exch")     ' remove combined field
            End If
        End If
        Return result
    End Function
    Private Shared Sub QSOFixup(QSO As Dictionary(Of String, String))
        ' Do some fixups to prevent silly errors causing mismatch
        For Each key In QSO.Keys
            QSO(key) = QSO(key).ToUpper.Trim           ' all in upper case and trimmed
        Next
        QSO("freq") = Regex.Replace(QSO("freq"), "\.G", "G")    ' many occurences of .G where the dot is not to spec
        If QSO("freq") = "120G" Then QSO("freq") = "122G"   ' 120G is an accepted abbreviation for 122G
        ' Sometimes frequecies above 30MHz are presented in kHz. The Cabrillo spec uses MHz for above 30MHz
        If IsNumeric(QSO("freq")) Then
            Dim freq As Long = CLng(QSO("freq"))     ' no decimal points
            Select Case freq
                Case 50 To 54
                    freq = 50 * MHz
                Case 144 To 148
                    freq = 144 * MHz
                Case 432 To 438
                    freq = 432 * MHz
                Case Else
                    freq *= 1000      ' frequency is KHz
            End Select
            ' search for band assuming freq is integer Hz
            For Each band As KeyValuePair(Of String, (low As Long, high As Long)) In Form1.bandRange
                If freq >= band.Value.low And freq <= band.Value.high Then
                    QSO("freq") = band.Key
                    Exit For
                End If
            Next
        End If
        ' Fix mode
        If QSO("mode") = "SSB" Then QSO("mode") = "PH"
        If QSO("mode") = "DIG" Then QSO("mode") = "DG"

        If QSO("datetime").Length > 16 Then
            ' Round date to nearest minute
            Dim dt = DateTime.Parse(QSO("datetime"))
            Dim ts = TimeSpan.FromMinutes(1)
            Dim Roundeddt = New DateTime(((dt.Ticks + (ts.Ticks / 2)) / ts.Ticks) * ts.Ticks)
            QSO("datetime") = Roundeddt.ToString("yyyy-MM-dd HH:mm")
        End If
    End Sub

    ' list of error flags. A good QSO will have no flags set
    <Flags()>
    Public Enum flagsEnum As Integer
        LoggedIncorrectCall = 1         ' receiver got callsign wrong
        LoggedIncorrectExchange = 2     ' receiver got exchange wrong
        LoggedIncorrectBand = 4         ' band differs between sender and receiver
        LoggedIncorrectLocator = 8      ' receiver got locator wrong
        OutsideContestHours = 16    ' outside of contest hours
        BadGrid = 32           ' bad gridsquare
        NonPermittedBand = 64
        NonPermittedMode = 128
        DuplicateQSO = 256        ' duplicate QSO within window
        NotInLog = 512            ' QSO not in log
        Outside8 = 1024           ' outside 8 hour window
    End Enum
    Const DisqualifyQSO = flagsEnum.OutsideContestHours Or flagsEnum.NonPermittedBand Or flagsEnum.NonPermittedMode Or flagsEnum.DuplicateQSO Or flagsEnum.LoggedIncorrectBand Or flagsEnum.LoggedIncorrectExchange Or flagsEnum.LoggedIncorrectLocator Or flagsEnum.Outside8 ' any of these flags disqualify a QSO
    Const ReworkWindow = 2       ' hours for duplicate window
    Const TimeTolerance = 10   ' times must be +/- minutes to match
    Private Sub CheckScoreLogsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CheckScoreLogsToolStripMenuItem.Click
        ' check logs and produce scores

        Dim contestID As Integer

        Dim sql As SqliteCommand, sqldr As SqliteDataReader
        Dim sqlQSO As SqliteCommand, sqlQSOdr As SqliteDataReader
        Dim Contestssql As SqliteCommand, Contestsdr As SqliteDataReader
        Dim Stationssql As SqliteCommand
        Dim updsql As SqliteCommand
        Dim count As Integer = 0, updated As Integer, TotalQSO As Integer

        If dlgContest.ShowDialog = DialogResult.OK Then
            contestID = dlgContest.DataGridView1.SelectedRows(0).Cells("id").Value
            Using connect As New SqliteConnection(CheckerDB)
                connect.Open()
                connect.CreateFunction("REGEXP", Function(pattern As String, Input As String) Regex.IsMatch(Input, pattern))    ' define a regexp function
                connect.CreateFunction("BASECALL", Function(input As String) basecall(input))       ' define a function to remove suffix from call
                connect.CreateFunction("DISTANCE", Function(A As String, B As String) GridDistance(A, B))
                connect.CreateFunction("SCORE", Function(band As String, distance As Integer) Score(band, distance))
                connect.CreateFunction("FREQUENCY", Function(band As String) frequency(band))
                Dim tr As SqliteTransaction = connect.BeginTransaction
                ' retrieve contest details
                Contestssql = connect.CreateCommand
                Contestssql.Transaction = tr
                Contestssql.CommandText = $"SELECT * FROM `Contests` WHERE `ContestID`={contestID}"
                Contestsdr = Contestssql.ExecuteReader()
                Contestsdr.Read()
                sql = connect.CreateCommand
                sql.Transaction = tr
                sqlQSO = connect.CreateCommand
                sqlQSO.Transaction = tr
                updsql = connect.CreateCommand
                updsql.Transaction = tr
                Stationssql = connect.CreateCommand
                Stationssql.Transaction = tr
                sql.CommandText = $"SELECT COUNT(*) AS Total FROM `QSO` WHERE `contestID`={contestID}"
                sqldr = sql.ExecuteReader()
                sqldr.Read()
                TotalQSO = sqldr("Total")
                sqldr.Close()
                TextBox2.AppendText($"Contest hours - {Contestsdr("Start")} UTC to {Contestsdr("Finish")} UTC{vbCrLf}")
                TextBox2.AppendText($"Permitted bands {Contestsdr("PermittedBands")}{vbCrLf}")
                TextBox2.AppendText($"Permitted modes {Contestsdr("PermittedModes")}{vbCrLf}")
                TextBox2.AppendText($"{vbCrLf}")

                ' Remove all existing flags and QSO matches
                StopWatch.Restart()
                sql.CommandText = $"UPDATE `QSO` SET `flags`=0,`match`=NULL, `score`=NULL, `distance`=NULL WHERE `contestID`={contestID}"
                count = sql.ExecuteNonQuery()
                StopWatch.Stop()
                TextBox2.AppendText($"Initialise {count} QSO took {StopWatch.Elapsed.Seconds}s{vbCrLf}")

                ' find all perfectly matching QSO (by time and call, band, exch and grid)
                ' The presumption is that most should match, so this makes finding ones with copy errors easier

                ' parameterize the match conditions so we can make appropriate combinations of them
                Dim callMatch As String = "basecall(A.rcvd_call)=basecall(B.sent_call) And basecall(B.rcvd_call)=basecall(A.sent_call)"
                Dim bandMatch As String = "A.band=B.band"
                Dim timeMatch As String = $"DATETIME(B.date) BETWEEN DATETIME(A.date,'-{TimeTolerance} minutes') AND DATETIME(A.date,'+{TimeTolerance} minutes')"
                Dim exchMatch As String = "A.sent_exch=B.rcvd_exch AND A.rcvd_exch=B.sent_exch"
                Dim gridMatch As String = "A.sent_grid=B.rcvd_grid AND A.rcvd_grid=B.sent_grid"

                StopWatch.Restart()
                count = 0
                sql.CommandText = $"SELECT     A.id  AS Aid,
                                               B.id  AS Bid
                                    FROM       `QSO` AS A
                                    INNER JOIN `QSO` AS B
                                    ON         A.contestID=B.contestID
                                    AND        {timeMatch}
                                    AND        {callMatch}
                                    AND        {bandMatch}
                                    AND        {exchMatch}
                                    AND        {gridMatch}
                                    WHERE      A.contestID={contestID}
                                    AND        A.id>B.id"
                sqldr = sql.ExecuteReader()
                While sqldr.Read
                    sqlQSO.CommandText = $"UPDATE `QSO` SET `match`={sqldr("Bid")} WHERE id={sqldr("Aid")}"
                    sqlQSO.ExecuteNonQuery()
                    sqlQSO.CommandText = $"UPDATE `QSO` SET `match`={sqldr("Aid")} WHERE id={sqldr("Bid")}"
                    sqlQSO.ExecuteNonQuery()
                    count += 2
                End While
                sqldr.Close()
                StopWatch.Stop()
                TextBox2.AppendText($"QSO perfect match analysis found {count} and took {StopWatch.Elapsed.Seconds}s{vbCrLf}")
                Application.DoEvents()    ' let textbox update

                ' Now look for not matched QSO that match on 3 out of 4 values - band, call, exch, grid
                ' Look for a matching pair of QSO, which match on 3 of the 4 criteria - call them A and B.
                ' if the B received value <> A sent value then B is wrong
                ' if the A received value <> B sent value then A is wrong
                ' for efficiency, constrain id of A > id of B to avoid processing all QSO twice

                ' Call
                StopWatch.Restart()
                count = 0
                sql.CommandText = $"SELECT     A.id        AS Aid,
                                               B.id        AS Bid,
                                               A.sent_call AS Asent_call,
                                               B.rcvd_call AS Brcvd_call,
                                               A.rcvd_call AS Arcvd_call,
                                               B.sent_call AS Bsent_call
                                    FROM       `QSO`       AS A
                                    INNER JOIN `QSO`       AS B
                                    ON         A.contestID=B.contestID
                                    AND        {timeMatch}
                                    AND        {exchMatch}
                                    AND        {bandMatch}
                                    AND        {gridMatch}
                                    WHERE      A.contestID={contestID}
                                    AND        A.match IS NULL
                                    AND        B.match IS NULL
                                    AND        A.id>B.id"
                sqldr = sql.ExecuteReader()
                While sqldr.Read
                    sqlQSO.CommandText = $"UPDATE `QSO` SET `match`={sqldr("Bid")} WHERE id={sqldr("Aid")}"     ' these two QSO match
                    sqlQSO.ExecuteNonQuery()
                    sqlQSO.CommandText = $"UPDATE `QSO` SET `match`={sqldr("Aid")} WHERE id={sqldr("Bid")}"     ' these two QSO match
                    sqlQSO.ExecuteNonQuery()
                    count += 2
                    If basecall(sqldr("Asent_call")) <> basecall(sqldr("Brcvd_call")) Then
                        ' B is wrong. flag error
                        sqlQSO.CommandText = $"UPDATE `QSO` SET `flags`=`flags` | {CInt(flagsEnum.LoggedIncorrectCall)} WHERE id={sqldr("Bid")}"
                        sqlQSO.ExecuteNonQuery()
                    End If
                    If basecall(sqldr("Bsent_call")) <> basecall(sqldr("Arcvd_call")) Then
                        ' A is wrong. flag error
                        sqlQSO.CommandText = $"UPDATE `QSO` SET `flags`=`flags` | {CInt(flagsEnum.LoggedIncorrectCall)} WHERE id={sqldr("Aid")}"
                        sqlQSO.ExecuteNonQuery()
                    End If
                End While
                sqldr.Close()
                StopWatch.Stop()
                TextBox2.AppendText($"Mismatched call analysis found {count} and took {StopWatch.Elapsed.Seconds}s{vbCrLf}")
                Application.DoEvents()    ' let textbox update

                ' Exchange
                StopWatch.Restart()
                count = 0
                sql.CommandText = $"SELECT     A.id        AS Aid,
                                               B.id        AS Bid,
                                               A.sent_exch AS Asent_exch,
                                               B.rcvd_exch AS Brcvd_exch,
                                               A.rcvd_exch AS Arcvd_exch,
                                               B.sent_exch AS Bsent_exch
                                        FROM       `QSO`       AS A
                                        INNER JOIN `QSO`       AS B
                                        ON         A.contestID=B.contestID
                                        AND        {timeMatch}   
                                        AND        {callMatch}
                                        AND        {bandMatch}
                                        AND        {gridMatch}
                                        WHERE      A.contestID={contestID}
                                        AND        A.match IS NULL
                                        AND        B.match IS NULL
                                        AND        A.id>B.id"
                sqldr = sql.ExecuteReader()
                While sqldr.Read
                    sqlQSO.CommandText = $"UPDATE `QSO` SET `match`={sqldr("Bid")} WHERE id={sqldr("Aid")}"     ' these two QSO match
                    sqlQSO.ExecuteNonQuery()
                    sqlQSO.CommandText = $"UPDATE `QSO` SET `match`={sqldr("Aid")} WHERE id={sqldr("Bid")}"     ' these two QSO match
                    sqlQSO.ExecuteNonQuery()
                    count += 2
                    If sqldr("Asent_exch") <> sqldr("Brcvd_exch") Then
                        ' B is wrong. flag error
                        sqlQSO.CommandText = $"UPDATE `QSO` SET `flags`=`flags` | {CInt(flagsEnum.LoggedIncorrectExchange)} WHERE id={sqldr("Bid")}"
                        sqlQSO.ExecuteNonQuery()
                    End If
                    If sqldr("Bsent_exch") <> sqldr("Arcvd_exch") Then
                        ' A is wrong. flag error
                        sqlQSO.CommandText = $"UPDATE `QSO` SET `flags`=`flags` | {CInt(flagsEnum.LoggedIncorrectExchange)} WHERE id={sqldr("Aid")}"
                        sqlQSO.ExecuteNonQuery()
                    End If
                End While
                sqldr.Close()
                StopWatch.Stop()
                TextBox2.AppendText($"Mismatched exchange analysis found {count} and took {StopWatch.Elapsed.Seconds}s{vbCrLf}")
                Application.DoEvents()    ' let textbox update

                ' Grid
                StopWatch.Restart()
                count = 0
                sql.CommandText = $"SELECT     A.id        AS Aid,
                                               B.id        AS Bid,
                                               A.sent_grid AS Asent_grid,
                                               B.rcvd_grid AS Brcvd_grid,
                                               A.rcvd_grid AS Arcvd_grid,
                                               B.sent_grid AS Bsent_grid
                                    FROM       `QSO`       AS A
                                    INNER JOIN `QSO`       AS B
                                    ON         A.contestID=B.contestID
                                    AND        {timeMatch}  
                                    AND        {callMatch}
                                    AND        {bandMatch}
                                    AND        {exchMatch}
                                    WHERE      A.contestID={contestID}
                                    AND        A.match IS NULL
                                    AND        B.match IS NULL
                                    AND        A.id>B.id"
                sqldr = sql.ExecuteReader()
                While sqldr.Read
                    sqlQSO.CommandText = $"UPDATE `QSO` SET `match`={sqldr("Bid")} WHERE id={sqldr("Aid")}"     ' these two QSO match
                    sqlQSO.ExecuteNonQuery()
                    sqlQSO.CommandText = $"UPDATE `QSO` SET `match`={sqldr("Aid")} WHERE id={sqldr("Bid")}"     ' these two QSO match
                    sqlQSO.ExecuteNonQuery()
                    count += 2
                    If sqldr("Asent_grid") <> sqldr("Brcvd_grid") Then
                        ' B is wrong. flag error
                        sqlQSO.CommandText = $"UPDATE `QSO` SET `flags`=`flags` | {CInt(flagsEnum.LoggedIncorrectLocator)} WHERE id={sqldr("Bid")}"
                        sqlQSO.ExecuteNonQuery()
                    End If
                    If sqldr("Bsent_grid") <> sqldr("Arcvd_grid") Then
                        ' A is wrong. flag error
                        sqlQSO.CommandText = $"UPDATE `QSO` SET `flags`=`flags` | {CInt(flagsEnum.LoggedIncorrectLocator)} WHERE id={sqldr("Aid")}"
                        sqlQSO.ExecuteNonQuery()
                    End If
                End While
                sqldr.Close()
                StopWatch.Stop()
                TextBox2.AppendText($"Mismatched gridsquare analysis found {count} and took {StopWatch.Elapsed.Seconds}s{vbCrLf}")
                Application.DoEvents()    ' let textbox update

                ' BAND
                StopWatch.Restart()
                count = 0
                sql.CommandText = $"SELECT     A.id   AS Aid,
                                               B.id   AS Bid,
                                               A.band AS Aband,
                                               B.band AS Bband
                                    FROM       `QSO`  AS A
                                    INNER JOIN `QSO`  AS B
                                    ON         A.contestID=B.contestID
                                    AND        {timeMatch}
                                    AND        {callMatch}
                                    AND        {gridMatch}
                                    AND        {exchMatch}
                                    WHERE      A.contestID={contestID}
                                    AND        A.match IS NULL
                                    AND        B.match IS NULL
                                    AND        A.id>B.id"
                sqldr = sql.ExecuteReader()
                While sqldr.Read
                    sqlQSO.CommandText = $"UPDATE `QSO` SET `match`={sqldr("Bid")} WHERE id={sqldr("Aid")}"     ' these two QSO match
                    sqlQSO.ExecuteNonQuery()
                    sqlQSO.CommandText = $"UPDATE `QSO` SET `match`={sqldr("Aid")} WHERE id={sqldr("Bid")}"     ' these two QSO match
                    sqlQSO.ExecuteNonQuery()
                    count += 2
                    If sqldr("Aband") <> sqldr("Bband") Then
                        ' both wrong. flag error
                        sqlQSO.CommandText = $"UPDATE `QSO` SET `flags`=`flags` | {CInt(flagsEnum.LoggedIncorrectBand)} WHERE id IN ({sqldr("Aid")},{sqldr("Bid")})"
                        sqlQSO.ExecuteNonQuery()
                    End If
                End While
                sqldr.Close()
                StopWatch.Stop()
                TextBox2.AppendText($"Mismatched band analysis found {count} and took {StopWatch.Elapsed.Seconds}s{vbCrLf}")
                Application.DoEvents()    ' let textbox update

                ' as a last resort, match on date, band and calls only. This will mean that both exchange and grid must have a mismatch
                StopWatch.Restart()
                count = 0
                sql.CommandText = $"SELECT     A.id   AS Aid,
                                               B.id   AS Bid,
                                               A.band AS Aband,
                                               B.band AS Bband,
                                               A.sent_exch AS Asent_exch,
                                               B.rcvd_exch AS Brcvd_exch,
                                               A.rcvd_exch AS Arcvd_exch,
                                               B.sent_exch AS Bsent_exch,
                                               A.sent_grid AS Asent_grid,
                                               B.rcvd_grid AS Brcvd_grid,
                                               A.rcvd_grid AS Arcvd_grid,
                                               B.sent_grid AS Bsent_grid
                                    FROM       `QSO`  AS A
                                    INNER JOIN `QSO`  AS B
                                    ON         A.contestID=B.contestID
                                    AND        {timeMatch}
                                    AND        {callMatch}
                                    AND        {bandMatch}
                                    WHERE      A.contestID={contestID}
                                    AND        A.match IS NULL
                                    AND        B.match IS NULL
                                    AND        A.id>B.id"
                sqldr = sql.ExecuteReader()
                While sqldr.Read
                    sqlQSO.CommandText = $"UPDATE `QSO` SET `match`={sqldr("Bid")} WHERE id={sqldr("Aid")}"     ' these two QSO match
                    sqlQSO.ExecuteNonQuery()
                    sqlQSO.CommandText = $"UPDATE `QSO` SET `match`={sqldr("Aid")} WHERE id={sqldr("Bid")}"     ' these two QSO match
                    sqlQSO.ExecuteNonQuery()
                    count += 2
                    ' could be a grid mismatch
                    If sqldr("Asent_grid") <> sqldr("Brcvd_grid") Then
                        ' B is wrong. flag error
                        sqlQSO.CommandText = $"UPDATE `QSO` SET `flags`=`flags` | {CInt(flagsEnum.LoggedIncorrectLocator)} WHERE id={sqldr("Bid")}"
                        sqlQSO.ExecuteNonQuery()
                    End If
                    If sqldr("Bsent_grid") <> sqldr("Arcvd_grid") Then
                        ' A is wrong. flag error
                        sqlQSO.CommandText = $"UPDATE `QSO` SET `flags`=`flags` | {CInt(flagsEnum.LoggedIncorrectLocator)} WHERE id={sqldr("Aid")}"
                        sqlQSO.ExecuteNonQuery()
                    End If
                    ' could be exchange mismatch
                    If sqldr("Asent_exch") <> sqldr("Brcvd_exch") Then
                        ' B is wrong. flag error
                        sqlQSO.CommandText = $"UPDATE `QSO` SET `flags`=`flags` | {CInt(flagsEnum.LoggedIncorrectExchange)} WHERE id={sqldr("Bid")}"
                        sqlQSO.ExecuteNonQuery()
                    End If
                    If sqldr("Bsent_exch") <> sqldr("Arcvd_exch") Then
                        ' A is wrong. flag error
                        sqlQSO.CommandText = $"UPDATE `QSO` SET `flags`=`flags` | {CInt(flagsEnum.LoggedIncorrectExchange)} WHERE id={sqldr("Aid")}"
                        sqlQSO.ExecuteNonQuery()
                    End If
                End While
                sqldr.Close()
                StopWatch.Stop()
                TextBox2.AppendText($"Last resort QSO matching found {count} and took {StopWatch.Elapsed.Seconds}s{vbCrLf}")
                Application.DoEvents()    ' let textbox update

                ' test gridsquares. Gridsquares can be 6 or 8 characters
                StopWatch.Restart()
                updsql.CommandText = $"UPDATE `QSO` SET `flags`=`flags` | {CInt(flagsEnum.BadGrid)} WHERE `contestID`={contestID} AND (`sent_grid` NOT REGEXP('^[A-R][A-R][0-9][0-9][A-X][A-X]([0-9][0-9])?$') OR `rcvd_grid` NOT REGEXP('^[A-R][A-R][0-9][0-9][A-X][A-X]([0-9][0-9])?$'))"
                count = updsql.ExecuteNonQuery()
                StopWatch.Stop()
                TextBox2.AppendText($"Incorrect grid square analysis found {count} and took {StopWatch.Elapsed.Seconds}s{vbCrLf}")
                If count > 0 Then
                    ' Display outside hours
                    TextBox2.AppendText($"QSO with bad gridsquare{vbCrLf}")
                    sql.CommandText = $"SELECT * FROM `QSO` WHERE `contestID`={contestID} AND `flags` & {CInt(flagsEnum.BadGrid)} <>0 ORDER BY sent_call,date"
                    sqldr = sql.ExecuteReader()
                    While sqldr.Read()
                        TextBox2.AppendText($"{sqldr("date")} {sqldr("band"),4} {sqldr("mode")} {sqldr("sent_call"),-10} {sqldr("sent_rst")}{sqldr("sent_exch")} {sqldr("sent_grid"),-8} {sqldr("rcvd_call"),-10} {sqldr("rcvd_rst")}{sqldr("rcvd_exch")} {sqldr("rcvd_grid")}{vbCrLf}")
                    End While
                    sqldr.Close()
                    TextBox2.AppendText($"{vbCrLf}")
                End If

                ' test for outside contest hours
                StopWatch.Restart()
                updsql.CommandText = $"UPDATE `QSO` SET `flags`=`flags` | {CInt(flagsEnum.OutsideContestHours)} WHERE `contestID`={contestID} AND (substr(`sent_call`,3,1)<>'6' AND NOT DATETIME(date) BETWEEN DATETIME('{Contestsdr("Start")}') AND DATETIME('{Contestsdr("Finish")}') OR (substr(`sent_call`,3,1)='6' AND NOT DATETIME(date) BETWEEN DATETIME('{Contestsdr("Start")}','+3 hours') AND DATETIME('{Contestsdr("Finish")}','+3 hours')))"
                count = updsql.ExecuteNonQuery()
                StopWatch.Stop()
                TextBox2.AppendText($"QSO outside contest hours analysis found {count} and took {StopWatch.Elapsed.Seconds}s{vbCrLf}")
                If count > 0 Then
                    ' Display outside hours
                    TextBox2.AppendText($"QSO outside contest hours{vbCrLf}")
                    sql.CommandText = $"SELECT * FROM `QSO` WHERE `contestID`={contestID} AND `flags` & {CInt(flagsEnum.OutsideContestHours)} <>0 ORDER BY sent_call,date"
                    sqldr = sql.ExecuteReader()
                    While sqldr.Read()
                        TextBox2.AppendText($"{sqldr("date")} {sqldr("band"),4} {sqldr("mode")} {sqldr("sent_call"),-10} {sqldr("sent_rst")}{sqldr("sent_exch")} {sqldr("sent_grid"),-8} {sqldr("rcvd_call"),-10} {sqldr("rcvd_rst")}{sqldr("rcvd_exch")} {sqldr("rcvd_grid")}{vbCrLf}")
                    End While
                    sqldr.Close()
                    TextBox2.AppendText($"{vbCrLf}")
                End If

                ' test for CATEGORY-BAND (SINGLE, FOUR or ALL)
                sql.CommandText = $"SELECT * FROM `Stations` WHERE `contestID`={contestID} AND `CategoryBand`<>''"
                sqldr = sql.ExecuteReader
                While sqldr.Read
                    Select Case sqldr("CategoryBand")
                        Case "SINGLE"
                            Dim ListofBands As New Dictionary(Of String, Integer)
                            ListofBands.Clear()
                            sqlQSO.CommandText = $"SELECT `band`,COUNT(*) as count FROM Stations AS S JOIN QSO AS Q ON S.station=Q.sent_call AND S.`contestID`=Q.`contestID` WHERE S.`contestID`={contestID} AND S.`station`='{sqldr("station")}' GROUP BY `band`"
                            sqlQSOdr = sqlQSO.ExecuteReader
                            While sqlQSOdr.Read
                                ListofBands.Add($"'{sqlQSOdr("band")}'", sqlQSOdr("count"))
                            End While
                            sqlQSOdr.Close()
                            If ListofBands.Count > 1 Then
                                ' There is more than one band in the log. Pick the band with the most QSO and disqualify the rest
                                Dim SortedList As List(Of String) = (From tPair As KeyValuePair(Of String, Integer) In ListofBands Order By tPair.Value Descending Select tPair.Key).ToList
                                SortedList.RemoveAt(0)      ' remove largest
                                updsql.CommandText = $"UPDATE `QSO` SET `flags`=`flags` | {CInt(flagsEnum.NonPermittedBand)} WHERE `contestID`={contestID} AND basecall(`sent_call`)=basecall('{sqldr("station")}') AND `band` IN ({Strings.Join(SortedList.ToArray, ",")})"
                                updsql.ExecuteNonQuery()
                            End If

                        Case "FOUR"
                            ' Disqualify any band not one of FOUR (50, 144, 432, 1.2G)
                            sqlQSO.CommandText = $"SELECT * FROM Stations AS S JOIN QSO AS Q ON S.station=Q.sent_call AND S.`contestID`=Q.`contestID` WHERE S.`contestID`={contestID} AND `station`= '{sqldr("station")}' AND `band` NOT IN ('50','144','432','1.2G')"
                            sqlQSOdr = sqlQSO.ExecuteReader
                            While sqlQSOdr.Read
                                updsql.CommandText = $"UPDATE `QSO` SET `flags`=`flags` | {CInt(flagsEnum.NonPermittedBand)} WHERE `id`={sqlQSOdr("id")}"
                                updsql.ExecuteNonQuery()
                            End While
                            sqlQSOdr.Close()

                        Case "ALL"
                            ' Do nothing
                        Case Else
                            MsgBox($"Unrecognised CategoryBand of {sqldr("CategoryBand")} for callsign {sqldr("station")}", vbCritical + vbOKOnly, "Bad CATEGORY BAND")
                    End Select
                End While
                sqldr.Close()

                StopWatch.Restart()
                Dim PermittedBands As String
                If Contestsdr("PermittedBands") = "*" Then
                    PermittedBands = QuotedCSV(bandRange.Keys.ToArray)
                Else
                    PermittedBands = QuotedCSV(Split(Contestsdr("PermittedBands"), ","))
                End If
                updsql.CommandText = $"UPDATE `QSO` SET `flags`=`flags` | {CInt(flagsEnum.NonPermittedBand)} WHERE `contestID`={contestID} AND `band` NOT IN ({PermittedBands})"
                count = updsql.ExecuteNonQuery()
                sqldr.Close()
                StopWatch.Stop()
                TextBox2.AppendText($"Non-permitted band analysis found {count} and took {StopWatch.Elapsed.Seconds}s{vbCrLf}")
                If count > 0 Then
                    ' Display bad band
                    TextBox2.AppendText($"QSO on non-permitted band{vbCrLf}")
                    sql.CommandText = $"SELECT * FROM `QSO` WHERE `contestID`={contestID} AND (`flags` & {CInt(flagsEnum.NonPermittedBand)}) <>0 ORDER BY sent_call,date"
                    sqldr = sql.ExecuteReader()
                    While sqldr.Read()
                        TextBox2.AppendText($"{sqldr("date")} {sqldr("band"),4} {sqldr("mode")} {sqldr("sent_call"),-10} {sqldr("sent_rst")}{sqldr("sent_exch")} {sqldr("sent_grid"),-8} {sqldr("rcvd_call"),-10} {sqldr("rcvd_rst")}{sqldr("rcvd_exch")} {sqldr("rcvd_grid")}{vbCrLf}")
                    End While
                    sqldr.Close()
                    TextBox2.AppendText($"{vbCrLf}")
                End If

                ' test for bad mode
                StopWatch.Restart()
                Dim PermittedModes As String
                If Contestsdr("PermittedModes") = "*" Then
                    PermittedModes = QuotedCSV(modeValidation)
                Else
                    PermittedModes = QuotedCSV(Split(Contestsdr("PermittedModes"), ","))
                End If
                updsql.CommandText = $"UPDATE `QSO` SET `flags`=`flags` | {CInt(flagsEnum.NonPermittedMode)} WHERE `contestID`={contestID} AND `mode` NOT IN ({PermittedModes})"
                count = updsql.ExecuteNonQuery()
                sqldr.Close()
                StopWatch.Stop()
                TextBox2.AppendText($"Non-permitted mode analysis found {count} and took {StopWatch.Elapsed.Seconds}s{vbCrLf}")
                Application.DoEvents()    ' let textbox update

                If count > 0 Then
                    ' Display bad mode
                    TextBox2.AppendText($"QSO on non-permitted mode{vbCrLf}")
                    sql.CommandText = $"SELECT * FROM `QSO` WHERE `contestID`={contestID} AND `flags` & {CInt(flagsEnum.NonPermittedMode)} <>0 ORDER BY sent_call,date"
                    sqldr = sql.ExecuteReader()
                    While sqldr.Read()
                        TextBox2.AppendText($"{sqldr("date")} {sqldr("band"),4} {sqldr("mode")} {sqldr("sent_call"),-10} {sqldr("sent_rst")}{sqldr("sent_exch")} {sqldr("sent_grid"),-8} {sqldr("rcvd_call"),-10} {sqldr("rcvd_rst")}{sqldr("rcvd_exch")} {sqldr("rcvd_grid")}{vbCrLf}")
                    End While
                    sqldr.Close()
                    TextBox2.AppendText($"{vbCrLf}")
                End If

                ' test for not in log
                ' NotInLog = there is no QSO match, but there does exist a log for the rcvd_call
                StopWatch.Restart()
                count = 0
                sql.CommandText = $"SELECT *
                                    FROM   QSO AS A
                                    WHERE  A.match IS NULL
                                           AND A.contestID = {contestID}
                                           AND EXISTS(SELECT *
                                                      FROM   QSO AS B
                                                      WHERE  B.contestID = {contestID}
                                                             AND basecall(B.sent_call) = basecall(A.rcvd_call))"
                sqldr = sql.ExecuteReader
                While sqldr.Read
                    updsql.CommandText = $"UPDATE `QSO` SET `flags`=`flags` | {CInt(flagsEnum.NotInLog)} WHERE id={sqldr("id")}"
                    count += updsql.ExecuteNonQuery()
                End While
                sqldr.Close()
                TextBox2.AppendText($"Not in log analysis found {count} and took {StopWatch.Elapsed.Seconds}s{vbCrLf}")
                Application.DoEvents()    ' let textbox update

                ' test for duplicates
                ' and duplicate is 2 QSO which match on calls & band and are within the rework window
                StopWatch.Restart()
                count = 0
                sqlQSO.CommandText = $"Select A.id AS Aid
                                    From QSO As A
                                    Where A.contestID = {contestID}
                                           And EXISTS(SELECT *
                                                      From QSO As B
                                                      Where A.contestID = B.contestID
                                                             And basecall(A.sent_call) = basecall(B.sent_call)
                                                             And basecall(A.rcvd_call) = basecall(B.rcvd_call)
                                                             And A.band = B.band
                                                             And A.date BETWEEN
                                                                 DateTime(B.date) AND
                                                                 DateTime(B.date, '+{ReworkWindow} hours', '-{TimeTolerance} minutes')
                                                             And A.id <> B.id)"
                sqlQSOdr = sqlQSO.ExecuteReader()
                If sqlQSOdr.HasRows Then
                    ' We have duplicate(s)
                    While sqlQSOdr.Read
                        updsql.CommandText = $"UPDATE `QSO` SET `flags`=`flags` | {CInt(flagsEnum.DuplicateQSO)} WHERE id={sqlQSOdr("Aid")}"
                        count += updsql.ExecuteNonQuery()
                    End While
                End If
                sqlQSOdr.Close()
                StopWatch.Stop()
                TextBox2.AppendText($"Duplicate QSO analysis found {count} and took {StopWatch.Elapsed.Seconds}s{vbCrLf}")
                Application.DoEvents()    ' let textbox update

                TextBox2.AppendText($"Duplicate QSO{vbCrLf}")
                sql.CommandText = $"SELECT * FROM `QSO` WHERE contestID={contestID} AND (`flags` & {CInt(flagsEnum.DuplicateQSO)})<>0 ORDER BY sent_call,band,rcvd_call,date"
                sqldr = sql.ExecuteReader()
                While sqldr.Read
                    ' Display duplicates
                    TextBox2.AppendText($"{sqldr("date")} {sqldr("band"),4} {sqldr("sent_call"),-10} {sqldr("rcvd_call"),-10}{vbCrLf}")
                End While
                sqldr.Close()

                ' Update list of active locators
                updsql.CommandText = $"UPDATE `Stations` SET `gridsquare`=(select group_concat(distinct `sent_grid`) as `gridsquare` from qso where `contestID`={contestID} AND Stations.station=sent_call group by sent_call)"
                updsql.ExecuteNonQuery()

                ' Calculate distance (km) where both grids are OK
                StopWatch.Restart()
                updsql.CommandText = $"UPDATE `QSO` SET `distance`= DISTANCE(`sent_grid`,`rcvd_grid`) WHERE `contestID`={contestID} AND (`flags` & {CInt(flagsEnum.BadGrid)})=0"
                updated = updsql.ExecuteNonQuery()
                StopWatch.Stop()
                TextBox2.AppendText($"{updated} distances calculated in {StopWatch.Elapsed.Seconds}s{vbCrLf}")
                Application.DoEvents()    ' let textbox update

                ' Calculate scores where both grids are OK, and band is OK
                StopWatch.Restart()
                updsql.CommandText = $"UPDATE `QSO` SET `score`=SCORE(`band`,`distance`) WHERE distance is not null AND `contestID`={contestID} AND (`flags` & {flagsEnum.BadGrid + flagsEnum.NonPermittedBand})=0"
                updated = updsql.ExecuteNonQuery()
                StopWatch.Stop()
                TextBox2.AppendText($"{updated} scores calculated in {StopWatch.Elapsed.Seconds}s{vbCrLf}")
                Application.DoEvents()    ' let textbox update

                ' for 8-hour section logs, find the best 8 hours (highest score) and disqualify all QSO outside it
                Dim QSOlist As New List(Of QSOdate)
                Dim Hours8 = New TimeSpan(8, 0, 0)          ' 8 hours in seconds
                StopWatch.Restart()
                count = 0
                sql.CommandText = $"SELECT * FROM `Stations` WHERE `contestID`={contestID} AND substr(`section`,2,1)='2'"     ' select all 8 hour logs
                sqldr = sql.ExecuteReader
                While sqldr.Read
                    count += 1
                    QSOlist.Clear()
                    sqlQSO.CommandText = $"SELECT * FROM `QSO` WHERE `contestID`={contestID} AND basecall(`sent_call`)='{sqldr("station")}' AND `section`='{sqldr("section")}' AND `score` IS NOT NULL ORDER BY `date`"     ' all QSO in this log
                    sqlQSOdr = sqlQSO.ExecuteReader
                    If sqlQSOdr.HasRows Then
                        While sqlQSOdr.Read
                            ' collect all QSO for analysis
                            QSOlist.Add(New QSOdate(sqlQSOdr("date"), sqlQSOdr("id"), sqlQSOdr("score")))
                        End While
                        sqlQSOdr.Close()
                        ' Now search for the highest scoring 8 hour window
                        Dim index As Integer = 0
                        Dim WindowStart As String = QSOlist(index).dte            ' start of window
                        Dim WindowEnd As String = Convert.ToDateTime(WindowStart).Add(Hours8).ToString("yyyy-MM-dd HH:mm")        ' end of 8 hour window
                        Dim BestStart As String = WindowStart       ' best starting date
                        Dim BestScore As Integer = QSOlist.Where(Function(QSO) QSO.dte >= WindowStart And QSO.dte < WindowEnd).Sum(Function(QSO) QSO.score)

                        Do
                            WindowStart = QSOlist(index).dte
                            WindowEnd = Convert.ToDateTime(WindowStart).Add(Hours8).ToString("yyyy-MM-dd HH:mm")
                            Dim score = QSOlist.Where(Function(QSO) QSO.dte >= WindowStart And QSO.dte < WindowEnd).Sum(Function(QSO) QSO.score)
                            If score > BestScore Then
                                ' remember start of best window
                                BestScore = score
                                BestStart = WindowStart
                            End If
                            ' move window forward 1 element
                            index += 1
                        Loop Until WindowEnd > QSOlist.Last().dte
                        ' disqualify all QSO before and after 8 hour window
                        Dim BeforeIds = QSOlist.Where(Function(QSO) QSO.dte < BestStart).Select(Function(QSO) QSO.id).ToArray     ' array of id's before window
                        Dim BestEnd = Convert.ToDateTime(BestStart).Add(Hours8).ToString("yyyy-MM-dd HH:mm")           ' end time of best 8 hours
                        Dim AfterIds = QSOlist.Where(Function(QSO) QSO.dte >= BestEnd).Select(Function(QSO) QSO.id).ToArray      ' array of id's after window
                        Dim AllIds = BeforeIds.Concat(AfterIds).ToArray         ' join before and after list
                        If AllIds.Any Then
                            updsql.CommandText = $"UPDATE `QSO` SET `flags`=`flags` | {CInt(flagsEnum.Outside8)} WHERE `id` IN ({String.Join(",", AllIds)})"
                            updated = updsql.ExecuteNonQuery
                        End If
                    Else
                        MsgBox($"No QSO data to find best 8 hours of for {sqldr("station")}", vbCritical + vbOKOnly, "No QSO to process")
                    End If
                    sqlQSOdr.Close()
                End While
                sqldr.Close()
                StopWatch.Stop()
                TextBox2.AppendText($"Best 8 hour analysis for {count} logs took {StopWatch.Elapsed.Seconds}s{vbCrLf}")
                Application.DoEvents()    ' let textbox update

                ' Display summary results
                TextBox2.AppendText($"{vbCrLf}Check summary for contest {Contestsdr("name")}{vbCrLf}Total QSO {TotalQSO}{vbCrLf}")

                Dim flgs As Array = System.Enum.GetValues(GetType(flagsEnum))  ' get enum values
                For Each flag In flgs
                    sql.CommandText = $"SELECT SUM(IIF(`flags` & {CInt(flag)}<>0,1,0)) AS Total FROM `QSO` WHERE `contestID`={contestID}"
                    sqldr = sql.ExecuteReader()
                    sqldr.Read()
                    TextBox2.AppendText($"Total {flag} - {sqldr("Total")} ({sqldr("Total") / TotalQSO * 100:f1}%){vbCrLf}")
                    sqldr.Close()
                Next

                ' Display matched QSO count
                sql.CommandText = $"SELECT COUNT(*) AS Total FROM `QSO` WHERE `contestID`={contestID} AND `match` is NOT null"
                sqldr = sql.ExecuteReader()
                sqldr.Read()
                TextBox2.AppendText($"Total matched QSO - {sqldr("Total")} ({sqldr("Total") / TotalQSO * 100:f1}%){vbCrLf}")
                sqldr.Close()

                ' Display total disqualified QSO
                sql.CommandText = $"SELECT COUNT(*) AS Total FROM `QSO` WHERE `contestID`={contestID} AND (`flags` & {DisqualifyQSO})<>0"
                sqldr = sql.ExecuteReader()
                sqldr.Read()
                TextBox2.AppendText($"Total disqualified QSO - {sqldr("Total")} ({sqldr("Total") / TotalQSO * 100:f1}%){vbCrLf}")
                sqldr.Close()

                ' Update ActualScore value
                updsql.CommandText = $"UPDATE `Stations` SET `ActualScore`=NULL,`Place`=NULL WHERE contestID={contestID}"    ' remove existing scores
                updated = updsql.ExecuteNonQuery
                updsql.CommandText = $"
UPDATE `Stations` AS S
SET    `ActualScore`=Q.score
FROM   (
                SELECT   sent_call,
                         SUM(score) AS score
                FROM     QSO
                WHERE    (flags & {DisqualifyQSO})=0 AND contestID={contestID}
                GROUP BY basecall(sent_call)) AS Q
WHERE  S.station=basecall(Q.sent_call)
AND    S.`contestID`={contestID}"
                updated = updsql.ExecuteNonQuery
                TextBox2.AppendText($"{updated} Total scores calculated{vbCrLf}")

                ' Calculate placings
                ' Query the actual scores, and have SQLite calculate RANK, partitioned by Category
                sql.CommandText = $"
SELECT station,
       section,
       subsection,
       `ActualScore`,
       RANK()
         OVER (
           PARTITION BY section, subsection
           ORDER BY `ActualScore` DESC) AS r
FROM   Stations
WHERE  `contestID` = {contestID}
       AND ActualScore IS NOT NULL
ORDER  BY section,
          subsection"
                sqldr = sql.ExecuteReader
                While sqldr.Read
                    ' Update all placings
                    updsql.CommandText = $"UPDATE `Stations` SET place={sqldr("r")} WHERE contestID={contestID} AND station='{sqldr("station")}' AND section='{sqldr("section")}' AND subsection='{sqldr("subsection")}'"
                    updsql.ExecuteNonQuery()
                End While
sqldr.close
                tr.Commit()
                connect.Close()
            End Using
        End If
    End Sub
    Class QSOdate
        ' encapsulate a single QSO score. Record date/time, id and score
        Property dte As String   ' date/time
        Property id As Integer    ' the QSO id
        Property score As Integer ' score of the QSO

        Sub New(dte As String, id As Integer, score As Integer)
            _dte = dte
            _id = id
            _score = score
        End Sub
    End Class
    Private Shared Function QuotedCSV(arr As Array) As String
        ' Convert an array of strings to a quoted string. Used in SQL IN () clauses
        Dim l As New List(Of String)
        For Each item In arr
            l.Add($"'{item}'")      ' add quotes around each item
        Next
        Return Join(l.ToArray, ",") ' return comma separted list
    End Function
    Private Shared Function GridDistance(Grid1 As String, Grid2 As String) As Integer
        ' Calculate distance between two grid locators in km.
        ' They could be 4, 6 or 8 in length
        Dim p1 As PointF = GridtoLatLon(Grid1), p2 As PointF = GridtoLatLon(Grid2)
        Return GCDistance(p1, p2)
    End Function
    Private Shared Function GridtoLatLon(grid As String) As PointF
        ' convert a grid square to lat/lon. Could be 4, 6 or 8 long
        Const OneMinute As Double = 1 / 60      ' 1 minute of arc
        Const OneSecond As Double = 1 / (60 * 60)   ' 1 second of arc
        Dim p As New PointF

        Trace.Assert(grid.Length >= 4, $"gridsquare must be minimum 4 characters{Environment.StackTrace}{vbCrLf}")
        p.X = (Asc(grid.Substring(0, 1)) - Asc("A")) * 20
        p.Y = (Asc(grid.Substring(1, 1)) - Asc("A")) * 10
        p.X += (Asc(grid.Substring(2, 1)) - Asc("0")) * 2
        p.Y += (Asc(grid.Substring(3, 1)) - Asc("0")) * 1
        If grid.Length > 4 Then
            p.X += (Asc(grid.Substring(4, 1)) - Asc("A")) * (5 * OneMinute)    ' 5' per
            p.Y += (Asc(grid.Substring(5, 1)) - Asc("A")) * (2.5 * OneMinute)  ' 2.5' per
        End If
        If grid.Length > 6 Then
            p.X += (Asc(grid.Substring(6, 1)) - Asc("0")) * (30 * OneSecond)    ' 30'' per
            p.Y += (Asc(grid.Substring(7, 1)) - Asc("0")) * (15 * OneSecond)    ' 15'' per
        End If
        ' Now correct position to center of gridsquare
        Select Case grid.Length
            Case 4
                p.X += 12 * 5 * OneMinute
                p.Y += 12 * 2.5 * OneMinute
            Case 6
                p.X += 5 * 30 * OneSecond
                p.Y += 5 * 15 * OneSecond
            Case 8      ' do nothing - it's accurate enough
        End Select
        p.X -= 180
        p.Y -= 90
        Return p
    End Function
    Private Shared Function GCDistance(p1 As PointF, p2 As PointF) As Integer
        ' Calculate Great Circle distance between 2 points
        'The Haversine formula according to Dr. Math.
        '    http://mathforum.org/library/drmath/view/51879.html
        'dlon = lon2 - lon1
        'dlat = lat2 - lat1
        'a = (sin(dlat / 2)) ^ 2 + cos(lat1) * cos(lat2) * (sin(dlon / 2)) ^ 2
        'c = 2 * atan2(sqrt(a), sqrt(1 - a))
        'd = R * c
        'Where()
        '     * dlon is the change in longitude
        '     * dlat is the change in latitude
        '     * c is the great circle distance in Radians.
        '     * R is the radius of a spherical Earth.
        '     * The locations of the two points in 
        '            spherical coordinates (longitude and 
        '            latitude) are lon1,lat1 and lon2, lat2.

        Dim dDistance As Double
        Dim dLat1InRad As Double
        Dim dLong1InRad As Double
        Dim dLat2InRad As Double
        Dim dLong2InRad As Double
        Dim dLongitude As Double
        Dim dLatitude As Double
        Dim a As Double
        Dim c As Double

        'convert to Radians
        dLat1InRad = p1.Y * (Math.PI / 180.0)
        dLong1InRad = p1.X * (Math.PI / 180.0)
        dLat2InRad = p2.Y * (Math.PI / 180.0)
        dLong2InRad = p2.X * (Math.PI / 180.0)

        dLongitude = dLong2InRad - dLong1InRad
        dLatitude = dLat2InRad - dLat1InRad

        'Intermediate result a.
        a = Math.Pow(Math.Sin(dLatitude / 2.0), 2.0) + Math.Cos(dLat1InRad) * Math.Cos(dLat2InRad) * Math.Pow(Math.Sin(dLongitude / 2.0), 2.0)

        'Intermediate result c (great circle distance in Radians).
        c = 2.0 * Math.Asin(Math.Sqrt(a))

        'Distance.
        Const EarthRadiusKms = 6376.5
        dDistance = EarthRadiusKms * c
        Return CInt(Math.Round(dDistance))
    End Function
    Private Shared Function Score(band As String, distance As Integer) As Integer
        ' Calculate the score of a QSO
        ' for any km up to threshold1, there are mult1 points
        ' for any km over threshold2, there are mult2 points per threshold2 or part thereof

        Trace.Assert(Form1.ScoreTable.ContainsKey(band), $"Band {band} missing from ScoreTable{Environment.StackTrace}{vbCrLf}")
        Trace.Assert(distance >= 0, $"Distance is <0{Environment.StackTrace}{vbCrLf}")
        Dim ScoreData As (multiplier As Single, threshold1 As Integer, mult1 As Integer, threshold2 As Integer, mult2 As Integer) = Form1.ScoreTable(band)  ' get relevant numbers for this band
        Dim zone1 As Integer = Math.Min(ScoreData.threshold1, distance)         ' km in zone1
        Dim zone2 As Integer = Math.Ceiling(Math.Max(distance - ScoreData.threshold1, 0) / ScoreData.threshold2)  ' km in zone 2
        Return Math.Ceiling(ScoreData.multiplier * (zone1 * ScoreData.mult1 + zone2 * ScoreData.mult2))   ' result is rounded up to next integer
    End Function

    Private Shared Function basecall(callsign As String) As String
        ' Analyse a compound call, e.g. VK5/OE2PAS/P, and return the base callsign, i.e. OE2PAS
        Dim s = Split(callsign, "/")
        If s.Length = 1 Then Return callsign
        Dim Sorted = From p In s Order By p.Length Descending Select p      ' sort by descending length
        Return Sorted(0)    ' return longest part
    End Function
    Private Shared Function frequency(band As String) As Long
        ' Calculate frequency in Hz for band

        Trace.Assert(Form1.bandRange.ContainsKey(band), $"Band {band} missing from bandRange{Environment.StackTrace}{vbCrLf}")
        Return Form1.bandRange(band).low   ' result is band frequency in Hz
    End Function

    Private Sub IndividualResultsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IndividualResultsToolStripMenuItem.Click
        ' produce a check report for an individual callsign
        Dim contestID As Integer, station As String, section As String
        Dim sql As SqliteCommand, sqldr As SqliteDataReader
        Dim sqlContest As SqliteCommand, sqldrContest As SqliteDataReader
        Dim sqlEntrant As SqliteCommand, sqldrEntrant As SqliteDataReader
        Dim sqlQSO As SqliteCommand, sqlQSOdr As SqliteDataReader
        Dim myQSO As SqliteCommand, myQSOdr As SqliteDataReader
        Dim QSOcounts As New List(Of QSOcount)

        If dlgContest.ShowDialog = DialogResult.OK Then
            contestID = dlgContest.DataGridView1.SelectedRows(0).Cells("id").Value
            dlgEntrant.Tag = contestID      ' pass contestID to dialog
            If dlgEntrant.ShowDialog() = DialogResult.OK Then
                station = dlgEntrant.DataGridView1.SelectedRows(0).Cells("station").Value
                section = dlgEntrant.DataGridView1.SelectedRows(0).Cells("section").Value
                Using connect As New SqliteConnection(CheckerDB)
                    connect.Open()
                    connect.CreateFunction("FREQUENCY", Function(band As String) frequency(band))
                    connect.CreateFunction("BASECALL", Function(input As String) basecall(input))     ' define a function to remove prefix/suffix from call
                    sql = connect.CreateCommand
                    sqlContest = connect.CreateCommand
                    sqlEntrant = connect.CreateCommand
                    sqlQSO = connect.CreateCommand
                    myQSO = connect.CreateCommand
                    sqlContest.CommandText = $"Select * FROM Contests WHERE contestID={contestID}"
                    sqldrContest = sqlContest.ExecuteReader()
                    sqldrContest.Read()
                    sqlEntrant.CommandText = $"Select * FROM Stations WHERE contestID={contestID} And station='{station}'"
                    sqldrEntrant = sqlEntrant.ExecuteReader()
                    sqldrEntrant.Read()

                    Using report As New StreamWriter($"{station} - {sqldrContest("name")} check report.html")

                        report.WriteLine("<!DOCTYPE html>
<html>
<style>
 .info table, .info th, .info td {
    border: 1px solid black;
}
 .info {
    border-collapse:collapse
}
 .info td {
    padding: 3px;
}
 .section table, .section th, .section td {
    border: 1px solid black;
}
 .section {
    border-collapse:collapse
}
 .section td {
    padding: 3px;
}
 .section {font-family: Arial, Helvetica, sans-serif; font-size: small;}
 .section td:nth-child(1) {width: 120px;}
 .section td:nth-child(2) {width: 250px;}
 .section td:nth-child(4),.section td:nth-child(5), .section td:nth-child(6), .section td:nth-child(7), .section td:nth-child(8), .section td:nth-child(9), .section td:nth-child(10), .section td:nth-child(11), .section td:nth-child(12),td:nth-child(13),td:nth-child(14),td:nth-child(15),td:nth-child(16) {
    text-align:right; width: 40px;
}
 .aligncenter {
     display:block;
     margin-left:auto;
     margin-right:auto 
}
 .zebra table, .zebra th, .zebra td {
    border: 1px solid lightblue;
}
 .zebra {
    border-collapse:collapse
}
 .zebra tr:nth-child(even) {
    background-color:#efefef;
}
 .zebra tr:nth-child(odd) {
    background-color:#e1e1e1;
}
 .zebra tr.deleted td{
    color:red;
}
 th {
    background-color:#e1e1e1
}
 .center {
    text-align: center;
}
 .right{
    text-align: Right;
}
 .left{
    text-align: Left;
}
 .boxed {
     border: 1px solid green ;
}
 h1 {
    text-align: center
}
td.correct{background-color: lightgreen;}
td.incorrect{background-color: tomato;}
</style>")

                        report.WriteLine($"Contest : {sqldrContest("name")}<br><br>")
                        report.WriteLine("<h2>Entrant</h2>")
                        report.WriteLine("<table")
                        report.WriteLine($"<tr><td>Name</td><td>{sqldrEntrant("name")}</td></tr>")
                        report.WriteLine($"<tr><td>Call</td><td>{station}</td></tr>")
                        If sqldrEntrant("operators") <> "" Then report.WriteLine($"<tr><td>Operators</td><td>{sqldrEntrant("operators")}</td></tr>")
                        report.WriteLine($"<tr><td>Section</td><td>{sqldrEntrant("section")}{sqldrEntrant("subsection")} - {Sections(sqldrEntrant("section"))}, {SubSections(sqldrEntrant("subsection"))}</td></tr>")
                        ' Calculate section entrants
                        sql.CommandText = $"SELECT COUNT(*) as Count FROM Stations WHERE contestID={contestID} AND section='{sqldrEntrant("section")}' AND subsection='{sqldrEntrant("subsection")}' AND place IS NOT NULL"
                        sqldr = sql.ExecuteReader
                        sqldr.Read()
                        Dim count = sqldr("Count")
                        sqldr.Close()
                        report.WriteLine($"<tr><td>Rank in Section</td><td>{nthNumber(sqldrEntrant("place"))} (Entries: {count})</td></tr>")
                        report.WriteLine($"<tr><td>email</td><td>{sqldrEntrant("email")}</td></tr>")
                        report.WriteLine("</table>")

                        report.WriteLine("<h2>Summary</h2>")

                        sqlQSO.CommandText = $"SELECT COUNT(*) AS CountedQSO, SUM(iif((flags & {DisqualifyQSO})=0,1,0)) as FinalQSO FROM `QSO` WHERE contestID={contestID} and basecall(sent_call)=basecall('{sqldrEntrant("station")}')"
                        sqlQSOdr = sqlQSO.ExecuteReader()
                        sqlQSOdr.Read()
                        report.WriteLine("<table>")
                        report.WriteLine($"<tr><td class='right' width=100px>{sqldrEntrant("ClaimedQSO")}</td><td>Claimed QSO (for reference)</td></tr>")
                        report.WriteLine($"<tr><td class='right'>{sqlQSOdr("CountedQSO")}</td><td>Counted QSO before checking</td></tr>")
                        report.WriteLine($"<tr><td class='right'>{sqlQSOdr("FinalQSO")}</td><td>Final QSO after checking</td></tr>")
                        report.WriteLine($"<tr><td class='right'>{(sqlQSOdr("CountedQSO") - sqlQSOdr("FinalQSO")) / sqlQSOdr("CountedQSO") * 100:f1}%</td><td>QSO reduction</td></tr>")
                        report.WriteLine("</table>")
                        sqlQSOdr.Close()

                        sqlQSO.CommandText = $"SELECT SUM(score) AS CalcScore, SUM(iif((flags & {DisqualifyQSO})=0,score,0)) as FinalScore FROM `QSO` WHERE contestID={contestID} and basecall(sent_call)=basecall('{sqldrEntrant("station")}')"
                        sqlQSOdr = sqlQSO.ExecuteReader()
                        sqlQSOdr.Read()
                        report.WriteLine("<br><table>")
                        report.WriteLine($"<tr><td class='right' width=100px>{sqldrEntrant("ClaimedScore")}</td><td>Claimed Score (for reference)</td></tr>")
                        report.WriteLine($"<tr><td class='right'>{sqlQSOdr("CalcScore")}</td><td>Calculated score before checking</td></tr>")
                        report.WriteLine($"<tr><td class='right'>{sqlQSOdr("FinalScore")}</td><td>Final score after checking</td></tr>")
                        report.WriteLine($"<tr><td class='right'>{(sqlQSOdr("CalcScore") - sqlQSOdr("FinalScore")) / sqlQSOdr("CalcScore") * 100:f1}%</td><td>Score reduction</td></tr>")
                        report.WriteLine("</table>")
                        sqlQSOdr.Close()

                        report.WriteLine("<h2>Results by Band</h2>")
                        ' Collect the data
                        sqlQSO.CommandText = $"SELECT band,COUNT(*) as Contacts,SUM(score) AS Claimed,SUM(case when (`flags` & {DisqualifyQSO})=0 then score else 0 end) AS Final,max(distance) as Longest,avg(distance) as Average FROM QSO WHERE contestID={contestID} AND sent_call='{station}' group by band "
                        sqlQSOdr = sqlQSO.ExecuteReader
                        QSOcounts.Clear()

                        While sqlQSOdr.Read
                            QSOcounts.Add(New QSOcount(sqlQSOdr("band"),
                                                   sqlQSOdr("Contacts"),
                                                   SuppressZero(sqlQSOdr("Claimed")),
                                                    SuppressZero(sqlQSOdr("Final")),
                                                    SuppressZero(sqlQSOdr("Longest")),
                                                   If(IsDBNull(sqlQSOdr("Average")), "", CInt(sqlQSOdr("Average")).ToString)))
                        End While
                        sqlQSOdr.Close()
                        QSOcounts.Sort(Function(a, b) frequency(a.Band).CompareTo(frequency(b.Band)))
                        ' display data
                        report.Write("<table class='info'><tr><th>Band</th>")
                        For Each item As QSOcount In QSOcounts
                            report.Write($"<th>{item.Band}</th>")
                        Next
                        report.WriteLine("<tr>")
                        report.Write("<tr><td>Contacts</td>")
                        For Each item As QSOcount In QSOcounts
                            report.Write($"<td class='right'>{item.Contacts}</td>")
                        Next
                        report.WriteLine("<tr>")
                        report.Write("<tr><td>Claimed</td>")
                        For Each item As QSOcount In QSOcounts
                            report.Write($"<td class='right'>{item.Claimed}</td>")
                        Next
                        report.WriteLine("<tr>")
                        report.Write("<tr><td>Final</td>")
                        For Each item As QSOcount In QSOcounts
                            report.Write($"<td class='right'>{item.Final}</td>")
                        Next
                        report.WriteLine("<tr>")
                        report.Write("<tr><td class='right'>Longest (km)</td>")
                        For Each item As QSOcount In QSOcounts
                            report.Write($"<td class='right'>{item.Longest}</td>")
                        Next
                        report.WriteLine("<tr>")
                        report.Write("<tr><td>Average (km)</td>")
                        For Each item As QSOcount In QSOcounts
                            report.Write($"<td class='right'>{item.Average}</td>")
                        Next
                        report.WriteLine("<tr>")
                        report.WriteLine("</table>")

                        report.WriteLine("<h2>Not in log (QSO Removed)</h2>")
                        sqlQSO.CommandText = $"SELECT * FROM `QSO` WHERE `contestID`={contestID} AND `sent_call`='{station}' AND (`flags` & {CInt(flagsEnum.NotInLog)}) <> 0 ORDER BY `date`"
                        sqlQSOdr = sqlQSO.ExecuteReader
                        If sqlQSOdr.HasRows Then
                            report.WriteLine($"<table class='info'><tr><th>Date</th><th>Band</th><th>Sent</th><th>Rcvd</th><th>Call</th><th>Impact</th></tr>")
                            While sqlQSOdr.Read
                                report.WriteLine($"<tr><td>{sqlQSOdr("date")}</td><td class='right'>{sqlQSOdr("band")}</td><td>{sqlQSOdr("sent_exch")}</td><td>{sqlQSOdr("rcvd_exch")}</td><td>{sqlQSOdr("rcvd_call")}</td><td class='right'>-{sqlQSOdr("score")} pts</td></tr>")
                            End While
                            report.WriteLine("</table>")
                        Else
                            report.WriteLine("None<br>")
                        End If
                        sqlQSOdr.Close()

                        report.WriteLine("<h2>Duplicate contact (QSO Removed)</h2>")
                        sqlQSO.CommandText = $"SELECT * FROM `QSO` WHERE `contestID`={contestID} AND `sent_call`='{station}' AND (`flags` & {CInt(flagsEnum.DuplicateQSO)}) <> 0 ORDER BY `date`"
                        sqlQSOdr = sqlQSO.ExecuteReader
                        If sqlQSOdr.HasRows Then
                            report.WriteLine($"<table class='info'><tr><th>Date</th><th>Band</th><th>Sent</th><th>Rcvd</th><th>Call</th><th>Impact</th></tr>")
                            While sqlQSOdr.Read
                                report.WriteLine($"<tr><td>{sqlQSOdr("date")}</td><td class='right'>{sqlQSOdr("band")}</td><td>{sqlQSOdr("sent_exch")}</td><td>{sqlQSOdr("rcvd_exch")}</td><td>{sqlQSOdr("rcvd_call")}</td><td class='right'>-{sqlQSOdr("score")} pts</td></tr>")
                            End While
                            report.WriteLine("</table>")
                        Else
                            report.WriteLine("None<br>")
                        End If
                        sqlQSOdr.Close()

                        report.WriteLine("<h2>Call Copied Incorrectly (QSO removed)</h2>")
                        sqlQSO.CommandText = $"SELECT   *,
                                                            A.rcvd_call AS Arcvd_call,
                                                            B.sent_call AS Bsent_call,
                                                            A.rcvd_exch AS Arcvd_exch,
                                                            A.sent_exch AS Asent_exch
                                                FROM     `QSO`       AS A
                                                JOIN     `QSO`       AS B
                                                ON       A.id=B.match
                                                WHERE    A.contestID={contestID}
                                                AND      basecall(A.sent_call)='{station}'
                                                AND      (A.flags & {CInt(flagsEnum.LoggedIncorrectCall)}) <> 0
                                                ORDER BY `date`"
                        sqlQSOdr = sqlQSO.ExecuteReader
                        If sqlQSOdr.HasRows Then
                            report.WriteLine($"<table class='info'><tr><th>Date</th><th>Band</th><th>Mode</th><th>Call</th><th>Sent</th><th>Rcvd</th><th>Correct</th><th>Impact</th></tr>")
                            While sqlQSOdr.Read
                                report.WriteLine($"<tr><td>{sqlQSOdr("date")}</td><td class='right'>{sqlQSOdr("band")}</td><td class='center'>{sqlQSOdr("mode")}</td><td class='incorrect center'>{sqlQSOdr("Arcvd_call")}</td><td>{sqlQSOdr("Asent_exch")}</td><td>{sqlQSOdr("Arcvd_exch")}</td><td class='correct'>{sqlQSOdr("Bsent_call")}</td><td class='right'>-{sqlQSOdr("score")} pts</td></tr>")
                            End While
                            report.WriteLine("</table>")
                        Else
                            report.WriteLine("None<br>")
                        End If
                        sqlQSOdr.Close()

                        report.WriteLine("<h2>Exchange Copied Incorrectly (QSO Removed)</h2>")
                        sqlQSO.CommandText = $"SELECT   *,
                                                         A.rcvd_call AS Arcvd_call,
                                                         B.sent_call AS Bsent_call,
                                                         A.rcvd_exch AS Arcvd_exch,
                                                         A.sent_exch AS Asent_exch,
                                                         B.sent_exch AS Bsent_exch
                                                FROM     `QSO`       AS A
                                                JOIN     `QSO`       AS B
                                                ON       A.id=B.match
                                                WHERE    A.contestID={contestID}
                                                AND      basecall(A.sent_call)='{station}'
                                                AND      (A.flags & {CInt(flagsEnum.LoggedIncorrectExchange)}) <> 0
                                                ORDER BY `date`"
                        sqlQSOdr = sqlQSO.ExecuteReader
                        If sqlQSOdr.HasRows Then
                            report.WriteLine($"<table class='info'><tr><th>Date</th><th>Band</th><th>Mode</th><th>Call</th><th>Sent</th><th>Rcvd</th><th>Correct</th><th>Impact</th></tr>")
                            While sqlQSOdr.Read
                                report.WriteLine($"<tr><td>{sqlQSOdr("date")}</td><td class='right'>{sqlQSOdr("band")}</td><td class='center'>{sqlQSOdr("mode")}</td><td class='center'>{sqlQSOdr("Arcvd_call")}</td><td>{sqlQSOdr("Asent_exch")}</td><td class='incorrect'>{sqlQSOdr("Arcvd_exch")}</td><td class='correct center'>{sqlQSOdr("Bsent_exch")}</td><td class='right'>-{sqlQSOdr("score")} pts</td></tr>")
                            End While
                            report.WriteLine("</table>")
                        Else
                            report.WriteLine("None<br>")
                        End If
                        sqlQSOdr.Close()

                        report.WriteLine("<h2>Exchange Possibly Sent Incorrectly (Information)</h2>")

                        report.WriteLine("<h2>Cross Band Contact (Using Lower Score Band in both logs)</h2>")
                        sqlQSO.CommandText = $"SELECT * FROM `QSO` WHERE `contestID`={contestID} AND basecall(`sent_call`)='{station}' AND (`flags` & {CInt(flagsEnum.LoggedIncorrectBand)}) <> 0 ORDER BY `date`"
                        sqlQSOdr = sqlQSO.ExecuteReader
                        If sqlQSOdr.HasRows Then
                            report.WriteLine($"<table class='info'><tr><th>Date</th><th>Band</th><th>Mode</th><th>Call</th><th>Sent</th><th>Rcvd</th><th>Other log</th><th>Impact</th></tr>")
                            While sqlQSOdr.Read
                                If Not IsDBNull(sqlQSOdr("match")) Then
                                    ' get my matching QSO
                                    myQSO.CommandText = $"SELECT * FROM `QSO` WHERE `id`={sqlQSOdr("match")}"
                                    myQSOdr = myQSO.ExecuteReader
                                    myQSOdr.Read()
                                End If
                                report.WriteLine(
$"<tr>
    <td>{sqlQSOdr("date")}</td>
    <td class='incorrect right'>{sqlQSOdr("band")}</td>
    <td class='center'>{sqlQSOdr("mode")}</td>
    <td>{sqlQSOdr("rcvd_call")}</td>
    <td>{sqlQSOdr("sent_exch")}</td>
    <td>{sqlQSOdr("rcvd_exch")}</td>
    <td class='center'>{myQSOdr("band")}</td>
    <td class='right'>-{sqlQSOdr("score")} pts</td>
 </tr>")
                                If Not myQSOdr.IsClosed Then myQSOdr.Close()
                            End While
                            report.WriteLine("</table>")
                        Else
                            report.WriteLine("None<br>")
                        End If
                        sqlQSOdr.Close()

                        report.WriteLine("<h2>Locator Copied Incorrectly (QSO removed)</h2>")
                        sqlQSO.CommandText = $"SELECT *,A.rcvd_call AS Arcvd_call,B.sent_call AS Bsent_call,A.rcvd_exch AS Arcvd_exch, A.sent_exch as Asent_exch, A.rcvd_grid AS Arcvd_grid, B.sent_grid AS Bsent_grid FROM `QSO` AS A JOIN `QSO` AS B ON A.id=B.match WHERE A.contestID={contestID} AND A.sent_call='{station}' AND (A.flags & {CInt(flagsEnum.LoggedIncorrectLocator)}) <> 0 ORDER BY `date`"
                        sqlQSOdr = sqlQSO.ExecuteReader
                        If sqlQSOdr.HasRows Then
                            report.WriteLine($"<table class='info'><tr><th>Date</th><th>Band</th><th>Mode</th><th>Call</th><th>Sent</th><th>Rcvd</th><th>Locator</th><th>Correct</th><th>Impact</th></tr>")
                            While sqlQSOdr.Read
                                report.WriteLine($"<tr><td>{sqlQSOdr("date")}</td><td class='right'>{sqlQSOdr("band")}</td><td class='center'>{sqlQSOdr("mode")}</td><td class='center'>{sqlQSOdr("Arcvd_call")}</td><td>{sqlQSOdr("Asent_exch")}</td><td>{sqlQSOdr("Arcvd_exch")}</td><td class='incorrect'>{sqlQSOdr("Arcvd_grid")}</td><td class='correct'>{sqlQSOdr("Bsent_grid")}</td><td class='right'>-{sqlQSOdr("score")} pts</td></tr>")
                            End While
                            report.WriteLine("</table>")
                        Else
                            report.WriteLine("None<br>")
                        End If
                        sqlQSOdr.Close()

                        report.WriteLine("<h2>Unique calls (worked once in your log only) (Information - QSO NOT removed)</h2>")
                        sqlQSO.CommandText = $"SELECT * FROM `QSO` AS Q JOIN (SELECT * FROM `QSO` WHERE contestID={contestID} GROUP BY `rcvd_call` HAVING COUNT(*)=1) AS X ON Q.rcvd_call=X.rcvd_call AND Q.contestID=X.contestID WHERE Q.`sent_call`='{station}' ORDER by Q.rcvd_call"
                        sqlQSOdr = sqlQSO.ExecuteReader
                        If sqlQSOdr.HasRows Then
                            report.WriteLine($"<table class='info'><tr><th>Date</th><th>Band</th><th>Mode</th><th>Call</th><th>Sent</th><th>Rcvd</th></tr>")
                            While sqlQSOdr.Read
                                report.WriteLine($"<tr><td>{sqlQSOdr("date")}</td><td class='right'>{sqlQSOdr("band")}</td><td class='center'>{sqlQSOdr("mode")}</td><td>{sqlQSOdr("rcvd_call")}</td><td>{sqlQSOdr("sent_exch")}</td><td>{sqlQSOdr("rcvd_exch")}</td></tr>")
                            End While
                            report.WriteLine("</table>")
                        Else
                            report.WriteLine("None<br>")
                        End If
                        sqlQSOdr.Close()

                        report.WriteLine("<h2>Stations copying your call/band/exchange/locator incorrectly (information)</h2>")
                        Dim BadCopy As Integer = flagsEnum.LoggedIncorrectCall Or flagsEnum.LoggedIncorrectBand Or flagsEnum.LoggedIncorrectExchange Or flagsEnum.LoggedIncorrectLocator Or flagsEnum.NotInLog

                        sqlQSO.CommandText = $"SELECT * FROM `QSO` WHERE `contestID`={contestID} AND basecall(`rcvd_call`)='{station}' AND (`flags` & {BadCopy})<>0 ORDER BY `sent_call`,`date`"
                        sqlQSOdr = sqlQSO.ExecuteReader
                        If sqlQSOdr.HasRows Then
                            report.WriteLine($"<table class='info'><tr><th>Call</th><th>Date</th><th>Band</th><th>Mode</th><th>My Call</th><th>Sent</th><th>Rcvd</th><th>Grid</th><th>Your log</th></tr>")
                            While sqlQSOdr.Read
                                Dim flags As Integer = sqlQSOdr("flags")
                                If Not IsDBNull(sqlQSOdr("match")) Then
                                    ' get my matching QSO
                                    myQSO.CommandText = $"SELECT * FROM `QSO` WHERE `id`={sqlQSOdr("match")}"
                                    myQSOdr = myQSO.ExecuteReader
                                    myQSOdr.Read()
                                End If
                                Dim comment As String
                                Dim CallClass As String = ""
                                If (flags And CInt(flagsEnum.NotInLog)) <> 0 Then
                                    CallClass = " class='incorrect'"
                                    comment = "Not in your log"
                                End If
                                Dim BandClass As String = ""
                                If (flags And CInt(flagsEnum.LoggedIncorrectBand)) <> 0 Then
                                    BandClass = " class='incorrect'"
                                    comment = myQSOdr("band")
                                End If
                                Dim RcvdExchClass As String = ""
                                If (flags And CInt(flagsEnum.LoggedIncorrectExchange)) <> 0 Then
                                    RcvdExchClass = " class='incorrect'"
                                    comment = myQSOdr("sent_exch")
                                End If
                                Dim RcvdGridClass As String = ""
                                If (flags And CInt(flagsEnum.LoggedIncorrectLocator)) <> 0 Then
                                    RcvdGridClass = " class='incorrect'"
                                    comment = myQSOdr("sent_grid")
                                End If
                                report.WriteLine($"<tr><td{CallClass}>{sqlQSOdr("sent_call")}</td><td>{sqlQSOdr("date")}</td><td{BandClass}>{sqlQSOdr("band")}</td><td>{sqlQSOdr("mode")}</td><td>{sqlQSOdr("rcvd_call")}</td><td>{sqlQSOdr("sent_exch")}</td><td{RcvdExchClass}>{sqlQSOdr("rcvd_exch")}</td><td{RcvdGridClass}>{sqlQSOdr("rcvd_grid")}</td><td class='correct'>{comment}</td></tr>")
                                If Not myQSOdr.IsClosed Then myQSOdr.Close()
                            End While
                            report.WriteLine("</table>")
                        Else
                            report.WriteLine("None<br>")
                        End If
                        sqlQSOdr.Close()

                        report.WriteLine($"<br>End of Report. Created: {Now.ToUniversalTime} UTC by FD Log Checker - Marc Hillman (VK3OHM)<br>")
                        report.WriteLine($"Process Log File name: {sqldrEntrant("filename")}")
                        report.WriteLine("</html>")
                        TextBox2.Text = $"Report produced in file {CType(report.BaseStream, FileStream).Name}{vbCrLf}"
                    End Using
                End Using
            End If
        End If
    End Sub
    Function nthNumber(number As Integer) As String
        ' calculate ordinal suffix for number
        Dim suffix As String = ""
        If number >= 4 And number <= 20 Then
            suffix = "th"
        Else
            Select Case number Mod 10
                Case 1
                    suffix = "st"
                Case 2
                    suffix = "nd"
                Case 3
                    suffix = "rd"
                Case Else
                    suffix = "th"
            End Select
        End If
        Return $"{number}{suffix}"
    End Function
    Private Class QSOcount
        ' class to represent a set of results for the "Results by Band" report
        Property Band As String
        Property Contacts As String
        Property Claimed As String
        Property Final As String
        Property Longest As String
        Property Average As String
        Sub New(band As String, Contacts As String, Claimed As String, Final As String, Longest As String, Average As String)
            _Band = band
            _Contacts = Contacts
            _Claimed = Claimed
            _Final = Final
            _Longest = Longest
            _Average = Average
        End Sub
    End Class

    Private Sub ProvisionalResultsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProvisionalResultsToolStripMenuItem.Click
        ' produce a provisional results report
        Dim contestID As Integer
        Dim sqlContest As SqliteCommand, sqldrContest As SqliteDataReader
        Dim sqlEntrant As SqliteCommand, sqldrEntrant As SqliteDataReader
        Dim sqlQSO As SqliteCommand, sqlQSOdr As SqliteDataReader
        Dim QSOcounts As New List(Of QSOcount)
        Dim bandList As New List(Of String)     ' list of all bands used in this contest
        Dim sqlBand As New List(Of String)

        If dlgContest.ShowDialog = DialogResult.OK Then
            contestID = dlgContest.DataGridView1.SelectedRows(0).Cells("id").Value
            dlgEntrant.Tag = contestID      ' pass contestID to dialog
            Using connect As New SqliteConnection(CheckerDB)
                connect.Open()
                connect.CreateFunction("BASECALL", Function(input As String) basecall(input))       ' define a function to remove suffix from call
                connect.CreateFunction("FREQUENCY", Function(band As String) frequency(band))
                sqlContest = connect.CreateCommand
                sqlEntrant = connect.CreateCommand
                sqlQSO = connect.CreateCommand
                sqlContest.CommandText = $"Select * FROM Contests WHERE contestID={contestID}"
                sqldrContest = sqlContest.ExecuteReader()
                sqldrContest.Read()
                Using report As New StreamWriter($"{sqldrContest("name")} Provisional Results.html")    ' open report file
                    report.WriteLine("<!DOCTYPE html>
<html>
<style>
 .info table, .info th, .info td {
    border: 1px solid black;
}
 .info {
    border-collapse:collapse
}
 .info td {
    padding: 3px;
}
 .section table, .section th, .section td {
    border: 1px solid black;
}
 .section {
    border-collapse:collapse
}
 .section td {
    padding: 3px;
}
 .section {font-family: Arial, Helvetica, sans-serif; font-size: small;}
 .section td:nth-child(1) {width: 120px;}
 .section td:nth-child(2) {width: 250px;}
 .section td:nth-child(4),.section td:nth-child(5), .section td:nth-child(6), .section td:nth-child(7), .section td:nth-child(8), .section td:nth-child(9), .section td:nth-child(10), .section td:nth-child(11), .section td:nth-child(12),td:nth-child(13),td:nth-child(14),td:nth-child(15),td:nth-child(16) {
    text-align:right; width: 40px;
}
 .aligncenter {
     display:block;
     margin-left:auto;
     margin-right:auto 
}
 .zebra table, .zebra th, .zebra td {
    border: 1px solid lightblue;
}
 .zebra {
    border-collapse:collapse
}
 .zebra tr:nth-child(even) {
    background-color:#efefef;rightdouble
}
 .zebra tr:nth-child(odd) {
    background-color:#e1e1e1;
}
 .zebra tr.deleted td{
    color:red;
}
 th {
    background-color:#e1e1e1
}
 .center {
    text-align: center;
}
 .right{
    text-align: Right;
}
 .left{
    text-align: Left;
}
 .boxed {
     border: 1px solid green ;
}
 h1 {
    text-align: center
}
td.rightthick {border-width: thin medium thin thin;}
td.leftthick {border-width: thin thin thin medium;}
</style>")

                    report.WriteLine($"<div class='center'>{sqldrContest("name")} {sqldrContest("Start")}</div>")
                    report.WriteLine("<h1>PROVISIONAL RESULTS</h1>")

                    ' get a list of all bands used in this contest
                    sqlQSO.CommandText = $"SELECT band FROM `QSO` WHERE `contestID`={contestID} and (`flags` & {CInt(flagsEnum.NonPermittedBand)})=0 GROUP BY band"
                    sqlQSOdr = sqlQSO.ExecuteReader
                    bandList.Clear()
                    While sqlQSOdr.Read
                        bandList.Add(sqlQSOdr("band"))
                    End While
                    sqlQSOdr.Close()
                    bandList.Sort(Function(a, b) frequency(a).CompareTo(frequency(b)))    ' sort in frequency order

                    ' Construct sql query for multiple bands
                    sqlBand.Clear()
                    For Each band In bandList
                        sqlBand.Add($"SUM(CASE WHEN band='{band}' THEN score ELSE 0 END) as B{bandList.IndexOf(band)}")
                    Next
                    Dim bandssql = $"{String.Join(",", sqlBand.ToArray())},SUM(score) AS Total" ' sql to get band counts and total
                    For Each section In Sections
                        report.WriteLine($"<h2>Section {section.Key} - {section.Value}</h2>")
                        ' make the table header
                        report.WriteLine("<table class='section'><tr><th>Call</th><th>Name</th><th>Active Lctr</th><th>Valid QSOs</th><Avg Km /QSO</th>")
                        For Each band In bandList
                            report.Write($"<th>{band}</th>")
                        Next
                        report.WriteLine("<th>Total</th></tr>")
                        ' Now do a sub section
                        For Each subsection In SubSections
                            sqlEntrant.CommandText = $"SELECT station,name,count(*) as Valid,gridsquare,{bandssql} FROM Stations AS S JOIN QSO AS Q ON S.station=Q.sent_call WHERE S.contestID={contestID} AND Q.contestID={contestID} AND S.Section='{section.Key}' AND S.subsection='{subsection.Key}' AND (`flags` & {DisqualifyQSO})=0 GROUP BY station ORDER By Total DESC"
                            sqldrEntrant = sqlEntrant.ExecuteReader
                            If sqldrEntrant.HasRows Then
                                ' create score table
                                report.WriteLine($"<tr><td class='center' colspan={bandList.Count + 5}><b>{subsection.Key} - {subsection.Value}</b></td></tr>")
                                While sqldrEntrant.Read
                                    report.WriteLine($"<tr><td>{sqldrEntrant("station")}</td><td>{sqldrEntrant("name")}</td><td>{sqldrEntrant("gridsquare")}</td><td class='rightthick'>{sqldrEntrant("Valid")}</td>")
                                    Dim col As Integer = 1
                                    For Each band In bandList
                                        Dim fieldname As String = $"B{bandList.IndexOf(band)}"
                                        Dim cls = If(col = 4, " class='rightthick'", "")
                                        report.Write($"<td{cls}>{SuppressZero(sqldrEntrant(fieldname))}</td>")
                                        col += 1
                                    Next
                                    report.WriteLine($"<td class='right leftthick'>{sqldrEntrant("Total")}</td></tr>")
                                End While
                            End If
                            sqldrEntrant.Close()
                        Next
                        report.WriteLine("</table>")
                    Next

                    report.WriteLine("<h2>Longest Distance Verified Contacts</h2>")
                    report.WriteLine("(Contact may be non-scoring, e.g. not in the best 8-hours, but BOTH logs must be received)<br><br>")
                    ' sub query selects the longest verified distance for each band
                    ' query then selects all QSO with longest band/distance
                    sqlEntrant.CommandText = $"SELECT Q.band as band,Q.date AS date,Q.distance AS distance,Q.sent_call AS sent_call,Q.rcvd_call AS rcvd_call,Q.sent_grid as sent_grid,Q.rcvd_grid as rcvd_grid FROM QSO AS Q JOIN (SELECT *,MAX(distance) AS distance FROM QSO WHERE `contestID`={contestID} AND distance is not null AND (`flags` & {DisqualifyQSO})=0 GROUP BY band) AS X ON Q.contestID=X.contestID AND Q.distance=X.distance AND Q.band=X.band where Q.sent_call<Q.rcvd_call GROUP BY Q.band,Q.sent_call,Q.rcvd_call ORDER BY frequency(Q.band),Q.date"
                    report.WriteLine("<table class='info'>")
                    report.WriteLine($"<tr><th>Band</th><th>Date</th><th class='right'>Distance (km)</th><th>Between</th><th>Locators</th></tr>")
                    Dim arrow = "&#x219e;&#x21a0"     ' double ended side arrow
                    Dim LastBand As String = ""
                    sqldrEntrant = sqlEntrant.ExecuteReader()
                    While sqldrEntrant.Read
                        Dim band As String = If(sqldrEntrant("band") = LastBand, """", sqldrEntrant("band"))  ' suppress band if same as previous
                        report.WriteLine($"<tr><td class='right'>{band}</td><td>{sqldrEntrant("date")}</td><td class='right'>{sqldrEntrant("distance")}</td><td class='center'>{sqldrEntrant("sent_call")} {arrow} {sqldrEntrant("rcvd_call")}</td><td class='center'>{sqldrEntrant("sent_grid")} {arrow} {sqldrEntrant("rcvd_grid")}</td></tr>")
                        LastBand = sqldrEntrant("band")
                    End While
                    report.WriteLine("</table>")
                    sqldrEntrant.Close()

                    report.WriteLine("<h2>Portable Locations</h2>")

                    report.WriteLine("<h2>Active Stations per Call Area</h2>")

                    report.WriteLine("<h2>Call Area to Call Area Contacts</h2>")

                    report.WriteLine("<h2>Multi-Op Portable Stations Operators</h2>")

                    report.WriteLine("<h2>Comments</h2>")
                    sqlEntrant.CommandText = $"SELECT * FROM Stations WHERE contestID={contestID} ORDER BY station"
                    sqldrEntrant = sqlEntrant.ExecuteReader
                    While sqldrEntrant.Read
                        If Not IsDBNull(sqldrEntrant("soapbox")) AndAlso sqldrEntrant("soapbox").length > 0 Then
                            report.WriteLine($"<h3>From the log of: {sqldrEntrant("station"),10} {sqldrEntrant("name")}</h3>")
                            report.WriteLine($"{sqldrEntrant("soapbox")}<br>")
                        End If
                    End While
                    sqldrEntrant.Close()
                    report.WriteLine($"End of Report. Created: {Now.ToUniversalTime} UTC by FD Log Checker - Marc Hillman (VK3OHM)<br>")
                    report.WriteLine("</html>")
                    TextBox2.Text = $"Report produced in file {CType(report.BaseStream, FileStream).Name}{vbCrLf}"
                End Using
            End Using
        End If
    End Sub
    Private Shared Function SuppressZero(value) As String
        ' Cause 0 to be printed as blank
        If IsDBNull(value) OrElse value = 0 Then
            Return ""
        Else
            Return value
        End If
    End Function

    Private Sub DeltaTimeAnalysisToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeltaTimeAnalysisToolStripMenuItem.Click
        Dim contestID As Integer, station As String, section As String
        Dim sqlStations As SqliteCommand, sqlStationsdr As SqliteDataReader
        Dim sqlQSO As SqliteCommand
        Dim sqlupd As SqliteCommand
        Dim Iteration As Integer = 0

        If dlgContest.ShowDialog = DialogResult.OK Then
            contestID = dlgContest.DataGridView1.SelectedRows(0).Cells("id").Value
            dlgEntrant.Tag = contestID      ' pass contestID to dialog
            If dlgEntrant.ShowDialog() = DialogResult.OK Then
                station = dlgEntrant.DataGridView1.SelectedRows(0).Cells("station").Value
                section = dlgEntrant.DataGridView1.SelectedRows(0).Cells("section").Value
                Using connect As New SqliteConnection(CheckerDB)
                    connect.Open()
                    Dim tr As SqliteTransaction = connect.BeginTransaction
                    Using tr
                        connect.CreateFunction("BASECALL", Function(input As String) basecall(input))       ' define a function to remove suffix from call
                        connect.CreateFunction("DELTAT", Function(a As String, b As String) DeltaT(a, b))
                        sqlQSO = connect.CreateCommand
                        sqlupd = connect.CreateCommand
                        sqlStations = connect.CreateCommand
                        sqlQSO.CommandText = $"UPDATE `Stations` SET DT=NULL WHERE contestID={contestID}"     ' initialise all DT
                        sqlQSO.ExecuteNonQuery()
                        sqlQSO.CommandText = $"UPDATE `Stations` SET DT=0 WHERE contestID={contestID} AND station='{station}'"     ' initialise all DT for station 0 to 0
                        sqlQSO.ExecuteNonQuery()
                        Do
                            sqlStations.CommandText = "SELECT *,
A.date AS Adate, B.date AS Bdate, S.station AS Sstation, T.DT AS DT, DELTAT(A.date,B.date) AS DELTA 
FROM     Stations AS S
JOIN     QSO      AS A
JOIN     QSO      AS B
JOIN     Stations AS T
ON       S.contestID=A.contestID AND S.contestID=T.contestID AND S.station=basecall(A.sent_call) AND A.match=B.id AND basecall(B.sent_call)=T.station
WHERE    S.DT IS NULL
AND		 T.DT IS NOT NULL
GROUP BY basecall(A.sent_call),
         basecall(A.rcvd_call)"
                            sqlStationsdr = sqlStations.ExecuteReader
                            updated = 0
                            Dim count = 0
                            While sqlStationsdr.Read
                                count += 1
                                sqlupd.CommandText = $"UPDATE `Stations` SET DT={sqlStationsdr("DT")}+{sqlStationsdr("DELTA")} WHERE contestID={contestID} AND station='{sqlStationsdr("Sstation")}'"
                                updated += sqlupd.ExecuteNonQuery
                            End While
                            Iteration += 1
                            TextBox2.AppendText($"{count} stations to be updated, {updated} DT updates on iteration {Iteration}{vbCrLf}")
                            sqlStationsdr.Close()
                            Application.DoEvents()
                        Loop Until updated = 0 Or Iteration >= 8
                        tr.Commit()
                    End Using
                End Using
            End If
        End If
    End Sub
    Function DeltaT(a As String, b As String) As Integer
        ' Calculate difference between 2 timestamps, in minutes
        Dim Atime As Date = Convert.ToDateTime(a)
        Dim Btime As Date = Convert.ToDateTime(b)
        Return Btime.Subtract(Atime).TotalMinutes
    End Function
    Private Sub SubmittedLogsToolStripMenuItem_Click_1(sender As Object, e As EventArgs) Handles SubmittedLogsToolStripMenuItem.Click
        ' produce a submitted logs report
        Dim contestID As Integer
        Dim sqlContest As SqliteCommand, sqldrContest As SqliteDataReader
        Dim sqlEntrant As SqliteCommand
        Dim sqlQSO As SqliteCommand, sqlQSOdr As SqliteDataReader
        Dim QSOcounts As New List(Of QSOcount)
        Dim bandList As New List(Of String)     ' list of all bands used in this contest
        Dim sqlBand As New List(Of String)

        If dlgContest.ShowDialog = DialogResult.OK Then
            contestID = dlgContest.DataGridView1.SelectedRows(0).Cells("id").Value
            dlgEntrant.Tag = contestID      ' pass contestID to dialog
            Using connect As New SqliteConnection(CheckerDB)
                connect.Open()
                connect.CreateFunction("BASECALL", Function(input As String) basecall(input))       ' define a function to remove suffix from call
                connect.CreateFunction("FREQUENCY", Function(band As String) frequency(band))
                sqlContest = connect.CreateCommand
                sqlEntrant = connect.CreateCommand
                sqlQSO = connect.CreateCommand
                sqlContest.CommandText = $"Select * FROM Contests WHERE contestID={contestID}"
                sqldrContest = sqlContest.ExecuteReader()
                sqldrContest.Read()
                Using report As New StreamWriter($"{sqldrContest("name")} Submitted Logs.html")    ' open report file
                    report.WriteLine("<!DOCTYPE html>
<html>
<style>
 .info table, .info th, .info td {
    border: 1px solid black;
}
 .info {
    border-collapse:collapse
}
 .info td {
    padding: 3px;
}
 .section table, .section th, .section td {
    border: 1px solid black;
}
 .section {
    border-collapse:collapse
}
 .section td {
    padding: 3px;
}
 .section {font-family: Arial, Helvetica, sans-serif; font-size: small;}
 .section td:nth-child(1) {width: 120px;}
 .section td:nth-child(2) {width: 250px;}
 .section td:nth-child(4),.section td:nth-child(5), .section td:nth-child(6), .section td:nth-child(7), .section td:nth-child(8), .section td:nth-child(9), .section td:nth-child(10), .section td:nth-child(11), .section td:nth-child(12),td:nth-child(13),td:nth-child(14),td:nth-child(15),td:nth-child(16) {
    text-align:right; width: 40px;
}
 .aligncenter {
     display:block;
     margin-left:auto;
     margin-right:auto 
}
 .zebra table, .zebra th, .zebra td {
    border: 1px solid lightblue;
}
 .zebra {
    border-collapse:collapse
}
 .zebra tr:nth-child(even) {
    background-color:#efefef;
}
 .zebra tr:nth-child(odd) {
    background-color:#e1e1e1;
}
 .zebra tr.deleted td{
    color:red;
}
 th {
    background-color:#e1e1e1
}
 .center {
    text-align: center;
}
 .right{
    text-align: Right;
}
 .left{
    text-align: Left;
}
 .boxed {
     border: 1px solid green ;
}
 h1 {
    text-align: center
}
</style>")

                    report.WriteLine($"<div class='center'>{sqldrContest("name")} {sqldrContest("Start")}</div>")
                    sqldrContest.Close()
                    report.WriteLine("<h2>List of Submitted Logs</h2>")
                    report.WriteLine("<table class='info'>")
                    report.WriteLine("<tr><th>#</th><th>Callsign</th><th>Category</th><th>email</th></tr>")
                    Dim index As Integer = 1
                    sqlContest.CommandText = $"SELECT * FROM `Stations` WHERE `contestID`={contestID} ORDER BY `station`,`section`"
                    sqldrContest = sqlContest.ExecuteReader
                    While sqldrContest.Read
                        report.WriteLine($"<tr><td class='right'>{index}</td><td>{sqldrContest("station")}</td><td class='center'>{sqldrContest("section")}</td><td>{sqldrContest("email")}</td></tr>")
                        index += 1
                    End While
                    sqldrContest.Close()
                    report.WriteLine("</table>")

                    report.WriteLine("<h2>Contest Statistics</h2>")
                    report.WriteLine("<table class='info'>")
                    sqlQSO.CommandText = $"SELECT COUNT(*) AS TotalQSO,
                                                   SUM(
                                                   CASE
                                                          WHEN (
                                                                        `flags` & {CInt(flagsEnum.NotInLog)})<>0 THEN 1
                                                          ELSE 0
                                                   END) AS NotInLog,
                                                   SUM(
                                                   CASE
                                                          WHEN (
                                                                        `flags` & {CInt(flagsEnum.LoggedIncorrectCall)})<>0 THEN 1
                                                          ELSE 0
                                                   END) AS BadCall,
                                                   SUM(
                                                   CASE
                                                          WHEN (
                                                                        `flags` & {CInt(flagsEnum.LoggedIncorrectExchange)})<>0 THEN 1
                                                          ELSE 0
                                                   END) AS BadExch,
                                                   SUM(
                                                   CASE
                                                          WHEN (
                                                                        `flags` & {CInt(flagsEnum.LoggedIncorrectLocator)})<>0 THEN 1
                                                          ELSE 0
                                                   END) AS BadGrid
                                            FROM   `QSO`
                                            WHERE  `contestID`={contestID}"
                    sqlQSOdr = sqlQSO.ExecuteReader
                    sqlQSOdr.Read()
                    report.WriteLine($"<tr><td>{sqlQSOdr("TotalQSO")}</td><td>Total number of contacts logged</td></tr>")
                    report.WriteLine($"<tr><td>{sqlQSOdr("NotInLog")}</td><td>Not in log</td></tr>")
                    report.WriteLine($"<tr><td>{sqlQSOdr("BadCall")}</td><td>Call copied incorrectly</td></tr>")
                    report.WriteLine($"<tr><td>{sqlQSOdr("BadExch")}</td><td>Exchange copied incorrectly</td></tr>")
                    report.WriteLine($"<tr><td>{sqlQSOdr("BadGrid")}</td><td>Grid square copied incorrectly</td></tr>")
                    report.WriteLine("</table>")
                    sqlQSOdr.Close()

                    report.WriteLine($"<br>End of Report. Created: {Now.ToUniversalTime} UTC by FD Log Checker - Marc Hillman (VK3OHM)<br>")
                    report.WriteLine("</html>")
                    TextBox2.Text = $"Report produced in file {CType(report.BaseStream, FileStream).Name}{vbCrLf}"
                End Using
            End Using
        End If
    End Sub
End Class
