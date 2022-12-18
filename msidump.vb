Imports System.Xml
Imports System.IO
Imports System.Runtime.InteropServices





Module msidump

    'Declare Function MsiGetSummaryInformation Lib "MSI" (ByRef hDatabase As UInt32, ByVal szDatabasePath As String, ByVal uiUpdateCount As Long, ByRef phSummaryInfo As UInt32) As Long
    'Declare Function MsiSummaryInfoGetProperty Lib "MSI" (ByRef hDatabase As UInt32, ByVal szDatabasePath As String, ByVal uiUpdateCount As Long, ByRef phSummaryInfo As UInt32) As Long
    '        [DllImport("msi.dll", EntryPoint = "MsiSummaryInfoGetPropertyW", CharSet = CharSet.Unicode, ExactSpelling = true)]
    '    internal static extern int MsiSummaryInfoGetProperty(uint summaryInfo, int property, out uint dataType, out int integerValue, ref FILETIME fileTimeValue, StringBuilder stringValueBuf, ref int stringValueBufSize);

            


    Sub Main(ByVal args As String())
        If args.Length > 0 Then
            If args(0).EndsWith("?") Or args(0).EndsWith("help") Or args(0).EndsWith("/?") Or args(0).EndsWith("-?") Then
                Usage()
            Else
                Dim o = New Opt(args)
                If o.IsDiffMode Then
                    If o.IsValidArgs Then
                        If File.Exists(o.RefMsiXmlPath) Then
                            System.Console.Out.WriteLine(o.RefMsiXmlName & ": " & o.RefMsiXmlPath)
                            Dim refX = New XMsi(o.RefMsiXmlName, o.RefMsiXmlPath)
                            Dim xDif = New XTDiff(refX)
                            For Each mx In o.GetMsiXmlNames()
                                System.Console.Out.WriteLine(mx & ": " & o.GetMsiXmlPath(mx))
                                xDif.AddMsiXml(mx, o.GetMsiXmlPath(mx))
                            Next
                            xDif.WriteResult(o)
                        Else

                            System.Console.Out.WriteLine("File Not Found: " & o.RefMsiXmlPath)
                            Return
                        End If
                    Else
                        System.Console.Out.WriteLine("Invalid parameters: " & Join(args, " "))
                        Usage()
                        Return
                    End If
                Else
                    Try
                        If File.Exists(o.MsiFile) Then
                            Dim oIn = CreateObject("WindowsInstaller.Installer")
                            Dim mDB = New MSIDB(oIn, o)
                            If Not Directory.Exists(o.OutDirectory) Then Directory.CreateDirectory(o.OutDirectory)

                            If o.HasOptions And o.QueryLevel > -1 Then '/l is specified
                                mDB.DumpQuery()
                            End If
                            mDB.DumpToXml()
                        Else
                            LogErr("File not found: " & o.MsiFile)
                            Usage()
                        End If
                    Catch e As Exception
                        LogErr("ErrorOccurred: " & e.Message)
                    End Try
                End If
            End If
        Else
            Usage()
        End If
    End Sub

    Public Sub Usage()
        Log("Usage:")
        Log("msidump.exe [/f] <msiPath> [/t table1;table2...] [/l table;store] [/x xslFilePath] [/o outputDirectory]")
        Log("Options:")
        Log("   <msiPath>                  MSI file to dump. Required")
        Log("   [/l table|store]           List for table or storage names. Optional")
        Log("   [/t table1;table2...]      MSI tables to dump. Optional")
        Log("   [/a store1;store2...]      Apply specified embedded MSTs. Optional")
        Log("   [/e mstfile1;mstfile2...]  Apply specified external MST file/s. Optional")
        Log("   [/x xslFilePath]           XML Stylesheet file path. Optional")
        Log("   [/n outputFileName]        Output filename. Optional")
        Log("   [/o outputDirectory]       Output directory. Optional")
        Log("   [/d n1=path1 n2=path2...]  DiffMode. n1, n2 are any unique names to the specified XML msidump file.")
        Log("   [/b]                       Suppress summary information stream dump.")
        Log("Example:")
        Log("   $>msidump c:\tmp\mps.msi")
        Log("      -dumps all tables and transform views to the same directory as the msi file.")
        Log("   $>msidump /d RTM=mps_rtm.msi.mxl B05=mps_b05.msi.mxl")
        Log("      -dumps all the changes made on the tables from RTM to B05.")

    End Sub
    Public Sub Log(ByVal txt As String)
        Console.Out.WriteLine(txt)
    End Sub
    Public Sub LogErr(ByVal txt As String)
        Console.Error.WriteLine(txt)
    End Sub
End Module



'classes
Public Class MSIDB
    Const XFMV = &H100
    Private bPatch As Boolean
    Private opt As Opt
    Private xForms As String()
    Private oInst As Object
    Private oDB As Object
    Private dicTable As Dictionary(Of String, MSITable)
    Private curXForm As String
    Private msiPath As String
    Private sumInfo As Object
    Private dbType As Integer = 0
    Private SISPROPS As Dictionary(Of Integer, String)
    Public _MSIPROPS As Dictionary(Of Integer, String) = New Dictionary(Of Integer, String) From {
    {2, "Title"}, {3, "Subject"}, {4, "Author"}, {5, "Keywords"}, {6, "Comments"},
    {7, "Templates"}, {8, "LastAuthor"}, {9, "PackageCode"}, {11, "LastPrinted"},
    {12, "DateCreated"}, {13, "DateSaved"}, {14, "Schema"}, {15, "ImageType"},
    {18, "CreatedWith"}, {19, "Security"}, {1, "CodePage"}
    }

    Public _MSPPROPS As Dictionary(Of Integer, String) = New Dictionary(Of Integer, String) From {
    {2, "Title"}, {3, "Subject"}, {4, "Author"}, {5, "Keywords"}, {6, "Comments"},
    {7, "Target"}, {8, "MSTNames"}, {9, "PatchCode"},
    {12, "DateCreated"}, {13, "DateSaved"}, {15, "Schema"},
    {18, "CreatedWith"}, {19, "Security"}, {1, "CodePage"}
    }
    Public Sub New(ByRef oInst As Object, ByRef opt As Opt)
        Me.oInst = oInst
        Me.dicTable = New Dictionary(Of String, MSITable)()
        Me.opt = opt
        InitDB()
    End Sub
    Private Sub InitDB()
        bPatch = Me.opt.MsiFile.EndsWith(".msp", StringComparison.CurrentCultureIgnoreCase)
        If bPatch Then
            dbType = 32
            SISPROPS = _MSPPROPS
        Else
            SISPROPS = _MSIPROPS
        End If

        Me.oDB = Me.oInst.OpenDatabase(Me.opt.MsiFile, dbType)
        'sumInfo = Me.oDB.SummaryInformation
    End Sub
    Public Sub ApplyTransform(ByVal xForm As String, Optional ByVal embXForm As Boolean = False)
        'If Not curXForm Is Nothing Then FreeTransformView()
        curXForm = xForm
        Dim x = xForm
        If embXForm Then x = ":" & x
        oDB.ApplyTransform(x, XFMV)
        sumInfo = Me.oDB.SummaryInformation
    End Sub
    'Public Sub FreeTransformView()
    '    If Not curXForm Is Nothing Then
    '        Dim oVw = oDB.OpenView("ALTER TABLE _TransformView FREE")
    '        oVw.Execute()
    '        oVw.Close()
    '        oVw = Nothing
    '        curXForm = Nothing
    '        System.Runtime.InteropServices.Marshal.ReleaseComObject(oVw)
    '    End If
    'End Sub
    Public ReadOnly Property AppliedTransform() As String
        Get
            Return Me.curXForm
        End Get
    End Property
    Public ReadOnly Property IsTransformed() As Boolean
        Get
            Return Not Me.curXForm Is Nothing
        End Get
    End Property
    Public ReadOnly Property Transforms() As String()
        Get
            If Me.xForms Is Nothing Then
                Me.xForms = ColumnToArray("_Storages", "Name")
            End If
            Return Me.xForms
        End Get
    End Property
    Public ReadOnly Property HasTransforms() As Boolean
        Get
            Return Me.Transforms.Length > 0
        End Get
    End Property
    Public ReadOnly Property Tables() As String()
        Get
            Return ColumnToArray("_Tables", 1)
        End Get
    End Property
    Public Function GetTable(ByVal tableName As String) As MSITable
        If Me.dicTable.ContainsKey(tableName) Then Return Me.dicTable.Item(tableName)
        Dim tab = New MSITable(Me.oDB, tableName)
        Me.dicTable.Add(tableName, tab)
        Return tab
    End Function
    Private Function ColumnToArray(ByVal tableName As String, ByVal columnIndex As Integer) As String()
        Return MSITable.ColToArray(oDB.OpenView("SELECT * FROM " & tableName), columnIndex)
    End Function
    Private Function ColumnToArray(ByVal tableName As String, ByVal columnName As String) As String()
        Return MSITable.ColToArray(oDB.OpenView("SELECT " & columnName & " FROM " & tableName), 1)
    End Function
    Public Function ObjectExist(ByVal tableName As String, ByVal columnName As String, ByVal valString As String) As Boolean
        Return MSITable.ColToArray(oDB.OpenView("SELECT " & columnName & " FROM " & tableName & " WHERE " & columnName & "='" & valString & "'"), 1).Length > 0
    End Function
    Public Sub WriteXML(ByRef xml As XmlTextWriter)
        Dim ts = Tables
        xml.WriteAttributeString(Vars.TABLE, ts.Length)

        For Each t As String In ts
            GetTable(t).WriteXML(xml)
        Next

    End Sub
    Private Sub WriteSIS(ByRef xml As XmlTextWriter)
        xml.WriteStartElement(Vars.TABLE)
        xml.WriteAttributeString(Vars.NAME, Vars._SISINFO)
        'metadata
        xml.WriteStartElement(Vars.HEADER)
        xml.WriteAttributeString(Vars.COUNT, 2)
        xml.WriteStartElement(Vars.FIELD)
        xml.WriteAttributeString(Vars.NAME, Vars._PROPERTY)
        xml.WriteAttributeString(Vars.TYPE, Vars._STRING)
        xml.WriteAttributeString(Vars.KEY, "1")
        xml.WriteEndElement() 'field
        xml.WriteStartElement(Vars.FIELD)
        xml.WriteAttributeString(Vars.NAME, Vars._VALUE)
        xml.WriteAttributeString(Vars.TYPE, Vars._STRING)
        xml.WriteEndElement() 'field
        xml.WriteEndElement() 'header

        'data

        xml.WriteStartElement(Vars.DATA)
        xml.WriteAttributeString(Vars.COUNT, SISPROPS.Count)
        For Each i As Integer In SISPROPS.Keys
            xml.WriteStartElement(Vars.ROW)
            xml.WriteAttributeString(Vars.KEY, SISPROPS(i))
            xml.WriteStartElement(Vars._PROPERTY)
            xml.WriteString(SISPROPS(i))
            xml.WriteEndElement() 'prop
            xml.WriteStartElement(Vars._VALUE)
            xml.WriteString(sumInfo.Property(i))
            xml.WriteEndElement() 'value
            xml.WriteEndElement() 'row
        Next

        xml.WriteEndElement() 'data
        xml.WriteEndElement() 'table
    End Sub
    Public Sub WriteXML(ByRef xml As XmlTextWriter, ByVal tableName As String, Optional ByVal attName As String = Nothing, Optional ByVal attVal As String = Nothing)
        GetTable(tableName).WriteXML(xml, attName, attVal)
    End Sub
    Public Sub DumpTransforms(ByRef xFs As String(), ByRef xml As XmlTextWriter)
        For Each xF As String In xFs
            If Me.ObjectExist("_Storages", "Name", xF) Then
                Me.ApplyTransform(xF, True)
                DumpToXml(xml)
                'mDB.FreeTransformView()
            Else 'for external transform files
                If File.Exists(xF) Then
                    Me.ApplyTransform(xF, False)
                    DumpToXml(xml)
                    'mDB.FreeTransformView()
                Else
                    LogErr("InvalidParameter: Storage " & xF & " does not exist!")
                End If
            End If
            If xFs.Length > 1 Then
                Me.oDB = Nothing
                Me.dicTable.Remove("_TransformView")
                InitDB()
            End If

        Next
    End Sub
    Public Sub DumpTransforms(ByRef xFs As String())
        If xFs.Length > 0 Then
            Dim xml = OpenDumpXMLWriter(Vars.MST)
            DumpTransforms(xFs, xml)
            CloseDumpXMLWriter(xml)
        End If
    End Sub
    Public Sub DumpToXml()
        Dim xml = OpenDumpXMLWriter()
        If (opt.DumpLevel And 2) Or opt.DumpLevel = 0 Then DumpToXml(xml)
        If (opt.DumpLevel And 1) Or opt.DumpLevel = 0 Then
            Dim sL = opt.StoreList
            If sL.Length = 0 Then
                sL = Me.Transforms
            End If
            DumpTransforms(sL, xml)
        End If
        If opt.DumpSIS Then WriteSIS(xml)
        CloseDumpXMLWriter(xml)
    End Sub

    Public Sub DumpToXML(ByRef xml As XmlTextWriter)
        If Me.IsTransformed Then
            'xml.WriteAttributeString("appliedMST", mDB.AppliedTransform)
            Me.WriteXML(xml, "_TransformView", Vars.MST, Me.AppliedTransform)
        Else
            Dim tL = Me.opt.TableList
            xml.WriteAttributeString(Vars.MST, String.Join(";", Me.Transforms))
            If tL.Length = 0 Then
                Me.WriteXML(xml)
            Else
                For Each t In tL
                    If Me.ObjectExist("_Tables", "Name", t) Then
                        Me.WriteXML(xml, t)
                    Else
                        LogErr("InvalidParameter: Table " & t & " does not exist!")
                    End If
                Next
            End If
        End If
    End Sub
    Public Function OpenDumpXMLWriter(Optional ByVal suffix As String = Nothing) As XmlTextWriter
        Dim xml = New XmlTextWriter(Me.opt.OutFilePath(suffix), Text.Encoding.UTF8)
        xml.WriteStartDocument()
        'xml.WriteProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
        If Not Me.opt.XslFile Is Nothing Then
            Dim PItext = "type='text/xsl' href='" + Me.opt.XslFile + "'"
            xml.WriteProcessingInstruction("xml-stylesheet", PItext)
        End If

        xml.WriteStartElement(Vars.ROOT)
        xml.WriteAttributeString(Vars.PATH, Me.opt.MsiFile)
        Return xml
    End Function
    Public Sub CloseDumpXMLWriter(ByRef xml As XmlTextWriter)
        xml.WriteEndElement() 'msidump
        xml.WriteEndDocument() 'root
        xml.Close()
    End Sub
    Public Sub DumpQuery()
        If Me.opt.QueryLevel And 1 Then 'table names only
            Print("#Tables: " & Me.Tables.Length)
            Print(String.Join(vbCrLf, Me.Tables))
        End If
        If Me.opt.QueryLevel And 2 Then 'storage names only
            Print("#Stores: " & Me.Transforms.Length)
            Print(String.Join(vbCrLf, Me.Transforms))
        End If
    End Sub
    Public Sub Print(ByVal txt As String)
        Console.Out.WriteLine(txt)
    End Sub
End Class

Public Class MSITransformEntry
    Private mLang, mKey, mCurrent, mData As String
    Public Sub New(ByVal sLang As String, ByVal sKey As String, ByVal sCurrent As String, ByVal sData As String)
        Me.mKey = sKey
        Me.mLang = sLang
        Me.mCurrent = sCurrent
        Me.mData = sData
    End Sub
    Public ReadOnly Property Key As String
        Get
            Return mKey
        End Get
    End Property
    Public ReadOnly Property Data As String
        Get
            Return mData
        End Get
    End Property
    Public ReadOnly Property Current As String
        Get
            Return mCurrent
        End Get
    End Property
    Public ReadOnly Property Language As String
        Get
            Return mLang
        End Get
    End Property

End Class

Public Class MSITransformEntryGroup
    Private mEntries As Dictionary(Of String, MSITransformEntry)
    Public Sub New()
        mEntries = New Dictionary(Of String, MSITransformEntry)()
    End Sub
    Public Sub AddEntry(ByRef mEntry As MSITransformEntry)
        mEntries.Add(mEntry.Key, mEntry)
    End Sub
    Public ReadOnly Property Count As Integer
        Get
            Return mEntries.Count
        End Get
    End Property
    Public Sub WriteXML(ByRef xml As XmlTextWriter)

    End Sub
End Class

Public Class MSITable

    Private tableName As String
    Private columnNames As String()
    Private columnTypes As String()
    Private keyColumns As List(Of String)
    Private tableRows As Dictionary(Of String, MSIRow)
    Public Sub New(ByRef oDB As Object, ByVal tableName As String)
        Me.tableName = tableName
        Dim oVw = oDB.OpenView("SELECT * FROM " & tableName)
        oVw.Execute()
        Me.columnNames = RecToArray(oVw.ColumnInfo(0))
        Me.keyColumns = RecToList(oDB.PrimaryKeys(tableName))
        SetTypes(oVw)
        Me.tableRows = New Dictionary(Of String, MSIRow)()
        Do
            Dim oRs = oVw.Fetch()
            If oRs Is Nothing Then Exit Do
            Dim row = RecToMsiRow(oRs)
            Me.tableRows.Add(row.Key, row)
        Loop
        'If tableName.Equals("_TransformView") Then
        '    oVw = oDB.OpenView("ALTER TABLE _TransformView FREE")
        '    oVw.Execute()
        'End If

        oVw.Close()
    End Sub
    Public ReadOnly Property Name() As String
        Get
            Return Me.tableName
        End Get
    End Property
    Public ReadOnly Property KeyFields() As String()
        Get
            Return Me.keyColumns.ToArray()
        End Get
    End Property
    Public ReadOnly Property FieldNames() As String()
        Get
            Return Me.columnNames
        End Get
    End Property
    Public ReadOnly Property FieldTypes() As String()
        Get
            Return Me.columnTypes
        End Get
    End Property
    Public ReadOnly Property FieldIndex(ByVal columnName As String) As Integer
        Get
            Return Array.IndexOf(Me.columnNames, columnName) + 1
        End Get
    End Property
    Public ReadOnly Property FieldType(ByVal columnName As String) As String
        Get
            Dim i = FieldIndex(columnName) - 1
            Dim t = Nothing
            If i > -1 Then t = Me.columnTypes(i)
            Return t
        End Get
    End Property
    Public Sub AddRow(ByRef msiRow As MSIRow)
        Me.tableRows.Add(msiRow.Key, msiRow)
    End Sub
    Public Function GetRow(ByVal key As String) As MSIRow
        Return Me.tableRows.Item(key)
    End Function
    Public ReadOnly Property GetRows() As ICollection(Of MSIRow)
        Get
            Return Me.tableRows.Values
        End Get
    End Property

    Public ReadOnly Property RowCount() As String
        Get
            Return Me.tableRows.Count
        End Get
    End Property
    Public ReadOnly Property GetRowKeys() As ICollection(Of String)
        Get
            Return Me.tableRows.Keys
        End Get
    End Property
    Private Sub SetTypes(ByRef oVw As Object) ' WindowsInstaller.View)
        's?	String, variable length (?=1-255)
        's0	String, variable length
        'i2	Short integer
        'i4	Long integer
        'v0	Binary Stream
        'g?	Temporary string (?=0-255)
        'j?	Temporary integer (?=0,1,2,4)
        Dim oRec = oVw.ColumnInfo(1)
        Dim oList = New List(Of String)()
        For i As Integer = 1 To oRec.FieldCount
            Dim s = oRec.StringData(i)
            If s.StartsWith("s", StringComparison.CurrentCultureIgnoreCase) Or s.StartsWith("l", StringComparison.CurrentCultureIgnoreCase) Or s.StartsWith("g", StringComparison.CurrentCultureIgnoreCase) Then
                s = Vars._STRING
            ElseIf s.StartsWith("i", StringComparison.CurrentCultureIgnoreCase) Or s.StartsWith("j", StringComparison.CurrentCultureIgnoreCase) Then
                s = Vars._INTEGER
            Else
                s = Vars._BINARY
            End If
            oList.Add(s)
        Next
        Me.columnTypes = oList.ToArray()
    End Sub
    Private Function RecToMsiRow(ByRef oRec As Object) As MSIRow
        Dim sKF = New List(Of String)()
        Dim sVl = New List(Of String)()
        Dim cTy, cNm, v As String
        For i As Integer = 1 To oRec.FieldCount
            cNm = Me.columnNames(i - 1)
            cTy = FieldType(cNm)
            If Vars._STRING.Equals(cTy) Then
                v = oRec.StringData(i)
            ElseIf Vars._INTEGER.Equals(cTy) Then
                v = oRec.IntegerData(i).ToString()
            Else
                v = Vars._BINARY
            End If
            If Me.keyColumns.Contains(cNm) Then sKF.Add(v)
            sVl.Add(v)
        Next
        Return New MSIRow(String.Join("_", sKF.ToArray()), sVl.ToArray())
    End Function
    Shared Function RecToArray(ByRef oRec) As String()
        Return (RecToList(oRec).ToArray())
    End Function
    Shared Function RecToList(ByRef oRec) As List(Of String)
        Dim oList = New List(Of String)()
        For i As Integer = 1 To oRec.FieldCount
            oList.Add(oRec.StringData(i))
        Next
        Return oList
    End Function
    Shared Function ColToArray(ByRef oView As Object, ByVal colIndex As Integer) As String()
        Dim oList = New List(Of String)()
        oView.Execute()
        Do
            Dim oRs = oView.Fetch()
            If oRs Is Nothing Then Exit Do
            oList.Add(oRs.StringData(colIndex))
        Loop
        oView.Close()
        Return oList.ToArray()
    End Function
    Public Function HeaderString(Optional ByVal delim As String = vbTab) As String
        Dim oList = New List(Of String)()
        oList.Add("RowKey " & Vars._STRING)
        For i As Integer = 0 To Me.columnNames.Length - 1
            Dim h = Me.columnNames(i)
            If Me.keyColumns.Contains(h) Then h = h & "*"
            oList.Add(h & " " & Me.columnTypes(i))
        Next
        Return String.Join(delim, oList.ToArray())
    End Function
    Public Sub Dump(Optional ByVal delim As String = vbTab)
        Log("--------------------------------------------------------")
        Log(HeaderString(delim))
        Log("--------------------------------------------------------")
        For Each r As MSIRow In GetRows
            Log(r.ToDelimString(delim))
        Next
        Log("--------------------------------------------------------")
    End Sub
    Sub Log(ByVal txt As String)
        Console.Out.WriteLine(txt)
    End Sub

    Public Sub WriteXML(ByRef xml As XmlTextWriter, Optional ByVal attName As String = Nothing, Optional ByVal attVal As String = Nothing)
        xml.WriteStartElement(Vars.TABLE)
        xml.WriteAttributeString(Vars.NAME, tableName)
        If Not attName Is Nothing Then
            xml.WriteAttributeString(attName, attVal)
        End If
        'metadata
        xml.WriteStartElement(Vars.HEADER)
        xml.WriteAttributeString(Vars.COUNT, Me.columnNames.Length)
        For i As Integer = 0 To Me.columnNames.Length - 1
            xml.WriteStartElement(Vars.FIELD)
            Dim h = Me.columnNames(i)
            xml.WriteAttributeString(Vars.NAME, h)
            xml.WriteAttributeString(Vars.TYPE, Me.columnTypes(i))
            If Me.keyColumns.Contains(h) Then xml.WriteAttributeString(Vars.KEY, "1")
            xml.WriteEndElement()
        Next
        xml.WriteEndElement()
        'data
        xml.WriteStartElement(Vars.DATA)
        Dim rows = GetRows
        xml.WriteAttributeString(Vars.COUNT, rows.Count)
        For Each r As MSIRow In rows
            xml.WriteStartElement(Vars.ROW)
            xml.WriteAttributeString(Vars.KEY, r.Key)
            Dim v = r.FieldValues()
            For i As Integer = 0 To Me.columnNames.Length - 1
                xml.WriteStartElement(Me.columnNames(i))
                xml.WriteString(v(i))
                xml.WriteEndElement()
            Next
            xml.WriteEndElement()
        Next
        xml.WriteEndElement()
        xml.WriteEndElement()
    End Sub
End Class

Public Class MSIRow
    Private pKey As String
    Private sFields As String()
    Public Sub New(ByVal sKey As String, Optional ByVal vals As String() = Nothing)
        Me.pKey = sKey
        If Not vals Is Nothing Then Me.sFields = vals
    End Sub
    Public ReadOnly Property Key() As String
        Get
            Return Me.pKey
        End Get
    End Property
    Public Property FieldValues() As String()
        Get
            Return Me.sFields
        End Get
        Set(ByVal value As String())
            Me.sFields = value
        End Set
    End Property
    Public ReadOnly Property FieldValue(ByVal index As Integer) As String
        Get
            If Not (Me.sFields Is Nothing) And (index > 0) And (index <= Me.FieldCount) Then Return Me.sFields(index)
            Return Nothing
        End Get
    End Property
    Public ReadOnly Property FieldCount() As Integer
        Get
            If Me.sFields Is Nothing Then Return 1
            Return Me.sFields.Length
        End Get
    End Property
    Public Function ToDelimString(Optional ByVal delim As String = vbTab) As String
        Return Me.Key & delim & String.Join(delim, Me.sFields)
    End Function
End Class

Public Class Opt
    Private msiPath As String
    Private tabList As List(Of String)
    Private stoList As Dictionary(Of String, Boolean)
    Private msiList As Dictionary(Of String, String)
    Private qryLvel As Integer = -1
    Private dmpLvel As Integer = 0
    Private bNoSis As Boolean = False
    Private ssFile As String
    Private outDir As String
    Private outFile As String = String.Empty
    Private bOption As Boolean
    Private refName As String
    Private bCmpMode As Boolean = False
    Private bVldArgs As Boolean = True

    Public Sub New(ByVal args As String())
        tabList = New List(Of String)()
        stoList = New Dictionary(Of String, Boolean)()
        msiList = New Dictionary(Of String, String)()
        If Not args(0).StartsWith("/") Then msiPath = args(0)
        ProcArgs(args)
    End Sub
    Private Sub ProcArgs(ByVal args As String())
        Dim v = String.Empty
        For Each s In args
            If s.StartsWith("/") Then
                v = s.ToLower()
                If v.Equals("/l") Then
                    qryLvel = 3 'default query level
                    bOption = True
                End If
                If v.Equals("/d") Then
                    bCmpMode = True
                End If
                If v.Equals("/b") Then
                    bNoSis = True
                End If
                If v.Equals("/t") Then
                    dmpLvel = dmpLvel Or 2
                    bOption = True
                End If
                If v.Equals("/a") Or v.Equals("/e") Then
                    dmpLvel = dmpLvel Or 1
                    bOption = True
                End If
            Else
                Select Case v
                    Case "/f" 'file
                        msiPath = s
                    Case "/t" 'tables to dump
                        AddTable(s.Split(";"c))
                    Case "/a" 'stores to apply
                        AddStore(s.Split(";"c), True)
                    Case "/e" 'external mst to apply
                        AddStore(s.Split(";"c), False)
                    Case "/l" 'list option
                        SetQueryLevel(s.Split(";"c))

                    Case "/x" 'list
                        ssFile = s
                        bOption = True
                    Case "/n" 'filename
                        outFile = s
                        bOption = True
                    Case "/o" 'output dir
                        outDir = s
                        bOption = True
                    Case "/d" 'diff
                        If s.IndexOf("="c) > 1 Then
                            Dim z = s.Split("="c)
                            Dim zt = z(0).Trim
                            If String.IsNullOrEmpty(refName) Then
                                refName = zt
                                msiPath = z(1)
                            Else
                                If Not msiList.ContainsKey(zt) Then msiList.Add(zt, z(1))
                            End If
                        Else
                            bVldArgs = False
                        End If
                End Select
            End If
        Next
    End Sub


    Private Sub AddTable(ByVal tab As String())
        For Each s In tab
            If Not tabList.Contains(s) Then tabList.Add(s)
        Next
    End Sub
    Private Sub AddStore(ByVal sto As String(), ByVal bEmbedded As Boolean)
        For Each s In sto
            If Not stoList.ContainsKey(s) Then stoList.Add(s, bEmbedded)
        Next
    End Sub
    Private Sub SetQueryLevel(ByVal tab As String())
        For Each s In tab
            If s.Equals("t", StringComparison.CurrentCultureIgnoreCase) Or s.Equals("table", StringComparison.CurrentCultureIgnoreCase) Then
                qryLvel = qryLvel And 1
            ElseIf s.Equals("s", StringComparison.CurrentCultureIgnoreCase) Or s.Equals("store", StringComparison.CurrentCultureIgnoreCase) Then
                qryLvel = qryLvel And 2
            End If
        Next
    End Sub
    Public ReadOnly Property IsValidArgs As Boolean
        Get
            Return bVldArgs
        End Get
    End Property
    Public ReadOnly Property DumpSIS As Boolean
        Get
            Return Not bNoSis
        End Get
    End Property
    Public ReadOnly Property DumpLevel() As Integer
        Get
            Return dmpLvel
        End Get
    End Property
    Public ReadOnly Property QueryLevel() As Integer
        Get
            Return qryLvel
        End Get
    End Property
    Public ReadOnly Property XslFile() As String
        Get
            Return ssFile
        End Get
    End Property
    Public ReadOnly Property MsiFile() As String
        Get
            Return msiPath
        End Get
    End Property
    Public ReadOnly Property OutDirectory() As String
        Get
            If outDir Is Nothing Then outDir = New FileInfo(msiPath).Directory.FullName
            Return outDir
        End Get
    End Property
    Public ReadOnly Property OutFilePath(Optional ByVal suffix As String = "") As String
        Get
            If outFile = String.Empty Then
                outFile = Path.GetFileName(msiPath).Replace(".xml", "")
                If suffix <> "" Then outFile = outFile & "_" & suffix
            End If

            If Not outFile.EndsWith(".xml", StringComparison.CurrentCultureIgnoreCase) Then outFile = outFile & ".xml"
            Return OutDirectory() & "\" & outFile
        End Get
    End Property
    Public ReadOnly Property DiffResultFilePath() As String
        Get
            Dim n = refName
            For Each k In msiList.Keys
                n = n & "-" & k
            Next
            Return OutFilePath(n & "_diff.xml")
        End Get
    End Property
    Public ReadOnly Property TableList() As String()
        Get
            Return tabList.ToArray()
        End Get
    End Property
    Public ReadOnly Property StoreList() As String()
        Get
            Dim s As String()
            ReDim s(stoList.Count - 1)
            stoList.Keys.CopyTo(s, 0)
            Return s
        End Get
    End Property
    Public ReadOnly Property HasOptions() As Boolean
        Get
            Return bOption
        End Get
    End Property
    Public ReadOnly Property RefMsiXmlName() As String
        Get
            Return refName
        End Get
    End Property
    Public ReadOnly Property RefMsiXmlPath() As String
        Get
            Return msiPath
        End Get
    End Property
    Public Function GetMsiXmlNames() As IEnumerable(Of String)
        Return msiList.Keys
    End Function
    Public Function GetMsiXmlPath(ByVal name As String) As String
        Return msiList(name)
    End Function
    Public ReadOnly Property IsDiffMode() As Boolean
        Get
            Return bCmpMode
        End Get
    End Property
End Class

Public Class Vars
    Public Const ROOT As String = "msidump"
    Public Const HEADER As String = "header"
    Public Const FIELD As String = "field"
    Public Const DATA As String = "data"
    Public Const TABLE As String = "table"
    Public Const ROW As String = "row"
    Public Const NAME As String = "name"
    Public Const COUNT As String = "count"
    Public Const KEY As String = "key"
    Public Const PATH As String = "path"
    Public Const MST As String = "mst"
    Public Const TYPE As String = "type"
    Public Const ADDED As String = "added"
    Public Const CHANGED As String = "changed"
    Public Const DELETED As String = "deleted"
    Public Const MODE As String = "mode"
    Public Const DIFF As String = "diff"
    Public Const EXECDATE As String = "execDate"
    Public Const INFO As String = "info"
    Public Const _IN As String = "in"
    Public Const _FROM As String = "from"
    Public Const NEWTABLE As String = "addedTable"
    Public Const MODTABLE As String = "modifiedTable"
    Public Const REMTABLE As String = "deletedTable"
    Public Const NEWROW As String = "addedRow"
    Public Const MODROW As String = "modifiedRow"
    Public Const REMROW As String = "deletedRow"
    Public Const _STRING As String = "[string]"
    Public Const _INTEGER As String = "[integer]"
    Public Const _BINARY As String = "[binary]"
    Public Const _SISINFO As String = "_SummaryInformationStream"
    Public Const _PROPERTY As String = "Property"
    Public Const _VALUE As String = "Value"
End Class



'------------------ xmlparser and compare/diff --------------------
Public Class XMsi
    Private sPath As String
    Private sLabel As String
    Private dTable As Dictionary(Of String, XTab)

    Public Sub New(ByVal tLabel As String, ByVal sXml As String)
        Me.sPath = sXml
        Me.sLabel = tLabel
        Me.dTable = New Dictionary(Of String, XTab)()
        ParseXml()
    End Sub

    Private Sub ParseXml()
        Dim mXml = New XmlDocument()
        mXml.Load(Me.sPath)
        Dim xNod = mXml.DocumentElement
        If Vars.ROOT = xNod.Name Then
            For Each cNod As XmlNode In xNod.ChildNodes
                If Vars.TABLE = cNod.Name Then
                    Dim tab = NodeToXTab(cNod)
                    Me.dTable.Add(tab.Name, tab)
                End If
            Next
        End If
    End Sub

    Private Function NodeToXTab(ByRef xNod As XmlNode) As XTab
        Dim tNam = xNod.Attributes(Vars.NAME).Value
        If xNod.Attributes.Count > 1 Then
            tNam = tNam & "::" & xNod.Attributes(Vars.MST).Value
        End If
        Dim tab = New XTab(sLabel, tNam)
        For Each tNod As XmlNode In xNod.ChildNodes
            If Vars.HEADER = tNod.Name Then
                'header info
                tab.SetHdr(NodeToXHdr(tNam, tNod))
            ElseIf Vars.DATA = tNod.Name Then
                'data
                For Each rNod As XmlNode In tNod.ChildNodes
                    tab.AddRow(NodeToXRow(tNam, rNod))
                Next
            End If
        Next

        Return tab
    End Function

    Private Function NodeToXRow(ByVal tabName As String, ByRef xNode As XmlNode) As XRow
        Dim rKey = xNode.Attributes(Vars.KEY).Value
        Dim row = New XRow(sLabel, tabName, rKey)
        For Each rNod As XmlNode In xNode.ChildNodes
            row.SetVal(rNod.Name, rNod.InnerText)
        Next
        Return row
    End Function

    Private Function NodeToXHdr(ByVal tName As String, ByRef xNod As XmlNode) As XHdr
        Dim hCnt = xNod.Attributes(Vars.COUNT).Value
        Dim hDr = New XHdr(tName, hCnt)
        'field nodes
        For Each fNod As XmlNode In xNod.ChildNodes
            Dim fNam = fNod.Attributes(Vars.NAME).Value
            Dim fTyp = fNod.Attributes(Vars.TYPE).Value
            Dim fKey As Boolean = False
            If fNod.Attributes.Count > 2 Then
                fKey = fNod.Attributes(Vars.KEY).Value.Equals("1")
            End If
            Dim xF = New XFld(fNam, fTyp, fKey)
            hDr.AddField(xF)
        Next
        Return hDr
    End Function

    Public Sub DumpTable()
        For Each k As String In Me.dTable.Keys
            Dim xT = dTable(k)
            System.Console.Out.WriteLine("-------------------------------")
            System.Console.Out.WriteLine(xT.Name & ": " & xT.Count)
            System.Console.Out.WriteLine("-[Hdr]-------------------------")
            xT.GetHdr().Dump()
            System.Console.Out.WriteLine("-[Dat] " & xT.Count & "-------------------------")
            For Each xR As String In xT.RowKeys
                xT.GetRow(xR).Dump()
            Next
            System.Console.Out.WriteLine("-------------------------------")
        Next
    End Sub

    Public ReadOnly Property TableNames() As IEnumerable(Of String)
        Get
            Return dTable.Keys
        End Get
    End Property

    Public Function GetXTab(ByVal tabName As String) As XTab
        Return dTable(tabName)
    End Function

    Public ReadOnly Property Label() As String
        Get
            Return sLabel
        End Get
    End Property

    Public Function TableExists(ByVal tabName As String) As Boolean
        Return dTable.ContainsKey(tabName)
    End Function
End Class

Public Class XTab
    Private sName, mLabel As String
    Private tRows As Dictionary(Of String, XRow)
    Private fHdr As XHdr
    Public Sub New(ByVal tLabel As String, ByVal tName As String)
        Me.mLabel = tLabel
        Me.sName = tName
        Me.tRows = New Dictionary(Of String, XRow)()
    End Sub
    Public Sub SetHdr(ByRef xH As XHdr)
        Me.fHdr = xH
    End Sub
    Public Function GetHdr() As XHdr
        Return Me.fHdr
    End Function
    Public Sub AddRow(ByRef xRow As XRow)
        Me.tRows.Add(xRow.Key, xRow)
    End Sub
    Public ReadOnly Property RowKeys() As ICollection(Of String)
        Get
            Return Me.tRows.Keys
        End Get
    End Property

    Public Function GetRow(ByVal rKey As String) As XRow
        Return Me.tRows(rKey)
    End Function
    Public ReadOnly Property Count() As String
        Get
            Return Me.tRows.Count
        End Get
    End Property
    Public ReadOnly Property Name() As String
        Get
            Return Me.sName
        End Get
    End Property
    Public ReadOnly Property MsiLabel() As String
        Get
            Return Me.mLabel
        End Get
    End Property
    Public Sub RemoveRow(ByVal rKey As String)
        Me.tRows.Remove(rKey)
    End Sub
    Public Function RowExists(ByVal rowKey As String) As Boolean
        Return tRows.ContainsKey(rowKey)
    End Function

    Public Sub WriteXML(ByRef xml As XmlTextWriter)
        xml.WriteStartElement(Vars.TABLE)
        xml.WriteAttributeString(Vars.NAME, Name)
        'header
        Me.fHdr.WriteXML(xml)
        'data
        xml.WriteStartElement(Vars.DATA)
        xml.WriteAttributeString(Vars.COUNT, tRows.Count)
        For Each r In tRows.Values
            r.WriteXML(xml)
        Next
        xml.WriteEndElement()
        xml.WriteEndElement()
    End Sub
End Class

Public Class XRow
    Private sKey, mLbl, tName As String
    Private rVal As Dictionary(Of String, String)
    Public Sub New(ByVal mLabel As String, ByVal tNam As String, ByVal rKey As String)
        Me.mLbl = mLabel
        Me.tName = tNam
        Me.sKey = rKey
        Me.rVal = New Dictionary(Of String, String)()
    End Sub
    Public ReadOnly Property Key() As String
        Get
            Return Me.sKey
        End Get
    End Property
    Public ReadOnly Property MsiLabel() As String
        Get
            Return Me.mLbl
        End Get
    End Property
    Public ReadOnly Property TableName() As String
        Get
            Return Me.tName
        End Get
    End Property
    Public Sub SetVal(ByVal cName As String, ByVal cVal As String)
        Me.rVal.Add(cName, cVal)
    End Sub
    Public Function GetVal(ByVal cName As String) As String
        Return Me.rVal(cName)
    End Function
    Public ReadOnly Property FieldNames() As IEnumerable(Of String)
        Get
            Return rVal.Keys
        End Get
    End Property
    Public Sub Dump()
        Dim s = "[R " & Me.sKey & "] "
        For Each k In rVal.Keys
            s = s & vbTab & k & "=" & rVal(k)
        Next
        System.Console.Out.WriteLine(s)
    End Sub
    Public Sub WriteXML(ByRef xml As XmlTextWriter)
        xml.WriteStartElement(Vars.ROW)
        xml.WriteAttributeString(Vars.KEY, Me.sKey)
        WriteValues(xml)
        xml.WriteEndElement()
    End Sub
    Public Sub WriteValues(ByRef xml As XmlTextWriter)
        For Each k In rVal.Keys
            xml.WriteStartElement(k)
            xml.WriteString(rVal(k))
            xml.WriteEndElement()
        Next
    End Sub

End Class

'cross-release diff
Public Class XTDiff
    Private refXMsi As XMsi
    Private newTC = 0, remTC = 0, modTC = 0
    Private cmpXMsis As Dictionary(Of String, XMsi)
    Private cmpXRes As Dictionary(Of String, XTRes) 'key is msiLabel
    Private grpTabs As Dictionary(Of String, XTabs) 'key is tableName

    Public Sub New(ByRef roMsi As XMsi)
        Me.refXMsi = roMsi
        Me.cmpXMsis = New Dictionary(Of String, XMsi)()
        Me.cmpXRes = New Dictionary(Of String, XTRes)()
        Me.grpTabs = New Dictionary(Of String, XTabs)()
    End Sub
    Public Sub AddMsiXml(ByVal name As String, ByVal path As String)
        If Me.cmpXMsis.ContainsKey(name) Then Return
        Dim xM = New XMsi(name, path)
        Me.cmpXMsis.Add(name, xM)
        Me.Compare(xM)
    End Sub
    Private Sub Compare(ByRef cmpX As XMsi)
        Dim xR = New XTRes(refXMsi.Label, cmpX.Label)
        For Each tN In refXMsi.TableNames
            System.Console.Out.WriteLine("Comparing Table: " & tN)
            If cmpX.TableExists(tN) Then
                If Not grpTabs.ContainsKey(tN) Then
                    grpTabs.Add(tN, New XTabs(refXMsi.GetXTab(tN)))
                End If
                If (grpTabs(tN).AddSubXTab(cmpX.GetXTab(tN))) Then
                    modTC = modTC + 1
                    System.Console.Out.WriteLine("ChangedTable: " & tN)
                End If

            Else
                xR.AddDelTables(refXMsi.GetXTab(tN))
                System.Console.Out.WriteLine("DeletedTable: " & tN)
            End If
        Next
        'check for added
        For Each tN In cmpX.TableNames
            If Not refXMsi.TableExists(tN) Then
                xR.AddNewTables(cmpX.GetXTab(tN))
                System.Console.Out.WriteLine("AddedTable: " & tN)
            End If
        Next
        newTC = newTC + xR.NewTableCount
        remTC = remTC + xR.DelTableCount

        Me.cmpXRes.Add(cmpX.Label, xR)
    End Sub
    Public Sub WriteResult(ByRef opt As Opt)

        Dim xml = New XmlTextWriter(opt.DiffResultFilePath, Text.Encoding.UTF8)
        If Not opt.XslFile Is Nothing Then
            Dim PItext = "type='text/xsl' href='" + opt.XslFile + "'"
            xml.WriteProcessingInstruction("xml-stylesheet", PItext)
        End If

        xml.WriteStartElement(Vars.ROOT)
        xml.WriteAttributeString(Vars.MODE, Vars.DIFF)
        xml.WriteAttributeString(Vars.EXECDATE, Now.ToString("yyyy.MM.dd.HH.mm.ss"))

        xml.WriteStartElement(Vars.INFO) 'info

        xml.WriteStartElement(opt.RefMsiXmlName) 'msiname
        xml.WriteAttributeString("ref", "1")
        xml.WriteString(opt.RefMsiXmlPath)
        xml.WriteEndElement() 'refmsixml
        For Each mx In opt.GetMsiXmlNames()
            xml.WriteStartElement(mx)
            xml.WriteString(opt.GetMsiXmlPath(mx))
            xml.WriteEndElement() 'submsixml
        Next
        xml.Flush()
        xml.WriteEndElement() 'info

        'added table
        If newTC > 0 Then
            For Each xm In Me.cmpXRes.Keys
                Dim xt = cmpXRes(xm)
                If xt.NewTableCount > 0 Then
                    xml.WriteStartElement(Vars.NEWTABLE)
                    xml.WriteAttributeString(Vars._IN, xm)
                    xml.WriteAttributeString(Vars.COUNT, xt.NewTableCount)
                    For Each nt In xt.NewTableNames
                        xt.GetNewXTab(nt).WriteXML(xml)
                    Next
                    xml.WriteEndElement()
                End If
            Next
        End If

        'removed table
        If remTC > 0 Then
            For Each xm In Me.cmpXRes.Keys
                Dim xt = cmpXRes(xm)
                If xt.DelTableCount > 0 Then
                    xml.WriteStartElement(Vars.REMTABLE)
                    xml.WriteAttributeString(Vars._IN, xm)
                    xml.WriteAttributeString(Vars.COUNT, xt.DelTableCount)
                    For Each nt In xt.DelTableNames
                        xt.GetDelXTab(nt).WriteXML(xml)
                    Next
                    xml.WriteEndElement()
                End If
            Next
        End If

        'modified table
        If modTC > 0 Then
            xml.WriteStartElement(Vars.MODTABLE)
            xml.WriteAttributeString(Vars.COUNT, Me.modTC)
            For Each tN In Me.grpTabs.Values
                If tN.XRDiff.HasChanges Then tN.XRDiff.WriteResult(xml)
            Next
            xml.WriteEndElement()
        End If
        xml.Flush()
        xml.Close()
    End Sub

   
End Class

Public Class XRDiff
    Private refXTab As XTab
    Private newRC = 0, remRC = 0, modRC = 0
    Private cmpXRes As Dictionary(Of String, XRRes) 'key is msiLabel
    Private grpRows As Dictionary(Of String, XRows) 'key is rowKey

    Public Sub New(ByRef roTab As XTab)
        Me.refXTab = roTab
        Me.cmpXRes = New Dictionary(Of String, XRRes)()
        Me.grpRows = New Dictionary(Of String, XRows)()
    End Sub
    Public ReadOnly Property HasChanges() As Boolean
        Get
            Return (newRC + remRC + modRC) > 0
        End Get
    End Property
    Public ReadOnly Property NewRowCount() As Integer
        Get
            Return newRC
        End Get
    End Property
    Public ReadOnly Property DelRowCount() As Integer
        Get
            Return remRC
        End Get
    End Property
    Public ReadOnly Property ModRowCount() As Integer
        Get
            Return modRC
        End Get
    End Property
    Public Function Compare(ByRef cmpX As XTab) As Boolean
        Dim xR = New XRRes(refXTab.MsiLabel, cmpX.MsiLabel)
        System.Console.Out.WriteLine(vbTab & "ComparingRows of Table: " & cmpX.Name)
        'loop on each row in the reference table and check if the rowkey exists on the compared subject
        For Each rK In refXTab.RowKeys
            If cmpX.RowExists(rK) Then
                If Not grpRows.ContainsKey(rK) Then
                    grpRows.Add(rK, New XRows(refXTab.GetRow(rK)))
                End If
                'add to group and do compare internally
                If grpRows(rK).AddSubXRow(cmpX.GetRow(rK)) > 0 Then
                    xR.AddModRow(refXTab.GetRow(rK))
                    System.Console.Out.WriteLine(vbTab & "ModifiedRow: " & rK)
                End If
            Else
                'ref rowkey does not exist in the compared subject -> indicating it is removed/deleted
                xR.AddDelRow(refXTab.GetRow(rK))
                System.Console.Out.WriteLine(vbTab & "DeletedRow: " & rK)
            End If
        Next
        'leftover cmp rowkeys indicate newly added rows in the compared subject
        For Each tN In cmpX.RowKeys
            If Not refXTab.RowExists(tN) Then
                xR.AddNewRow(cmpX.GetRow(tN))
                System.Console.Out.WriteLine(vbTab & "AddedRow: " & tN)
            End If
        Next

        newRC = newRC + xR.NewRowCount
        remRC = remRC + xR.DelRowCount
        modRC = modRC + xR.ModRowCount
        Me.cmpXRes.Add(cmpX.MsiLabel, xR)
        System.Console.Out.WriteLine(vbTab & "New:" & newRC & ", Del:" & remRC & ", Mod:" & modRC & ", HasChanges: " & HasChanges)
        Return HasChanges
    End Function
    Public Sub WriteResult(ByRef xml As XmlTextWriter)

        xml.WriteStartElement(Vars.TABLE)
        xml.WriteAttributeString(Vars.NAME, refXTab.Name)
        'added rows
        If newRC > 0 Then
            For Each xm In Me.cmpXRes.Keys
                Dim xt = cmpXRes(xm)
                If xt.NewRowCount > 0 Then
                    xml.WriteStartElement(Vars.NEWROW)
                    xml.WriteAttributeString(Vars._IN, xm)
                    xml.WriteAttributeString(Vars.COUNT, xt.NewRowCount)
                    For Each nt In xt.NewRowKeys
                        xt.GetNewXRow(nt).WriteXML(xml)
                    Next
                    xml.WriteEndElement()
                End If

            Next
        End If

        'removed rows
        If remRC > 0 Then
            For Each xm In Me.cmpXRes.Keys
                Dim xt = cmpXRes(xm)
                If xt.DelRowCount > 0 Then
                    xml.WriteStartElement(Vars.REMROW)
                    xml.WriteAttributeString(Vars._IN, xm)
                    xml.WriteAttributeString(Vars.COUNT, xt.DelRowCount)
                    For Each nt In xt.DelRowKeys
                        xt.GetDelXRow(nt).WriteXML(xml)
                    Next
                    xml.WriteEndElement()
                End If

            Next
        End If

        If modRC > 0 Then
            For Each xm In Me.cmpXRes.Keys
                Dim xt = cmpXRes(xm)
                If xt.ModRowCount > 0 Then
                    xml.WriteStartElement(Vars.MODROW)
                    xml.WriteAttributeString(Vars._IN, xm)
                    xml.WriteAttributeString(Vars.COUNT, xt.ModRowCount)
                    For Each nt In xt.ModRowKeys
                        Me.grpRows(nt).WriteChangedValues(xml)
                    Next
                    xml.WriteEndElement()
                End If

            Next
            'xml.WriteStartElement(Vars.CHANGED)
            'For Each x In Me.grpRows.Values
            '    If x.IsModified Then x.WriteChangedValues(xml)
            'Next
            'xml.WriteEndElement()
        End If
        xml.WriteEndElement()
    End Sub
End Class

'grouped rows
Public Class XRows
    Private refRow As XRow
    Private modFlds As List(Of String)
    Private subRows As List(Of XRow)
    Public Sub New(ByRef oRow As XRow)
        Me.refRow = oRow
        Me.subRows = New List(Of XRow)()
        Me.modFlds = New List(Of String)()
    End Sub
    Public Function AddSubXRow(ByRef xRow As XRow) As Integer
        Me.subRows.Add(xRow)
        Me.Compare(xRow)
        Return modFlds.Count
    End Function
    Public ReadOnly Property Count() As Integer
        Get
            Return Me.subRows.Count
        End Get
    End Property
    Public Function SubXRows() As IList(Of XRow)
        Return Me.subRows
    End Function
    Public ReadOnly Property Key() As String
        Get
            Return Me.refRow.Key
        End Get
    End Property
    Public Sub ToXml(ByRef xml As XmlTextWriter)
        xml.WriteStartElement(Vars.ROW)
        xml.WriteAttributeString(Vars.KEY, Me.refRow.Key)
        For Each k In refRow.FieldNames
            xml.WriteStartElement(k)
            Dim rV = refRow.GetVal(k)
            xml.WriteStartElement(refRow.MsiLabel) 'refNode
            xml.WriteString(rV)   'refString
            xml.WriteEndElement()            'endNode
            For Each sRow In Me.subRows
                xml.WriteStartElement(sRow.MsiLabel)
                Dim cV = sRow.GetVal(k)
                If rV <> cV Then xml.WriteAttributeString(Vars.CHANGED, "1")
                xml.WriteString(cV)
                xml.WriteEndElement()
            Next
            xml.WriteEndElement()
        Next
        xml.WriteEndElement()
    End Sub
    Public Sub WriteChangedValues(ByRef xml As XmlTextWriter)
        xml.WriteStartElement(Vars.ROW)
        xml.WriteAttributeString(Vars.KEY, Me.refRow.Key)
        For Each k In Me.modFlds
            xml.WriteStartElement(k)
            xml.WriteStartElement(refRow.MsiLabel) 'refNode
            Dim rV = refRow.GetVal(k)
            xml.WriteString(rV)   'refString
            xml.WriteEndElement()            'endNode
            For Each sRow In Me.subRows
                xml.WriteStartElement(sRow.MsiLabel)
                xml.WriteString(sRow.GetVal(k))
                xml.WriteEndElement()
            Next
            xml.WriteEndElement()
        Next
        xml.WriteEndElement()
    End Sub
    Public ReadOnly Property IsModified() As Boolean
        Get
            Return modFlds.Count > 0
        End Get
    End Property

    '-1: diff key, 0: no changes 1>: number of changed fields
    Private Function Compare(ByRef cmpRow As XRow) As Boolean
        For Each k In refRow.FieldNames
            If (Not modFlds.Contains(k)) And (refRow.GetVal(k) <> cmpRow.GetVal(k)) Then Me.modFlds.Add(k)
        Next
        Return IsModified
    End Function
End Class

'grouped tables
Public Class XTabs
    Private refTab As XTab
    Private refNam As String
    Private subTabs As Dictionary(Of String, XTab)
    Private xrD As XRDiff
    Public Sub New(ByRef oTab As XTab)
        Me.refTab = oTab
        Me.subTabs = New Dictionary(Of String, XTab)()
        Me.xrD = New XRDiff(oTab)
    End Sub
    Public Function AddSubXTab(ByRef xTab As XTab) As Boolean
        Me.subTabs.Add(xTab.MsiLabel, xTab)
        Return Me.xrD.Compare(xTab)
    End Function

    Public ReadOnly Property Count() As Integer
        Get
            Return Me.subTabs.Count
        End Get
    End Property
    Public Function SubXRows() As IList(Of XTab)
        Return Me.subTabs
    End Function
    Public ReadOnly Property TableName As String
        Get
            Return Me.refTab.Name
        End Get
    End Property
    Public Sub ToXml(ByRef xml As XmlTextWriter)
        xml.WriteStartElement(Vars.TABLE)
        xml.WriteAttributeString(Vars.NAME, Me.refTab.Name)
        For Each k In refTab.RowKeys
            xml.WriteStartElement(k)
            Dim rV = refTab.GetRow(k)
            xml.WriteStartElement(refTab.MsiLabel) 'refNode
            'xml.WriteString(rV)   'refString
            xml.WriteEndElement()            'endNode
            For Each sTab In Me.subTabs.Keys
                xml.WriteStartElement(sTab)
                'Dim cV = sRow.GetVal(k)
                'If rV <> cV Then xml.WriteAttributeString(Vars.CHANGED, "1")
                'xml.WriteString(cV)
                xml.WriteEndElement()
            Next
            xml.WriteEndElement()
        Next
        xml.WriteEndElement()
    End Sub
    Public ReadOnly Property XRDiff() As XRDiff
        Get
            Return Me.xrD
        End Get
    End Property
End Class


'result
Public Class XTRes
    Private sFr, sTo As String
    Private newTables, remTables As Dictionary(Of String, XTab)
    Public Sub New(ByVal refName As String, ByVal cmpName As String)
        sFr = refName
        sTo = cmpName
        newTables = New Dictionary(Of String, XTab)()
        remTables = New Dictionary(Of String, XTab)()
    End Sub
    Public ReadOnly Property RefName() As String
        Get
            Return sFr
        End Get
    End Property
    Public ReadOnly Property CmpName() As String
        Get
            Return sTo
        End Get
    End Property
    Public Sub AddNewTables(ByRef xT As XTab)
        newTables.Add(xT.Name, xT)
    End Sub
    Public Sub AddDelTables(ByRef xT As XTab)
        remTables.Add(xT.Name, xT)
    End Sub
    Public ReadOnly Property NewTableCount() As Integer
        Get
            Return newTables.Count
        End Get
    End Property
    Public ReadOnly Property DelTableCount() As Integer
        Get
            Return remTables.Count
        End Get
    End Property
    Public ReadOnly Property NewTableNames() As IEnumerable(Of String)
        Get
            Return newTables.Keys
        End Get
    End Property
    Public ReadOnly Property DelTableNames() As IEnumerable(Of String)
        Get
            Return remTables.Keys
        End Get
    End Property
    Public Function GetNewXTab(ByVal tabName As String) As XTab
        Return newTables(tabName)
    End Function
    Public Function GetDelXTab(ByVal tabName As String) As XTab
        Return remTables(tabName)
    End Function
    Public Function IsDeleted(ByVal tabName As String) As Boolean
        Return remTables.ContainsKey(tabName)
    End Function
End Class

Public Class XRRes
    Private sFr, sTo As String
    Private newRows, remRows, modRows As Dictionary(Of String, XRow)
    Public Sub New(ByVal refName As String, ByVal cmpName As String)
        sFr = refName
        sTo = cmpName
        newRows = New Dictionary(Of String, XRow)()
        remRows = New Dictionary(Of String, XRow)()
        modRows = New Dictionary(Of String, XRow)()
    End Sub
    Public ReadOnly Property RefName() As String
        Get
            Return sFr
        End Get
    End Property
    Public ReadOnly Property CmpName() As String
        Get
            Return sTo
        End Get
    End Property
    Public Sub AddNewRow(ByRef xR As XRow)
        newRows.Add(xR.Key, xR)
    End Sub
    Public Sub AddDelRow(ByRef xR As XRow)
        remRows.Add(xR.Key, xR)
    End Sub
    Public Sub AddModRow(ByRef xR As XRow)
        modRows.Add(xR.Key, xR)
    End Sub
    Public ReadOnly Property NewRowCount() As Integer
        Get
            Return newRows.Count
        End Get
    End Property
    Public ReadOnly Property DelRowCount() As Integer
        Get
            Return remRows.Count
        End Get
    End Property
    Public ReadOnly Property ModRowCount() As Integer
        Get
            Return modRows.Count
        End Get
    End Property
    Public ReadOnly Property NewRowKeys() As IEnumerable(Of String)
        Get
            Return newRows.Keys
        End Get
    End Property
    Public ReadOnly Property DelRowKeys() As IEnumerable(Of String)
        Get
            Return remRows.Keys
        End Get
    End Property
    Public ReadOnly Property ModRowKeys() As IEnumerable(Of String)
        Get
            Return modRows.Keys
        End Get
    End Property
    Public Function GetNewXRow(ByVal rowKey As String) As XRow
        Return newRows(rowKey)
    End Function
    Public Function GetDelXRow(ByVal rowKey As String) As XRow
        Return remRows(rowKey)
    End Function
    Public Function GetModXRow(ByVal rowKey As String) As XRow
        Return modRows(rowKey)
    End Function
End Class

Public Class XHdr
    Private dicFields As Dictionary(Of String, XFld)
    Private dicKeys As Dictionary(Of String, XFld)
    Private tabName As String
    Public Sub New(ByVal tName As String, ByVal hdrCount As Integer)
        Me.tabName = tName
        Me.dicFields = New Dictionary(Of String, XFld)(hdrCount)
        Me.dicKeys = New Dictionary(Of String, XFld)(3)
    End Sub
    Public ReadOnly Property Name() As String
        Get
            Return Me.tabName
        End Get
    End Property
    Public Sub AddField(ByRef xFld As XFld)
        Me.dicFields.Add(xFld.Name, xFld)
        If xFld.IsKey Then
            Me.dicKeys.Add(xFld.Name, xFld)
        End If

    End Sub
    Public Function GetField(ByVal fldName As String) As XFld
        Return Me.dicFields(fldName)
    End Function
    Public ReadOnly Property KeyCount As Integer
        Get
            Return Me.dicKeys.Count
        End Get
    End Property
    Public ReadOnly Property FieldCount As Integer
        Get
            Return Me.dicFields.Count
        End Get
    End Property
    Public ReadOnly Property KeyNames() As ICollection(Of String)
        Get
            Return Me.dicKeys.Keys
        End Get
    End Property
    Public ReadOnly Property FieldNames() As ICollection(Of String)
        Get
            Return Me.dicFields.Keys
        End Get
    End Property

    Public Sub Dump()
        For Each x In dicFields.Keys
            dicFields(x).Dump()
        Next
    End Sub
    Public Sub WriteXML(ByRef xml As XmlTextWriter)
        xml.WriteStartElement(Vars.HEADER)
        xml.WriteAttributeString(Vars.COUNT, FieldCount)
        For Each f In dicFields.Values
            f.WriteXML(xml)
        Next
        xml.WriteEndElement()
    End Sub
End Class

Public Class XFld
    Private fieldName, fieldType As String
    Private bIsKey As Boolean

    Public Sub New(ByVal fldName As String, ByVal fldType As String, ByVal bKey As Boolean)
        Me.fieldName = fldName
        Me.fieldType = fldType
        Me.bIsKey = bKey
    End Sub

    Public ReadOnly Property Name() As String
        Get
            Return fieldName
        End Get
    End Property

    Public ReadOnly Property Type() As String
        Get
            Return fieldType
        End Get
    End Property

    Public ReadOnly Property IsKey() As Boolean
        Get
            Return bIsKey
        End Get
    End Property

    Public Sub Dump()
        System.Console.Out.WriteLine("[F] " & vbTab & fieldName & "." & fieldType & "." & bIsKey)
    End Sub

    Public Sub WriteXML(ByRef xml As XmlTextWriter)
        xml.WriteStartElement(Vars.FIELD)
        xml.WriteAttributeString(Vars.NAME, fieldName)
        xml.WriteAttributeString(Vars.TYPE, fieldType)
        If IsKey() Then xml.WriteAttributeString(Vars.KEY, "1")
        xml.WriteEndElement()
    End Sub
End Class