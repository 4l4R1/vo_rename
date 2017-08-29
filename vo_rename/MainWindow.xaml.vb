Imports System.Globalization
Imports System.IO
Imports System.Windows.Media

Class MainWindow
    Implements IDisposable
    Private watcher As New FileSystemWatcher
    'Private tDir As String = "E:\Eetrigraafika\"
    'Private tDir As String = "\\192.168.73.22\eetrigraafika\"
    Private wDir As String = "E:\_home_\var\audacious"
    'Private eDir As String = "E:\Eetrigraafika\Kanal2\sk1415\voiceover\"
    'Private fDir As String = "\\INCA2\eetrigraafika\kanal2\sk1415\voiceover\"
    Private eDir As String = "E:\Caspar\Eetrigraafika\Kanal2\16-sk\voiceover"
    'Private fDir As String = "\\Caspar_KX-HP\Caspar\Eetrigraafika\Kanal2\16-sk\voiceover"
    Private bDir As String = "E:\Eetrigraafika\Kanal2\sk16\voiceover"

    Private dDir As New List(Of String)
    'From {
    '    "\\CASPAR_K2-HP\Caspar\Eetrigraafika\Kanal2\16-sk\voiceover",
    '    "\\CASPAR_K11-HP\Caspar\Eetrigraafika\Kanal2\16-sk\voiceover",
    '    "\\CASPAR_K12-HP\Caspar\Eetrigraafika\Kanal2\16-sk\voiceover",
    '    "\\CASPAR1-HP\Caspar\Eetrigraafika\Kanal2\16-sk\voiceover"}

    'Private gDir As String = "\\192.168.73.71\Metus2\kanal2\guide\"
    Private setupFn As String = "E:\Eetrigraafika\conf\RenameSetup.xml"

    Private currentFn As String = String.Empty
    Private lFiles As New List(Of String)
    Delegate Sub SimpleDelegate()
    Private delExport As New SimpleDelegate(AddressOf export)
    Private delRename As New SimpleDelegate(AddressOf rename)

    'Private lTemplate As String = "{0:MMdd_HHmm}" + vbTab + "{1}"
    'Private lTemplate As String = "{0:MMdd_HHmm}" + vbTab + "{0:yyMMdd}-{1}" + vbTab + "{2}"
    'Private lTemplate As String = "{0:MM.dd-HH:mm}" + vbTab + "{1:yyMMdd}-{2}" + vbTab + "{3}"
    'Private RejectedLabels As New List(Of String)

    Private startTimes As New List(Of TimeSpan)
    Private stopTimes As New List(Of TimeSpan)

    Private RenameList As New List(Of String) From {"14-100-00052", "14-100-00053", "15-100-00022", "15-100-00023", "07-100-00259"}

    Private Sub MainWindow_Loaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Loaded

        'For Each timeStr As String In {"20:05", "20:05", "19:45", "19:45", "19:45", "19:30", "19:30"}	 'algusajad esmasp - pühap 	   et 2005, muidu 19:45	LP 1930
        '	startTimes.Add(GetSpan(timeStr))
        'Next

        'For Each timeStr As String In {"23:45", "00:05", "00:05", "00:05", "00:05", "00:05", "00:05"}	'lõpuajad esmasp - pühap 
        '	stopTimes.Add(GetSpan(timeStr))
        'Next

        'LoadLabels()
        LoadSetup()

        If cName.Items.Count > 0 Then cName.SelectedIndex = 0

        initWatcher()
    End Sub

    Private Sub MainWindow_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Closing
        If Not String.IsNullOrWhiteSpace(currentFn) Then e.Cancel = True
    End Sub

    Private Sub bOK_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles bOK.Click
        If String.IsNullOrWhiteSpace(currentFn) Then
            If MessageBox.Show("Send all files to " + dDir.First, String.Format("{0} files in " + eDir, lFiles.Count), MessageBoxButton.OKCancel) = MessageBoxResult.OK Then
                'For Each fn As String In lFiles
                '    If File.Exists(fn) Then
                '        Try
                '            File.Copy(fn, fn.Replace(eDir, fDir), True)
                '        Catch ex As Exception

                '        End Try
                '    End If
                'Next

                'lFiles.Clear()
                Distribute()
                Archive()

                End

            End If
            Return
        End If

        export()
        cName.Foreground = Brushes.DarkRed
        Title = "VO Renamer"
    End Sub

    Private Sub bAdd_Click(ByVal sender As Object, ByVal e As RoutedEventArgs) Handles bAdd.Click
        'Dim msg1 As String = If(cName.SelectedIndex < 0, String.Empty, cName.SelectedItem.ToString)
        'Dim msg2 As String = InputBox("Name (MMdd_HHmm)", "Add New label", msg1)
        ''If msg.Length < 9 Then Return
        'If String.IsNullOrWhiteSpace(msg2) Then Return


        'If msg1.Substring(0, 9) = msg2.Substring(0, 9) Then
        '	cName.Items(cName.SelectedIndex) = msg2
        'Else
        '	cName.Items.Insert(cName.SelectedIndex, msg2)
        '	cName.SelectedIndex -= 1
        'End If


        EditXml(setupFn)
    End Sub

    Private Sub cName_SelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles cName.SelectionChanged
        If String.IsNullOrWhiteSpace(currentFn) Then
            'Me.Title = "VO Renamer"
            Return
        End If

        watcher.EnableRaisingEvents = False
        rename()
        watcher.EnableRaisingEvents = True
    End Sub

    Private Sub export()
        Dim expFn As String = currentFn.Replace(wDir, eDir)
        If File.Exists(expFn) Then
            If MessageBox.Show(expFn + " exists. Overwrite?", expFn, MessageBoxButton.OKCancel) = MessageBoxResult.OK Then
                File.Delete(expFn)
            Else
                Return
            End If

        End If

        Try
            File.Move(currentFn, expFn)
            currentFn = String.Empty
            If Not lFiles.Contains(expFn) Then lFiles.Add(expFn)
        Catch ex As Exception

        End Try


        cName.Foreground = Brushes.Black
        If cName.SelectedIndex = cName.Items.Count - 1 Then Return
        cName.SelectedIndex += 1
    End Sub

    Private Sub rename()
        If cName.SelectedIndex < 0 Then Return

        Dim fnSplit() As String = cName.SelectedItem.ToString.Split(vbTab.First)
        Dim fn As String = fnSplit(1) ' cName.SelectedItem.ToString.Substring(0, 9)
        Dim fi As New FileInfo(currentFn)
        fn = fi.DirectoryName + "\" + fn + fi.Extension
        If fn = currentFn Then Return
        If File.Exists(fn) Then
            If MessageBox.Show(fn + " exists. Overwrite?", fn, MessageBoxButton.OKCancel) = MessageBoxResult.OK Then
                File.Delete(fn)
            Else
                Return
            End If
        End If
        Try
            File.Move(currentFn, fn)
            currentFn = fn
        Catch ex As Exception

        End Try

        cName.Foreground = Brushes.DarkBlue

        If cName.SelectedIndex = cName.Items.Count - 1 Then
            Title = "VO Renamer"
            Return
        End If

        'Me.Title = cName.Items(cName.SelectedIndex + 1).ToString.Replace(vbTab, "  ")
        fnSplit = cName.Items(cName.SelectedIndex + 1).ToString.Split(vbTab.First)
        Title = fnSplit(0) + "  " + fnSplit(2)
    End Sub

    Private Sub initWatcher()

        Try
            With watcher
                '.SynchronizingObject = SynchronizationContext
                .EnableRaisingEvents() = False
                .Path = wDir
                .Filter = "*.wav"
                .NotifyFilter = IO.NotifyFilters.LastWrite Or IO.NotifyFilters.CreationTime

                AddHandler watcher.Changed, AddressOf Watcher_Changed
                AddHandler watcher.Created, AddressOf Watcher_Changed

                .EnableRaisingEvents() = True
            End With
        Catch ex As Exception
            'tPaste.Text = ex.Message
        End Try

    End Sub

    Private Sub Watcher_Changed(ByVal sender As System.Object, ByVal e As System.IO.FileSystemEventArgs) 'Handles watcher.Changed, watcher.Created

        Select Case e.ChangeType
            Case IO.WatcherChangeTypes.Changed, IO.WatcherChangeTypes.Created
                watcher.EnableRaisingEvents = False
                Dim fn As String = e.FullPath
                Do Until ExclusiveAccessToFile(fn)
                    System.Threading.Thread.Sleep(100)
                Loop
                If Not String.IsNullOrWhiteSpace(currentFn) Then
                    'export())
                    Dispatcher.Invoke(delExport)
                End If

                currentFn = e.FullPath
                'rename()
                Dispatcher.Invoke(delRename)
                watcher.EnableRaisingEvents = True
            Case Else
                'nothing
        End Select
    End Sub

    Private Function ExclusiveAccessToFile(ByVal fullPath As String) As Boolean
        Try
            Dim f As Stream
            f = File.Open(fullPath, FileMode.Append, FileAccess.Write, FileShare.None)
            f.Close()
            Return True
        Catch e As Exception
            Return False
        End Try
    End Function

    Private Sub LoadSetup()
        Dim sDoc As XDocument = XDocument.Load(setupFn)
        startTimes.Clear()
        stopTimes.Clear()

        For Each xDay As XElement In sDoc.Root.Elements
            startTimes.Add(GetSpan(xDay.@from))
            stopTimes.Add(GetSpan(xDay.@to))
        Next

        LoadLabels()
        InitDest("E:\Caspar\Eetrigraafika\conf\syncSetup.xml")
    End Sub

    Private Sub LoadLabels()
        Dim _ptrn As String = "\\192.168.73.71\Metus2\kanal2\guide\programme_guide_Kanal 2_{0:dd.MM.yyyy}-{1:dd.MM.yyyy}.xml"

        Dim nr As Integer = Now.DayOfWeek - 1
        If nr < 0 Then nr = 6 'pühapäev
        Dim dMond As Date = Now.Date.AddDays(-1 * nr)   'kontrollitava nädala esmaspäev
        If nr > 3 Then dMond = dMond.AddDays(7) 'reedest otsi uut nädalat

        Dim lLabel As New List(Of String)

        For i As Integer = 0 To 1 'kontrollitav ja järgmine nädal
            Dim Fn As String = String.Format(_ptrn, dMond.AddDays(i * 7), dMond.AddDays(i * 7 + 6))
            If File.Exists(Fn) Then AddLabels(Fn, lLabel)
        Next

        With cName.Items
            .Clear()
            For Each lbl As String In lLabel
                .Add(lbl)
            Next
        End With
    End Sub

    Private Sub AddLabels(ByVal fn As String, ByVal lbl As List(Of String))
        Try
            Dim guideDoc As XDocument
            guideDoc = XDocument.Load(fn)

            For Each el As XElement In guideDoc.Root.Elements("row")
                Dim dCntrl As Date
                Dim dEl As IEnumerable(Of XElement) = el.Elements("Kuupaev")

                If dEl.Count > 0 Then
                    dCntrl = Date.Parse(dEl.First.Value).Date
                    Dim nr As Integer = dCntrl.DayOfWeek - 1
                    If nr < 0 Then nr = 6 'pühapäev

                    dEl = el.Elements("Kell")
                    If dEl.Count > 0 Then
                        Dim ts As TimeSpan = GetSpan(dEl.First.Value)
                        Dim hMin As Integer = (Integer.Parse(dEl.First.Value.Substring(0, 2)) - 1) Mod 24
                        Dim hMax As Integer = (hMin + 2) Mod 24
                        Dim pr As String = PrRename(el.Elements("Production_nr").First.Value.Replace("/"c, "-"))



                        If ts.TotalMinutes > startTimes(nr).TotalMinutes AndAlso ts.TotalMinutes < stopTimes(nr).TotalMinutes Then   ' (19:30.st 00:05ni)
                            'Dim label As String = String.Format(lTemplate, dCntrl + ts, dCntrl, pr, el.Elements("Saade").First.Value)
                            Dim label As String = $"{dCntrl + ts:MM.dd-HH:mm}{vbTab}{dCntrl:yyMMdd}-{hMin:00}-{hMax:00},{pr}{vbTab}{el.Elements("Saade").First.Value}"

                            lbl.Add(label)    'dCntrl.ToString("MMdd_HHmm"))

                        End If
                    End If
                End If
            Next
        Catch ex As Exception
            'Log.ErrOut(ex)
        End Try
    End Sub

    Private Function PrRename(name As String) As String
        For Each N As String In RenameList
            If name.StartsWith(N) Then Return N
        Next
        Return name
    End Function

    Private Function GetSpan(timeStr As String) As TimeSpan
        Dim h_m() As String = timeStr.Split(":"c)
        Dim ts As TimeSpan = New TimeSpan(Integer.Parse(h_m(0)), Integer.Parse(h_m(1)), 0)

        If ts.TotalHours < 6 Then
            ts += New TimeSpan(24, 0, 0)
        End If

        Return ts
    End Function

    Private Sub EditXml(ByVal fn As String)
        If Not File.Exists(fn) Then Return
        Dim sDate As Date = New FileInfo(fn).LastWriteTime

        Hide()
        Dim fnCopy As String = fn.Replace(".xml", "-1.xml")
        File.Copy(fn, fnCopy, True)

        Dim newProc As Process
        newProc = Process.Start("C:\WINDOWS\NOTEPAD.EXE ", fn)
        newProc.WaitForExit()

        If New FileInfo(fn).LastWriteTime > sDate Then
            Dim i As Integer = cName.SelectedIndex
            LoadSetup()
            If i > cName.Items.Count - 1 Then i = cName.Items.Count - 1
            cName.SelectedIndex = i
        End If

        Show()
    End Sub

    Private Sub Distribute()
        For Each fn As String In lFiles
            If File.Exists(fn) Then
                'Try
                '    File.Copy(fn, fn.Replace(eDir, fDir), True)
                'Catch ex As Exception

                'End Try
                For Each dn As String In dDir
                    Try
                        File.Copy(fn, fn.Replace(eDir, dn), True)
                    Catch ex As Exception

                    End Try
                Next
            End If
        Next
    End Sub

    Private Sub Archive()
        For Each fi As FileInfo In New DirectoryInfo(eDir).GetFiles
            Try
                If Char.IsDigit(fi.Name.First) Then
                    Dim d As Date = Date.ParseExact(fi.Name.Substring(0, 6), "yyMMdd", CultureInfo.InvariantCulture)
                    If (Now - d).TotalHours > 48 Then
                        Dim fn As String = fi.FullName.Replace(eDir, bDir)
                        If File.Exists(fn) Then File.Delete(fn)
                        fi.MoveTo(fn)
                        'fn = fn.Replace(bDir, fDir)
                        'If File.Exists(fn) Then File.Delete(fn)
                        For Each dn As String In dDir
                            Try
                                Dim f As String = fn.Replace(bDir, dn)
                                If File.Exists(f) Then File.Delete(f)
                            Catch ex As Exception

                            End Try
                        Next
                    End If
                End If

            Catch ex As Exception

            End Try
        Next
    End Sub

    Private Sub InitDest(ByVal fn As String)
        If Not File.Exists(fn) Then Return
        Dim initDoc = XDocument.Load(fn)

        Dim voDir As String = initDoc.Root.Element("vo").Value ' "Eetrigraafika\Kanal2\16-sk\voiceover"
        If Not voDir.EndsWith("\", StringComparison.Ordinal) Then voDir += "\"
        For Each el As XElement In initDoc.Root.Elements("dest")
            Dim d = el.Value
            If Not d.EndsWith("\", StringComparison.Ordinal) Then d += "\"
            dDir.Add(d + voDir)
        Next
    End Sub

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
                If Not watcher Is Nothing Then watcher.Dispose()
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region

    'Public Sub Dispose() Implements IDisposable.Dispose
    '    If Not watcher Is Nothing Then watcher.Dispose()
    'End Sub
End Class
