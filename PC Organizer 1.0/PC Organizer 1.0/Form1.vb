Imports Microsoft.Win32

Public Class Form1
    Private intTabs% = 1
    Private mclrMain As Color
    Private myReg As Registry
    Private aTyp() As String

    Public Enum ICON_SIZE
        small = 0
        large = 1
    End Enum

    Private IconSize = ICON_SIZE.small

    Private Sub Page_Click(ByVal obj As Object, ByVal e As EventArgs)
        Dim p As TabPage = obj
        Debug.Print(p.Text)
        p.AllowDrop = True
    End Sub

    Public dd As New DragDrop
    Public fn$ = ""
    Public fnLast$ = ""

    Public Structure myIcon
        Public i As Icon
        Public f As String
        Public p As String
    End Structure

    Private Icons As New List(Of myIcon)

    Private Sub DisposeIcons()
        For Each i As myIcon In Icons
            i.i.Dispose()
        Next
        Icons.Clear()
    End Sub

    Private Sub Listview_Click(ByVal obj As Object, ByVal e As EventArgs)
        Dim lv As ListView = obj
        Dim s$ = lv.SelectedItems(0).ToString
        Dim t$ = s.Replace("ListViewItem: {", "").Replace("}", "")
        MsgBox(t)
    End Sub

    Private Sub SetIcon(ByVal pathname As String, ByRef p As TabPage)
        DisposeIcons()
        Dim a$()
        'Get the icon for this dropped file
        Dim m As myIcon
        m.i = IconHelper.GetIconFrom(pathname, IconSize.Small, False)
        m.p = pathname
        Erase a
        a = m.p.Split("\")
        m.f = a(UBound(a))
        Icons.Add(m)
        m.i = IconHelper.GetIconFrom(pathname, IconSize.Large, False)
        m.p = pathname
        m.f = a(UBound(a))
        Icons.Add(m)

        'Figure out the correct icon for the icon size setting
        m = Icons(IconSize)
        Dim n1% = ImageList1.Images.Count
        Dim n2% = ImageList2.Images.Count

        Select Case IconSize
            Case ICON_SIZE.large
                ImageList2.Images.Add(m.f, m.i)
                p.ImageKey = n2
            Case Else
                ImageList1.Images.Add(m.f, m.i)
                p.ImageKey = n1
        End Select

        p.SuspendLayout()

        Dim lv As ListView
        Dim lb As ListBox
        If p.Controls.Count < 1 Then
            lv = New ListView
            lb = New ListBox
            lb.Visible = False
            p.Controls.Add(lv)
            p.Controls.Add(lb)
            AddHandler lv.Click, AddressOf Listview_Click
        Else
            lv = p.Controls(0)
            lb = p.Controls(1)
        End If

        lv.BackColor = p.BackColor
        lv.Dock = DockStyle.Fill

        Dim li As New ListViewItem

        Select Case IconSize
            Case ICON_SIZE.large
                lv.View = View.LargeIcon
                lv.LargeImageList = ImageList2
                li.ImageIndex = n2
            Case Else
                lv.View = View.SmallIcon
                lv.SmallImageList = ImageList1
                li.ImageIndex = n1
        End Select
        Select Case IconSize
            Case ICON_SIZE.large
                lv.Items.Add(p.ImageKey, m.f, li.ImageIndex - 1)
            Case Else
                lv.Items.Add(p.ImageKey, m.p, li.ImageIndex - 1)
        End Select
        lb.Items.Add(m.p)
        p.ResumeLayout()
        Me.Refresh()
    End Sub

    Private Sub Page_DragEnter(ByVal obj As Object, ByVal e As System.Windows.Forms.DragEventArgs)
        Debug.Print("DragEnter")
    End Sub

    Private Sub AddFile(ByVal strFilePath As String, ByVal tp As TabPage)
        Me.SetIcon(strFilePath, tp)
    End Sub

    Private Sub Page_DragDrop(ByVal obj As Object, ByVal e As System.Windows.Forms.DragEventArgs)
        Dim p As TabPage = obj, sLast$ = ""
        If (e.Data.GetDataPresent(DataFormats.FileDrop)) Then
            Dim strFiles() As String = e.Data.GetData(DataFormats.FileDrop)
            Dim intCount As Integer
            For intCount = 0 To strFiles.Length - 1
                If strFiles(intCount) <> sLast Then
                    AddFile(strFiles(intCount), p)
                    sLast = strFiles(intCount)
                End If
            Next
        End If
    End Sub

    Private Sub Page_DragLeave(ByVal obj As Object, ByVal e As System.Windows.Forms.DragEventArgs)
        Debug.Print("DragLeave")
    End Sub

    Private Sub Page_DragOver(ByVal obj As Object, ByVal e As System.Windows.Forms.DragEventArgs)
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub MakeTabs(ByVal aT() As String, ByVal aC() As String, ByVal aF() As String, ByRef aTyp() As String)
        Dim n% = 0, tpi% = 0, bDone As Boolean, j% = 0, u% = 0, sTypes$ = ""
        For i = 0 To aC.Length - 1
            If aT(i).Length > 0 Then
                If Not aT(i) Like "Main*" Then
                    Me.TabControl1.TabPages.Add(aT(i))
                End If
                tpi = Me.TabControl1.TabPages.Count - 1
                Me.TabControl1.TabPages(i).AllowDrop = True
                AddHandler Me.TabControl1.TabPages(i).Click, AddressOf Page_Click
                AddHandler Me.TabControl1.TabPages(i).DragDrop, AddressOf Page_DragDrop
                AddHandler Me.TabControl1.TabPages(i).DragEnter, AddressOf Page_DragEnter
                AddHandler Me.TabControl1.TabPages(i).DragLeave, AddressOf Page_DragLeave
                AddHandler Me.TabControl1.TabPages(i).DragOver, AddressOf Page_DragOver
                If aC(i).Length > 0 Then
                    Me.TabControl1.TabPages(i).BackColor = Color.FromArgb(aC(i))
                    Me.TabControl1.TabPages(i).UseVisualStyleBackColor = False
                End If

                'Get file types for this tab
                Dim bFound As Boolean = False
                For u = LBound(aTyp) To UBound(aTyp)
                    If aTyp(u).Contains("Tab") AndAlso aTyp(u).Contains(aT(i)) Then
                        sTypes = "Tab=" & aT(i) & "|"
                        For v = u + 1 To aTyp.Length - 1
                            sTypes &= aTyp(v) & "|"
                            u += 1
                            bFound = True
                            Exit For
                        Next
                    End If
                    If bFound Then
                        bFound = False
                        Exit For
                    End If
                Next

                'Get icons for this tab
                If aF(j).Contains("Tab") AndAlso aF(j).Contains(aT(i)) Then
                    For j = n To aF.Length - 1
                        j += 1
                        For k = j To aF.Length - 1
                            If Not aF(k).Contains("Tab") AndAlso aF(k).Length > 0 Then
                                If System.IO.File.Exists(aF(k)) Then
                                    Dim s$ = System.IO.Path.GetExtension(aF(k))
                                    s = s.Replace(".", "")
                                    If sTypes.Contains("Tab=" & aT(i) & "|" & s) Then
                                        Me.AddFile(aF(k), Me.TabControl1.TabPages(tpi))
                                    Else
                                        'Extension is not allowed for this tab, do you want to add it?
                                        Dim ans As MsgBoxResult = _
                                            MsgBox("Add this extension (" & s & ") for the tab (" & _
                                                   Me.TabControl1.TabPages(tpi).Text & ")?", _
                                                   MsgBoxStyle.YesNo, "Allowable File Extensions")
                                        If ans = MsgBoxResult.Yes Then
                                            Dim y% = -1
                                            Me.AddFile(aF(k), Me.TabControl1.TabPages(tpi))
                                            For x = LBound(aTyp) To UBound(aTyp)
                                                If aTyp(x).Contains(Me.TabControl1.TabPages(tpi).Text) Then
                                                    y = x
                                                    Exit For
                                                End If
                                            Next
                                            bFound = False
                                            If y < 0 Then
                                                For x = LBound(aTyp) To UBound(aTyp)
                                                    If aTyp(x) Is Nothing OrElse aTyp(x).Length < 1 Then
                                                        aTyp(x) = "Tab=" & Me.TabControl1.TabPages(tpi).Text
                                                        If x + 2 > aTyp.Count Then ReDim Preserve aTyp(aTyp.Count + 1)
                                                        aTyp(x + 1) = s
                                                        bFound = True
                                                        Exit For
                                                    End If
                                                Next
                                                If Not bFound Then
                                                    ReDim Preserve aTyp(aTyp.Count + 1)
                                                    For q = LBound(aTyp) To UBound(aTyp)
                                                        If aTyp(q) Is Nothing OrElse aTyp(q) < 1 Then
                                                            aTyp(q) = "Tab=" & Me.TabControl1.TabPages(tpi).Text
                                                            If q + 1 > aTyp.Count Then ReDim Preserve aTyp(aTyp.Count + 1)
                                                            aTyp(q + 1) = s
                                                        End If
                                                    Next
                                                End If
                                            Else
                                                If y + 1 > aTyp.Count Then
                                                    ReDim Preserve aTyp(aTyp.Count + 1)
                                                    aTyp(y + 1) = s
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            ElseIf Not bDone Then
                                n = k
                                bDone = True
                                Exit For
                            End If
                        Next
                        If bDone Then
                            bDone = False
                            j += 1
                            Exit For
                        End If
                    Next
                End If


            End If
            If aC(i).Length > 0 Then
                Me.TabControl1.TabPages(i).BackColor = Color.FromArgb(aC(i))
                Me.TabControl1.TabPages(i).UseVisualStyleBackColor = False
            End If
        Next
    End Sub

    Private Sub GetTabInfo()
        Dim sT$ = myReg.GetKey(, , "Tabs")
        Dim sC$ = myReg.GetKey(, , "TabColors")
        Dim sF$ = myReg.GetKey(, , "TabFiles")
        Dim sTyp$ = myReg.GetKey(, , "TabExt")
        Dim aT = sT.Split("|")
        Dim aC = sC.Split("|")
        Dim aF = sF.Split("|")
        Erase aTyp
        aTyp = sTyp.Split("|")
        Me.MakeTabs(aT, aC, aF, aTyp)
    End Sub

    Private Sub SaveTabInfo()
        Dim s$ = ""
        For Each t As TabPage In Me.TabControl1.TabPages
            s &= t.Text & "|"
        Next
        myReg.SetKey(s, , , "Tabs")
        s = ""
        For Each t As TabPage In Me.TabControl1.TabPages
            s &= t.BackColor.ToArgb.ToString & "|"
        Next
        myReg.SetKey(s, , , "TabColors")
        s = ""
        For Each t As TabPage In Me.TabControl1.TabPages
            'Dim lv As ListView = t.Controls(1)
            If t.Controls.Count > 0 Then
                Dim lv As ListView = t.Controls(0)
                Dim lb As ListBox = t.Controls(1)
                s &= "Tab=" & t.Text & "|"
                For Each p As String In lb.Items
                    s &= p & "|"
                Next
            End If
        Next
        myReg.SetKey(s, , , "TabFiles")
        s = ""
        For i = LBound(aTyp) To UBound(aTyp)
            If aTyp(i) IsNot Nothing Then
                s &= aTyp(i) & "|"
            Else
                Exit For
            End If
        Next
        myReg.SetKey(s, , , "TabExt")
    End Sub

    Private Sub GetFormInfo()
        Dim s$ = myReg.GetKey(, , "FormInfo")
        Dim a$() = s.Split("|")
        Dim x As New Point
        If a.Length > 3 AndAlso (a(0) <> "0" OrElse a(1) <> "0" OrElse a(2) <> "") Then
            x.X = a(0)
            x.Y = a(1)
            Me.Location = x
            Me.Width = a(2)
            Me.Height = a(3)
        End If
        Me.GetTabInfo()
    End Sub

    Private Sub SaveFormInfo()
        Dim s$ = Me.Location.X.ToString & "|" & Me.Location.Y.ToString & "|" & Me.Width & "|" & Me.Height & "|"
        myReg.SetKey(s, , , "FormInfo")
        Me.SaveTabInfo()
    End Sub

    Private Sub AddToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles AddToolStripMenuItem.Click
        Dim a$() = {"Main"}, i% = 1
        For Each t As TabPage In TabControl1.TabPages
            If t.Text <> "Main" Then
                ReDim Preserve a(i)
                If t.Text Like "Tab*" Then
                    a(i) = t.Text.Replace("Tab", "")
                End If
                i += 1
            End If
        Next
        i -= 1
        Dim n% = 1
        If i > 0 Then
            For j = 1 To i
                If IsNumeric(a(j)) Then
                    n = a(j)
                End If
                n += 1
            Next
        Else
            n = i + 1
        End If
        intTabs = n + 1
        Dim s$ = InputBox("Tab Name", "Add Tab", "Tab" & intTabs - 1)
        If s.Length > 0 Then
            Me.TabControl1.TabPages.Add(intTabs, s)
            Dim arb% = mclrMain.ToArgb + intTabs * 100
            Me.TabControl1.TabPages(intTabs - 1).BackColor = Color.FromArgb(arb)
            'setval()
            mclrMain = Color.FromArgb(arb)
        End If
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        If TabControl1.SelectedIndex > 0 Then
            TabControl1.TabPages.RemoveAt(TabControl1.SelectedIndex)
            intTabs -= 1
        End If
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Me.SaveFormInfo()
        myReg.Dispose()
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        mclrMain = Me.TabControl1.TabPages(Me.TabControl1.TabPages.Count - 1).BackColor
        myReg = New Registry(My.Computer.Registry.CurrentUser.ToString, "PC Organizer", "LastRun", Now.ToString)
        Me.GetFormInfo()
    End Sub

    Private Sub RenameToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RenameToolStripMenuItem.Click
        Dim n% = Me.TabControl1.SelectedIndex
        Dim s$ = Me.TabControl1.SelectedTab.Text
        If Not s Like "Main*" Then
            Dim t$ = InputBox("Rename Tab?", "Tab Rename", s)
            If t <> s AndAlso t.Length > 0 Then
                Me.TabControl1.SelectedTab.Text = t
            End If
        Else
            MsgBox("You cannot rename the 'MAIN' tab", MsgBoxStyle.Exclamation, "No Can Do!")
        End If
    End Sub


    Private Sub TabControl1_Selecting(ByVal sender As Object, ByVal e As System.Windows.Forms.TabControlCancelEventArgs) Handles TabControl1.Selecting
        RenameToolStripMenuItem.Enabled = Me.TabControl1.SelectedIndex > 0
        DeleteToolStripMenuItem.Enabled = Me.TabControl1.SelectedIndex > 0
    End Sub

End Class
