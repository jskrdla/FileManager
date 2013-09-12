Imports System.IO

Public Class Form1

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        SearchDirectory(New DirectoryInfo("D:\Temp\Downloads\"))
    End Sub

    Private Sub SearchDirectory(ByVal di As DirectoryInfo)
        For Each item As DirectoryInfo In di.GetDirectories
            SearchDirectory(item)
        Next
        For Each fi As FileInfo In di.GetFiles("*.mp3")
            MoveFile(fi)
        Next
        For Each fi As FileInfo In di.GetFiles("*.mp4")
            MoveFile(fi)
        Next
        For Each fi As FileInfo In di.GetFiles("*.m4v")
            MoveFile(fi)
        Next
    End Sub

    Private Sub MoveFile(ByVal fi As FileInfo)
        Dim sName As String = fi.Name.Replace(fi.Extension, "")
        Dim iCounter As Integer = 0
        While File.Exists("D:\Temp\" & sName & fi.Extension)
            iCounter += 1
            sName = fi.Name & iCounter.ToString
        End While
        fi.MoveTo("D:\Temp\" & sName & fi.Extension)
        ListBox1.Items.Add("Moved " & fi.Name & fi.Extension & " to Temp\" & sName & fi.Extension)
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Timer1.Enabled = False
        SearchDirectory(New DirectoryInfo("D:\Temp\Downloads\"))
        Timer1.Enabled = True
    End Sub


    Private Sub Form1_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Me.WindowState <> FormWindowState.Minimized Then
            Dim result As DialogResult = MessageBox.Show("Do you want to exit?", "Exit?", MessageBoxButtons.YesNoCancel)
            If result = Windows.Forms.DialogResult.No Then
                Me.WindowState = FormWindowState.Minimized
                Me.Visible = False
                e.Cancel = True
            ElseIf result = Windows.Forms.DialogResult.Cancel Then
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        Me.Visible = True
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub ShowToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ShowToolStripMenuItem.Click
        Me.Visible = True
        Me.WindowState = FormWindowState.Normal
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Timer1.Enabled = True
    End Sub
End Class
