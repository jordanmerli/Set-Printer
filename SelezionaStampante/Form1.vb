Imports System.Drawing.Printing
Public Class Form1
    Dim strPrinterName As String
    Private Sub ContextMenuStrip1_ItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ContextMenuStrip1.ItemClicked
        'MsgBox(e.ClickedItem.Text)         'se tolgo l'apostrofo compare una messagebox con scritto l'elemento cliccato
        If e.ClickedItem.Text.Trim <> "CHIUDI" Then     'se clicco qualcosa che non sia Chiudi
            strPrinterName = e.ClickedItem.Text         'imposto la stampante selezionata
            Call SetDefaultSystemPrinter(strPrinterName)   'e chiamo la funzione SetDefaultSystemPrinter(strPrinterName) che ho creato sotto
        Else
            Call EsciProgramma()            'altrimenti chiamo la funzione EsciProgramma
        End If
    End Sub
    Private Sub ContextMenuStrip1_Opening(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening
        ContextMenuStrip1.Items.Clear()
        For Each strPrinterName As String In PrinterSettings.InstalledPrinters      'per ogni stampante installata
            Dim itm As New ToolStripMenuItem 'ToolStripButton
            itm.Text = strPrinterName
            itm.DisplayStyle = ToolStripItemDisplayStyle.Text
            ContextMenuStrip1.Items.Add(itm)
            Dim PD As New PrintDocument, StampPred As String
            StampPred = PD.DefaultPageSettings.PrinterSettings.PrinterName
            If StampPred = itm.Text Then
                itm.Checked = True
                itm.CheckState = CheckState.Checked
            End If
        Next
        Dim itm2 As New ToolStripMenuItem 'ToolStripButton
        itm2.Text = ""
        itm2.DisplayStyle = ToolStripItemDisplayStyle.Text
        ContextMenuStrip1.Items.Add(itm2)
        Dim itm3 As New ToolStripMenuItem 'ToolStripButton
        itm3.Text = "CHIUDI"
        itm3.DisplayStyle = ToolStripItemDisplayStyle.Text
        ContextMenuStrip1.Items.Add(itm3.Text)
    End Sub
    Sub EsciProgramma()         'funzione esci programma
        Dim jordan
        jordan = MsgBox("Sei sicuro di voler uscire?", MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Esci!")
        If jordan = vbYes Then
            End
        End If
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Me.Hide()           'all'avvio nascondo il form
    End Sub
    'Private Sub NotifyIcon1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.Click
    'Dim Position As System.Drawing.Point
    'Position.X = MousePosition.X
    'Position.Y = MousePosition.Y
    'ContextMenuStrip1.Show(Position.X, Position.Y)
    'End Sub
    '
    'Private Sub NotifyIcon1_MouseDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
    'Dim Position As System.Drawing.Point
    'Position.X = MousePosition.X
    'Position.Y = MousePosition.Y
    'ContextMenuStrip1.Show(Position.X, Position.Y)
    'End Sub
    Private Function SetDefaultSystemPrinter(ByVal strPrinterName As String) As Boolean         'funzione imposta stampante predefinita
        ' Esempio: SetDefaultSystemPrinter("Lexmark T620") - local printer
        ' Esempio: SetDefaultSystemPrinter("\\Server01\Lexmark T522") - network printer
        Dim strOldPrinter As String
        Dim WshNetwork As Object
        Dim pd As New PrintDocument
        ' Set the default system printer
        Try
            strOldPrinter = pd.PrinterSettings.PrinterName '/ Get the system default printer
            WshNetwork = Microsoft.VisualBasic.CreateObject("WScript.Network")
            WshNetwork.SetDefaultPrinter(strPrinterName)
            pd.PrinterSettings.PrinterName = strPrinterName '/ Specify the printer to use.
            If pd.PrinterSettings.IsValid Then '/ Check that the printer exists
                Return True
            Else
                'MessageBox.Show("Printer <" & strPrinterName & "> is invalid.")
                WshNetwork.SetDefaultPrinter(strOldPrinter)
                Return False
            End If
        Catch exptd As Exception
            WshNetwork.SetDefaultPrinter(strOldPrinter)
            Return False
        Finally
            WshNetwork = Nothing
            pd = Nothing
        End Try
    End Function '/ SetDefaultSystemPrinter
End Class
