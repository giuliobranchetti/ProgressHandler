Imports System.ComponentModel
Imports System.Windows.Forms

Public Class ProgressHandler
    Inherits BackgroundWorker

    Protected _pb As ProgressBar
    Protected _label As Label
    Protected _tsLabel As ToolStripLabel

    Protected _max As Integer
    Protected _progress As Integer

    Public Event StartWork(sender As Object, e As ProgressChangedEventArgs)

    Public Sub New()
        MyBase.New()

        _pb = Nothing
        _label = Nothing
        _tsLabel = Nothing

        _max = 0
        _progress = 0

        WorkerReportsProgress = True

        Control.CheckForIllegalCrossThreadCalls = False

        AddHandler StartWork, AddressOf HandleStartWork
        AddHandler ProgressChanged, AddressOf HandleProgressChanged
        AddHandler RunWorkerCompleted, AddressOf HandleRunWorkerCompleted
    End Sub

    Public ReadOnly Property Progress
        Get
            Return _progress
        End Get
    End Property

    Public ReadOnly Property Max
        Get
            Return _max
        End Get
    End Property

    Public Function WithProgressBar(ByRef pb As ProgressBar) As ProgressHandler
        _pb = pb
        Return Me
    End Function

    Public Function WithProgressBar(ByRef tspb As ToolStripProgressBar) As ProgressHandler
        Return WithProgressBar(tspb.ProgressBar)
    End Function

    Public Function WithLabel(ByRef label As Label) As ProgressHandler
        _label = label
        Return Me
    End Function

    Public Function WithLabel(ByRef label As ToolStripLabel) As ProgressHandler
        _tsLabel = label
        Return Me
    End Function

    Public Sub Run()
        _progress = 0

        RaiseEvent StartWork(Me, New ProgressChangedEventArgs(_progress, ""))

        RunWorkerAsync()
    End Sub

    Public Sub Run(argument As Object)
        _progress = 0

        RaiseEvent StartWork(Me, New ProgressChangedEventArgs(_progress, ""))

        RunWorkerAsync(argument)
    End Sub

    Public Sub SetMax(max As Integer)
        _max = max
    End Sub

    Public Sub Increment(Optional amount As Integer = 1)
        _progress += amount

        ReportProgress(_progress)
    End Sub

    Public Sub Increment(userState As Object, Optional amount As Integer = 1)
        _progress += amount

        ReportProgress(_progress, userState)
    End Sub

    Private Sub HandleStartWork(sender As Object, e As ProgressChangedEventArgs)
        If Not IsNothing(_pb) Then
            ' la progress bar lavora sempre in percentuale
            _pb.Maximum = 100
            _pb.Value = 0
            _pb.Visible = True
        End If

        If Not IsNothing(_label) Then
            ' mentre la label mostra il valore effettivo
            If e.UserState Is Nothing Then
                _label.Text = $"{_progress}/{_max}"
            Else
                _label.Text = e.UserState.ToString
            End If
        End If

        If Not IsNothing(_tsLabel) Then
            ' mentre la label mostra il valore effettivo
            If e.UserState Is Nothing Then
                _tsLabel.Text = $"{_progress}/{_max}"
            Else
                _tsLabel.Text = e.UserState.ToString
            End If
        End If
    End Sub

    Private Sub HandleProgressChanged(sender As Object, e As ProgressChangedEventArgs)
        If Not IsNothing(_pb) Then
            _pb.Value = e.ProgressPercentage * 100 / _max
            _pb.Refresh()
        End If

        If Not IsNothing(_label) Then
            If e.UserState Is Nothing Then
                _label.Text = $"{_progress}/{_max}"
            Else
                _label.Text = e.UserState.ToString
            End If
        End If

        If Not IsNothing(_tsLabel) Then
            If e.UserState Is Nothing Then
                _tsLabel.Text = $"{_progress}/{_max}"
            Else
                _tsLabel.Text = e.UserState.ToString
            End If
        End If
    End Sub

    Private Sub HandleRunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs)
        If Not IsNothing(_pb) Then
            _pb.Visible = False
        End If

        If Not IsNothing(_label) Then
            If e.Result Is Nothing Then
                _label.Text = $"{_max} record trovati"
            Else
                _label.Text = e.Result.ToString
            End If
        End If

        If Not IsNothing(_tsLabel) Then
            If e.Result Is Nothing Then
                _tsLabel.Text = $"{_max} record trovati"
            Else
                _tsLabel.Text = e.Result.ToString
            End If
        End If
    End Sub

End Class
