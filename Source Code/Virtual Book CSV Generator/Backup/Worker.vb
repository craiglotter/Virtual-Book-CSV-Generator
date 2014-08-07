Imports System.IO
Imports System.Text


Public Class Worker

    Inherits System.ComponentModel.Component

    ' Declares the variables you will use to hold your thread objects.

    Public WorkerThread As System.Threading.Thread

    Public filequeue1 As String
    Private filecount As Long
    Private blankcount As Long
    Private fullcount As Long
    Private filereader As StreamReader
    Private filewriter As StreamWriter

    Public Event WorkerFileProcessing(ByVal filename As String, ByVal queue As Integer)
    Public Event WorkerStatusMessage(ByVal message As String, ByVal statustag As Integer)
    Public Event WorkerError(ByVal Message As Exception)
    Public Event WorkerFileCount(ByVal Result As Long, ByVal count As Integer)
    Public Event WorkerComplete(ByVal queue As Integer)




#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)

    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        
    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub Error_Handler(ByVal message As Exception)
        Try
            If (Not WorkerThread.ThreadState.ToString.IndexOf("Aborted") > -1) And (Not WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1) Then
                RaiseEvent WorkerError(message)
            End If
        Catch ex As Exception
            MsgBox("An error occurred in Virtual Book CSV Generator's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub



    Public Sub ChooseThreads(ByVal threadNumber As Integer)
        Try
            ' Determines which thread to start based on the value it receives.
            Select Case threadNumber
                Case 1
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerFileCount_Routine)
                    WorkerThread.Start()
                Case 2
                    filewriter.Close()
                    filereader.Close()
            End Select
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub



   

    Private Sub WorkerFileCount_Routine()
        RaiseEvent WorkerStatusMessage("Generating CSV File", 1)
        filecount = 0
        fullcount = 0
        blankcount = 0

        RaiseEvent WorkerFileCount(filecount, 1)
        RaiseEvent WorkerFileCount(fullcount, 2)
        RaiseEvent WorkerFileCount(blankcount, 3)

        Try
            FileCountRunner(filequeue1)
        Catch ex As Exception
            Error_Handler(ex)
        End Try

        RaiseEvent WorkerFileCount(filecount, 1)
        RaiseEvent WorkerFileCount(fullcount, 2)
        RaiseEvent WorkerFileCount(blankcount, 3)

        RaiseEvent WorkerStatusMessage("CSV File Generated", 1)
        RaiseEvent WorkerComplete(0)
    End Sub

    Private Sub FileCountRunner(ByVal filename As String)
        Try
            filereader = New StreamReader(filename, True)
            filewriter = New StreamWriter(filename & ".csv", False)
            Dim line As String
            line = filereader.ReadLine
            RaiseEvent WorkerStatusMessage("Examining: " & filename, 2)
            RaiseEvent WorkerStatusMessage("Generating: " & filename & ".csv", 3)

            While Not IsNothing(line) = True
                If (Not WorkerThread.ThreadState.ToString.IndexOf("Aborted") > -1) And (Not WorkerThread.ThreadState.ToString.IndexOf("AbortRequested") > -1) Then
                    filecount = filecount + 1
                    RaiseEvent WorkerFileCount(filecount, 1)
                    If line.Length > 0 And Not line = "" Then
                        fullcount = fullcount + 1
                        RaiseEvent WorkerFileCount(fullcount, 2)

                        Dim outline As String
                        outline = line
                        outline = line.Replace(vbTab, ",")
                        Dim spoutline = outline.Split(",")

                        Dim firstname As String
                        firstname = Trim(spoutline(2))
                        Dim st As Integer
                        Dim spfirstname() As String = firstname.Split(" ")
                        firstname = ""
                        For st = 0 To spfirstname.Length - 2
                            firstname = firstname & spfirstname(st) & " "
                        Next st
                        firstname = firstname.Trim()

                        outline = Trim(spoutline(0)) & "," & firstname & "," & Trim(spoutline(1)) & ", ," & Trim(spoutline(0)) & "@mail.uct.ac.za,cache2.uct.ac.za,8080, , , , , , "
                        filewriter.WriteLine(outline)

                    Else
                        blankcount = blankcount + 1
                        RaiseEvent WorkerFileCount(blankcount, 3)
                    End If

                    line = filereader.ReadLine
                Else
                    Exit While
                End If
            End While

            filereader.Close()
            filewriter.Close()

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub 'ProcessDirectory

End Class
