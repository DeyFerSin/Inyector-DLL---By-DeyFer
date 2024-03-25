Imports System.Runtime.InteropServices
Public Class Form1
#Region "Variables y referencias"
    Public Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Integer, ByVal lpAddress As Integer, ByVal dwSize As Integer, ByVal flAllocationType As Integer, ByVal flProtect As Integer) As Integer
    Public Const MEM_COMMIT = 4096, PAGE_EXECUTE_READWRITE = &H40
    Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Integer, ByVal lpBaseAddress As Integer, ByVal lpBuffer As Byte(), ByVal nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Integer
    Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Integer, ByVal lpProcName As String) As Integer
    Private Declare Function GetModuleHandle Lib "Kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Integer
    Public Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Integer, ByVal lpThreadAttributes As Integer, ByVal dwStackSize As Integer, ByVal lpStartAddress As Integer, ByVal lpParameter As Integer, ByVal dwCreationFlags As Integer, ByRef lpThreadId As Integer) As Integer
    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function AllocConsole() As Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Shared Function FreeConsole() As Boolean
    End Function
#End Region
#Region "Inicio de la aplicacion"
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'AllocConsole()
        Console.WriteLine("Inicio de la app")
    End Sub
#End Region
#Region "Seccion de Timer´s"
    Private Sub InjectTimer_Tick(sender As Object, e As EventArgs) Handles InjectTimer.Tick
        If IO.File.Exists(OpenFileDialog1.FileName) Then
            Dim TargetProcess As Process() = Process.GetProcessesByName(ProcessTextBox.Text)
            If TargetProcess.Length = 0 Then
                StatusLabel.ForeColor = Color.Red
                StatusLabel.Text = ("Esperando a " + ProcessTextBox.Text + ".exe")
            Else
                InjectTimer.Stop()
                DelayTimer.Start()
            End If
        Else
        End If
    End Sub
    Private Sub DelayTimer_Tick(sender As Object, e As EventArgs) Handles DelayTimer.Tick
        If DelayNumeric.Value = 0 Then
            Inject(ProcessTextBox.Text)
            DelayTimer.Enabled = False
            StatusLabel.ForeColor = Color.Lime
            StatusLabel.Text = "La DLL a sido inyectado con exito!"
            For i = 0 To (DllListBox.Items.Count + -1)
                If CloseCheckBox.Checked = True Then End
            Next i
        Else
            DelayNumeric.Value = DelayNumeric.Value - 1
        End If
    End Sub
    Private Sub TimerData_Tick(sender As Object, e As EventArgs) Handles TimerData.Tick
        CarbonFiberTheme1.Text = "Inyector DLL - By DeyFer"
    End Sub
#End Region
#Region "Inyeccion"
    Private Sub Inject(processName As String)
        Dim DllPath As String = RutaCompletaDLL
        If (Process.GetProcessesByName(ProcessTextBox.Text).Length = 0) Then StatusLabel.Text = "No es posible encontrar el proceso " + processName

        Dim TargetHandle As IntPtr = Process.GetProcessesByName(ProcessTextBox.Text)(0).Handle
        If (TargetHandle.Equals(IntPtr.Zero)) Then
            StatusLabel.Text = "El proceso " + processName + " ha fracasado."
            Exit Sub
        End If

        Dim GetAdrOfLLBA As IntPtr = GetProcAddress(GetModuleHandle("Kernel32"), "LoadLibraryA")
        If (GetAdrOfLLBA.Equals(IntPtr.Zero)) Then
            StatusLabel.Text = "No se puede obtener la direccion por defecto de la funcion de carga API"
            Exit Sub
        End If

        Dim OperaChar As Byte() = System.Text.Encoding.Default.GetBytes(DllPath)
        Dim DllMemPathAdr = VirtualAllocEx(TargetHandle, 0&, &H64, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
        If (DllMemPathAdr.Equals(IntPtr.Zero)) Then
            StatusLabel.Text = "Proceso " + processName + "error de solicitud."
            Exit Sub
        End If

        If (WriteProcessMemory(TargetHandle, DllMemPathAdr, OperaChar, OperaChar.Length, 0) = False) Then
            StatusLabel.Text = "Proceso " + processName + "error al escribir la memoria!"
            Exit Sub
        End If
        CreateRemoteThread(TargetHandle, 0, 0, GetAdrOfLLBA, DllMemPathAdr, 0, 0)
        StatusLabel.Text = "Proceso " + processName + "se ha realizado la inyeccion correctamente"
    End Sub
    'Boton de inyeccion
    Private Sub InjectButton_Click(sender As Object, e As EventArgs) Handles InjectButton.Click
        If IO.File.Exists(OpenFileDialog1.FileName) Then
            Dim TargetProcess As Process() = Process.GetProcessesByName(ProcessTextBox.Text)
            If TargetProcess.Length = 0 Then
                StatusLabel.ForeColor = Color.Red
                StatusLabel.Text = ("Esperando a " + ProcessTextBox.Text + ".exe")
            Else
                InjectTimer.Stop()
                DelayTimer.Start()
            End If
        Else
        End If
    End Sub
    'Opciones de inyeccion

    'Automatica
    Private Sub AutoRadioButton_CheckedChanged(sender As Object) Handles AutoRadioButton.CheckedChanged
        InjectButton.Enabled = False
        InjectTimer.Enabled = True
    End Sub
    'Manual
    Private Sub ManualRadioButton_CheckedChanged(sender As Object) Handles ManualRadioButton.CheckedChanged
        InjectButton.Enabled = True
        InjectTimer.Enabled = False
    End Sub
#End Region
    'Boton de busqueda de las DLL´s
    Private Sub BrowseButton_Click(sender As Object, e As EventArgs) Handles BrowseButton.Click
        ' Establecer las propiedades del OpenFileDialog
        OpenFileDialog1.Filter = "DLL (*.dll) |*.dll" ' Filtrar por archivos JPG
        OpenFileDialog1.FilterIndex = 1 ' Establecer el índice del filtro predeterminado
        OpenFileDialog1.RestoreDirectory = True ' Restaurar el directorio original después de seleccionar un archivo

        OpenFileDialog1.ShowDialog()
    End Sub
    'Boton de eliminar una sola DLL (La que seleccionas)
    Private Sub RemoveButton_Click(sender As Object, e As EventArgs) Handles RemoveButton.Click
        For i As Integer = (DllListBox.SelectedItems.Count - 1) To 0 Step -1
            DllListBox.Items.Remove(DllListBox.SelectedItems(i))
        Next
    End Sub
    'Boton para eliminar todas las DLL que se encuentran en el listbox
    Private Sub ClearAllButton_Click(sender As Object, e As EventArgs) Handles ClearAllButton.Click
        DllListBox.Items.Clear()
    End Sub
    'Dialogo para la busqueda de las DLL
    Dim RutaCompletaDLL As String
    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        Dim FileName As String
        FileName = OpenFileDialog1.FileName.Substring(OpenFileDialog1.FileName.LastIndexOf("\"))
        RutaCompletaDLL = OpenFileDialog1.FileName

        Dim DllFileName As String = FileName.Replace("\", "")
        Console.WriteLine("Ruta completa:" & RutaCompletaDLL)
        Console.WriteLine("Dll:" & DllFileName)
        DllListBox.Items.Add(DllFileName)
    End Sub
End Class
