VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desconectado"
   ClientHeight    =   7050
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14790
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7050
   ScaleWidth      =   14790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConexao 
      Caption         =   "Conectar"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12000
      TabIndex        =   10
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSketch 
      Caption         =   "Sketch"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "Upload"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdCompile 
      Caption         =   "Compile"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox txtSketch 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   8295
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   7
      DTREnable       =   -1  'True
      RThreshold      =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1440
      Top             =   6360
   End
   Begin VB.ComboBox cboChrEspecial 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Form1.frx":169B2
      Left            =   9600
      List            =   "Form1.frx":169C2
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtSend 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1200
      TabIndex        =   3
      Top             =   600
      Width           =   8295
   End
   Begin VB.ComboBox cboBaudRate 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      ItemData        =   "Form1.frx":169FA
      Left            =   12000
      List            =   "Form1.frx":16A16
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.ComboBox cboCommPort 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "Form1.frx":16A4F
      Left            =   13440
      List            =   "Form1.frx":16A51
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtConsole 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   5895
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1080
      Width           =   14775
   End
   Begin VB.Menu mMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mSend 
         Caption         =   "Send"
      End
      Begin VB.Menu mClear 
         Caption         =   "Clear"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Shell
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Sleep
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Variável global
Dim scan As Boolean
Dim compileCmd As String

Private Sub Form_Load()
   Me.Caption = App.Title & "_v" & App.Major & "." & App.Minor & " by Dalçóquio Automação"
   writeConsole ("Desconectado !!!")
   
   Call scanCommPort
   
   cmdCompile.Enabled = True
   cmdUpload.Enabled = True
   txtSketch.Locked = True
   txtConsole.Locked = False
   
   cboBaudRate.Text = cboBaudRate.List(3) ' 9600
   cboChrEspecial.Text = cboChrEspecial.List(0) ' None
   
   ' Verifica se existe a pasta "C:\arduino-cli"
   Dim folderPath As String
   folderPath = "C:\arduino-cli"
   If Not Dir(folderPath, vbDirectory) <> "" Then
       MsgBox "Pasta " & folderPath & " não foi encontrada !!!", vbExclamation, "DALÇOQUIO AUTOMAÇÃO"
       End ' FECHA O APLICATIVO
   End If
   
End Sub

Private Sub cmdConexao_Click()
   If MSComm1.PortOpen = False Then
      'Conectar
      MSComm1.Settings = cboBaudRate.Text & "n,8,1"
      MSComm1.CommPort = Mid(cboCommPort.Text, 4, 6)
      MSComm1.PortOpen = True
      cmdConexao.Caption = "Desconectar"
      writeConsole ("Conectado na " & cboCommPort.Text & "," & MSComm1.Settings)
   Else
      'Desconectar
      MSComm1.PortOpen = False
      cmdConexao.Caption = "Conectar"
      writeConsole ("Desconectado !!!")
   End If

End Sub

Private Sub scanCommPort()
   scan = True
   cboCommPort.Enabled = False
   
   cboCommPort.Clear
   Dim i As Integer
   For i = 1 To 16 'Procura portas COM de 1 a 16
      MSComm1.CommPort = i
      On Error Resume Next 'ignora o tratamento de erro
      MSComm1.PortOpen = True 'tenta abrir a porta
      If Err.Number = 0 Then 'a porta está disponível
         cboCommPort.AddItem "COM" & i
         cboCommPort.ListIndex = 1
         MSComm1.PortOpen = False 'fecha a porta
      End If
      On Error GoTo 0 'ativa o tratamento de erro novamente
   Next i
   
   cboCommPort.Text = cboCommPort.List(0)
   cboCommPort.Enabled = True
   scan = False

End Sub

Private Sub cboBaudRate_Click()
      MSComm1.Settings = cboBaudRate.Text & "n,8,1"
      
      If MSComm1.PortOpen = True Then
         writeConsole ("Conectado na " & cboCommPort.Text & "," & MSComm1.Settings)
      End If
    
End Sub

Private Sub MSComm1_OnComm()
   On Error GoTo Erro
      Dim strData As String
      Do While MSComm1.InBufferCount > 0
          If strData = Empty Then Exit Do
          strData = MSComm1.Input(MSComm1.InBufferCount)
      Loop
      txtConsole.Text = txtConsole.Text + MSComm1.Input
      txtConsole.SelStart = Len(txtConsole.Text)
   Exit Sub
   
Erro:
   MsgBox Error, vbExclamation, "DALCOQUIO AUTOMAÇÃO"
   
End Sub

Private Sub cmdSend_Click()
   If MSComm1.PortOpen = True Then
      Call sendDado
      writeConsole ("Enviado com sucesso...")
   End If
   
End Sub

Private Sub txtSend_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And MSComm1.PortOpen = True Then
      Call sendDado
   End If

End Sub

Private Sub sendDado()
   If cboChrEspecial = "None" Then
      MSComm1.output = txtSend.Text
   ElseIf cboChrEspecial = "Nova Linha" Then
      MSComm1.output = txtSend.Text & vbLf 'vbNewLine
   ElseIf cboChrEspecial = "Retorno de Carro" Then
      MSComm1.output = txtSend.Text & vbCr
   ElseIf cboChrEspecial = "Ambos, NL e CR" Then
      MSComm1.output = txtSend.Text & vbCrLf
   End If

End Sub

Private Sub txtConsole_DblClick()
   mSend.Visible = False
   mClear.Visible = True
   PopupMenu mMenu
   mSend.Visible = True
   mClear.Visible = True

End Sub

Private Sub mClear_Click()
   txtConsole.Text = Empty
   
End Sub

Private Sub writeConsole(mensagem As String)
   txtConsole.Text = txtConsole.Text & "> " & mensagem & vbCrLf
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If MSComm1.PortOpen = True Then
      MSComm1.PortOpen = False
      writeConsole ("Desconectado !!!")
   End If
   writeConsole ("Fechando o sistema...")
   DoEvents
   Sleep (1000)
   End

End Sub

'////////////////////////////////////////////////////////////////////////////////////////////

Private Sub txtArduinoCLI_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       ' Define o caminho do arduino-cli
       Dim arduinoCliPath As String
       arduinoCliPath = "C:\arduino-cli.exe" ' Substitua pelo caminho correto do arduino-cli em seu sistema
   
       ' Define o comando para arduino-cli
       Dim commandCmd As String
       commandCmd = txtArduinoCLI.Text  ' comando user para arduino-cli
   
       ' Executa o comando para arduino-cli e abre a janela do prompt de comando
       ShellExecute Me.hWnd, "open", "cmd.exe", "/k " & commandCmd, vbNullString, vbNormalFocus
       
       KeyAscii = 0
   End If

End Sub

Private Sub cmdSketch_Click()
    ' Define o filtro para exibir apenas arquivos de texto
    CommonDialog1.Filter = "Arquivos de Texto (*.ino)|*.ino"
    
    ' Abre o diálogo de seleção de arquivo
    CommonDialog1.ShowOpen
    
    ' Obtém o caminho completo do arquivo selecionado
    Dim filePath As String
    filePath = CommonDialog1.FileName
    
    ' Exibe o caminho do arquivo no TextBox
    txtSketch.Text = filePath
    txtSketch.ToolTipText = txtSketch.Text
    
End Sub

Private Sub cmdCompile_Click()
    ' Define o caminho do arduino-cli
    Dim arduinoCliPath As String
    arduinoCliPath = "C:\arduino-cli.exe" ' Substitua pelo caminho correto do arduino-cli em seu sistema

    ' Define o caminho do arquivo .ino
    Dim sketchPath As String
    sketchPath = txtSketch.Text  ' Substitua "seu_arquivo.ino" pelo nome correto do seu arquivo .ino
    
    ' Define o comando para compilar o arquivo .ino
    Dim compileCmd As String
    'compileCmd = "arduino-cli compile --fqbn arduino:avr:uno -v -e " & sketchPath ' com arquivo binário, com detalhes processo
     compileCmd = "arduino-cli compile --fqbn arduino:avr:uno " & sketchPath ' sem arquivo binário, sem detalhes processo

    ' Executa o comando de compilação e abre a janela do prompt de comando
    'ShellExecute Me.hWnd, "open", "cmd.exe", "/k " & compileCmd, vbNullString, vbNormalFocus
    
    ' Escreve no console inicio do compile
    writeConsole (compileCmd)
    writeConsole ("Compile iniciado, aguarde...")
    DoEvents
    
    ' Executa script.vbs
    Call runCommand(compileCmd)
    
End Sub

Private Sub cmdUpload_Click()
    ' Se conectado, desconecta !!!
    If MSComm1.PortOpen = True Then Call cmdConexao_Click
    
    ' Define o caminho do arduino-cli
    Dim arduinoCliPath As String
    arduinoCliPath = "arduino-cli.exe" ' Substitua pelo caminho correto do arduino-cli em seu sistema

    ' Define o caminho do arquivo .ino
    Dim sketchPath As String
    sketchPath = txtSketch.Text ' Substitua "seu_arquivo.ino" pelo nome correto do seu arquivo .ino

    ' Define o comando para fazer o upload do arquivo compilado para o Arduino Uno
    Dim uploadCmd As String
    uploadCmd = "arduino-cli upload -p " & cboCommPort.Text & " --fqbn arduino:avr:uno -v " & sketchPath

    ' Executa o comando de upload e abre a janela do prompt de comando
    'ShellExecute Me.hWnd, "open", "cmd.exe", "/k " & uploadCmd, vbNullString, vbNormalFocus
        
    ' Escreve no console inicio do upload
    writeConsole (uploadCmd)
    writeConsole ("Upload iniciado, aguarde...")
    DoEvents
    
    ' Executa script.vbs
    Call runCommand(uploadCmd)
    
End Sub

Private Sub runCommand(strCommand As String)
    Dim cmd As String
    Dim output As String
    
    ' Executa o arquivo VBScript para obter a saída do comando do cmd
    cmd = "wscript.exe " & App.Path & "\command.vbs" & " """ & strCommand & """" & " " & 0
    output = CStr(CreateObject("WScript.Shell").Exec(cmd).StdOut.ReadAll())
    
    ' Cola a saída do cmd no TextBox
    writeConsole (output)
End Sub

Private Sub Timer1_Timer()
   If txtSketch.Text <> Empty Then
      cmdCompile.Enabled = True
      cmdUpload.Enabled = True
   Else
      cmdCompile.Enabled = False
      cmdUpload.Enabled = False
   End If
   
End Sub



