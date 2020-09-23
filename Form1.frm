VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   WindowState     =   2  'Maximiert
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   120
      Pattern         =   "*.png"
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const bytMaxSize As Byte = 128
Private Const bytMinSize As Byte = 64

Private Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    Color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type
Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR
Dim Dx As DirectX8
Dim D3D As Direct3D8
Dim D3DDevice As Direct3DDevice8
Dim MouseBox(0 To 3) As TLVERTEX
Dim D3DX As D3DX8
Dim TexWithA As Direct3DTexture8
Dim DispMode As D3DDISPLAYMODE
Dim D3DWindow As D3DPRESENT_PARAMETERS
Dim sngHeight As Single, sngWidth As Single, sngLeft As Single, sngTop As Single
Dim sngIndex As Single, sngUBound As Single, sngStep As Single, sngStartTop As Single, sngStartLeft As Single
Dim sngFrom() As Single, strTest As String

Private Function CreateTLVertex(X As Single, Y As Single, Z As Single, rhw As Single, Color As Long, Specular As Long, tu As Single, tv As Single) As TLVERTEX
CreateTLVertex.X = X
CreateTLVertex.Y = Y
CreateTLVertex.Z = Z
CreateTLVertex.rhw = rhw
CreateTLVertex.Color = Color
CreateTLVertex.Specular = Specular
CreateTLVertex.tu = tu
CreateTLVertex.tv = tv
End Function

Private Sub Form_Click()
MsgBox sngIndex
End
End Sub

Private Sub Form_Load()
Set Dx = New DirectX8
Set D3D = Dx.Direct3DCreate()
Set D3DX = New D3DX8

Me.Height = Screen.Height
Me.Width = Screen.Width

D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
D3DWindow.Windowed = 1
D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
D3DWindow.BackBufferFormat = DispMode.Format
D3DWindow.hDeviceWindow = Me.hWnd

Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, Me.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)

File1.Path = App.Path & "\Icons\"
ReDim Preserve sngFrom(File1.ListCount)
sngUBound = File1.ListCount - 1
sngStartTop = (Me.Height / Screen.TwipsPerPixelX) - bytMaxSize
sngStartLeft = ((Me.Width / Screen.TwipsPerPixelX) - (File1.ListCount * bytMinSize)) / 2
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y < Me.Height - (bytMaxSize * Screen.TwipsPerPixelY) Or X < (sngStartLeft * Screen.TwipsPerPixelX) Or X > sngFrom(UBound(sngFrom)) Then
    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
    D3DDevice.BeginScene
    D3DDevice.SetVertexShader FVF

    sngStartLeft = ((Me.Width / Screen.TwipsPerPixelX) - (File1.ListCount * bytMinSize)) / 2
    sngLeft = sngStartLeft
    For a = 0 To File1.ListCount - 1
        sngHeight = bytMinSize
        sngWidth = bytMinSize
        sngTop = sngStartTop + bytMaxSize - bytMinSize

        MouseBox(0) = CreateTLVertex(sngLeft, sngTop, 0, 1, &HFFFFFF, 0, 0, 0)
        MouseBox(1) = CreateTLVertex(sngLeft + sngWidth, sngTop, 0, 1, &HFFFFFF, 0, 1, 0)
        MouseBox(2) = CreateTLVertex(sngLeft, sngTop + sngHeight, 0, 1, &HFFFFFF, 0, 0, 1)
        MouseBox(3) = CreateTLVertex(sngLeft + sngWidth, sngTop + sngHeight, 0, 1, &HFFFFFF, 0, 1, 1)

        D3DDevice.SetTexture 0, D3DX.CreateTextureFromFile(D3DDevice, App.Path & "\Icons\" & File1.List(a))
        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, MouseBox(0), Len(MouseBox(0))
        D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False
        
        sngFrom(a) = sngLeft * Screen.TwipsPerPixelX
        sngLeft = sngLeft + sngWidth
    Next a
    sngFrom(UBound(sngFrom)) = sngLeft * Screen.TwipsPerPixelX
    
    D3DDevice.EndScene
    D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
    Exit Sub
End If

For a = 0 To sngUBound
    If X >= sngFrom(a) And X <= sngFrom(a + 1) Then
        sngIndex = a
        Exit For
    End If
Next a
sngLeft = sngStartLeft
sngStep = (X - sngFrom(sngIndex)) / (2 * Screen.TwipsPerPixelX)

D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
D3DDevice.BeginScene
D3DDevice.SetVertexShader FVF

For a = 0 To sngUBound
    If a <> sngIndex And (a <> sngIndex - 1 And a <> sngIndex + 1) Then
        sngHeight = bytMinSize
        sngWidth = bytMinSize
        sngTop = sngStartTop + bytMaxSize - bytMinSize
    ElseIf a = sngIndex - 1 Then
        sngHeight = bytMaxSize - sngStep
        sngWidth = bytMaxSize - sngStep
        sngTop = sngStartTop + bytMaxSize - (bytMaxSize - sngStep)
    ElseIf a = sngIndex Then
        sngHeight = bytMaxSize
        sngWidth = bytMaxSize
        sngTop = sngStartTop + bytMaxSize - bytMaxSize
    ElseIf a = sngIndex + 1 Then
        sngHeight = bytMinSize + sngStep
        sngWidth = bytMinSize + sngStep
        sngTop = sngStartTop + bytMaxSize - (bytMinSize + sngStep)
    End If

    MouseBox(0) = CreateTLVertex(sngLeft, sngTop, 0, 1, &HFFFFFF, 0, 0, 0)
    MouseBox(1) = CreateTLVertex(sngLeft + sngWidth, sngTop, 0, 1, &HFFFFFF, 0, 1, 0)
    MouseBox(2) = CreateTLVertex(sngLeft, sngTop + sngHeight, 0, 1, &HFFFFFF, 0, 0, 1)
    MouseBox(3) = CreateTLVertex(sngLeft + sngWidth, sngTop + sngHeight, 0, 1, &HFFFFFF, 0, 1, 1)

    D3DDevice.SetTexture 0, D3DX.CreateTextureFromFile(D3DDevice, App.Path & "\Icons\" & File1.List(a))
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, MouseBox(0), Len(MouseBox(0))
    D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, False

    sngFrom(a) = sngLeft * Screen.TwipsPerPixelX
    sngLeft = sngLeft + sngWidth
Next a
sngFrom(UBound(sngFrom)) = sngLeft * Screen.TwipsPerPixelX
sngStartLeft = ((Me.Width / Screen.TwipsPerPixelX) - (sngLeft - sngStartLeft)) / 2

D3DDevice.EndScene
D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set D3DDevice = Nothing
Set D3D = Nothing
Set Dx = Nothing
End
End Sub
