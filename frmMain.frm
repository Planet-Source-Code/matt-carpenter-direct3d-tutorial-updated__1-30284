VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'General Declarations
Dim pos As D3DVECTOR
'These three components are the big daddy of what your 3D Program.
Dim DX_Main As New DirectX7 ' The DirectX core file (The heart of it all)
Dim DD_Main As DirectDraw4 ' The DirectDraw Object
Dim D3D_Main As Direct3DRM3 ' The direct3D Object

'DirectInput Components
Dim DI_Main As DirectInput ' The DirectInput core
Dim DI_Device As DirectInputDevice ' The DirectInput device
Dim DI_State As DIKEYBOARDSTATE 'Array Holding the state of the keys

'DirectDraw Surfaces- Where the screen is drawn.
Dim DS_Front As DirectDrawSurface4 ' The frontbuffer (What you see on the screen)
Dim DS_Back As DirectDrawSurface4 ' The backbuffer, (Where everything is drawn before it's put on the screen.)
Dim SD_Front As DDSURFACEDESC2 ' The SurfaceDescription
Dim DD_Back As DDSCAPS2 ' General Surface Info

'ViewPort and Direct3D Device
Dim D3D_Device As Direct3DRMDevice3 'The Main Direct3D Retained Mode Device
Dim D3D_ViewPort As Direct3DRMViewport2 'The Direct3D Retained Mode Viewport (Kinda the camera)

'The Frames
Dim FR_Root As Direct3DRMFrame3 'The Main Frame (The other frames are put under this one (Like a tree))
Dim FR_Camera As Direct3DRMFrame3 'Another frame, just happens to be called 'camera'. We will use this
                                                   'as the viewport. Hence, the name camera. Doesn't have to be
                                                   'called 'camera'. It could be called 'BillyBob' for all I care.
Dim FR_Light As Direct3DRMFrame3 'This frame contains our, guess what, spotlight!
Dim FR_Building As Direct3DRMFrame3 'Frame containing our 1st mesh that will be put in this "game".

'Meshes (3D objects loaded from a .x file)
Dim MS_Building As Direct3DRMMeshBuilder3

'Lights
Dim LT_Ambient As Direct3DRMLight 'Our Main (Ambient) light that illuminates everything (not just part of something
                                                'like a spotlight.
Dim LT_Spot As Direct3DRMLight 'Our Spot light, makes it look more realistic.

'Camera Positions
Dim xx As Long
Dim yy As Long
Dim zz As Long

Dim esc As Boolean 'If Escape is pressed, the DX_Input sub will make it true and the main loop will end.
'Incase you haven't caught on, the prefix FR = frame, MS = Mesh, & LT = light.

Dim BackGround As Direct3DRMTexture3 'This will be the texture that holds our background.
'Note: Texture files must have side lengths that are devisible by 2!
'======================================================================================

Private Sub DX_Init()
 'Type, not copy and paste!
 'This sub will initialize all your components and set them up.
 Set DD_Main = DX_Main.DirectDraw4Create("") 'Create the DirectDraw Object

 DD_Main.SetCooperativeLevel Form1.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE 'Set Screen Mode (Full 'Screen)
 DD_Main.SetDisplayMode 640, 480, 32, 0, DDSDM_DEFAULT 'Set Resolution and BitDepth (Lets use 32-bit color)
 
 SD_Front.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
 SD_Front.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_3DDEVICE Or DDSCAPS_COMPLEX Or _
 DDSCAPS_FLIP 'I used the line-continuation ( _ ) because the whole thing wouldn't fit on one line...
 SD_Front.lBackBufferCount = 1 'Make one backbuffer
 Set DS_Front = DD_Main.CreateSurface(SD_Front) 'Initialize the front buffer (the screen)
 'The Previous block of code just created the screen and the backbuffer.
 
 DD_Back.lCaps = DDSCAPS_BACKBUFFER
 
 Set DS_Back = DS_Front.GetAttachedSurface(DD_Back)
 DS_Back.SetForeColor RGB(255, 255, 255)
 'The backbuffer was initialized and the DirectDraw text color was set to white.

 Set D3D_Main = DX_Main.Direct3DRMCreate() 'Creates the Direct3D Retained Mode Object!!!!!

 Set D3D_Device = D3D_Main.CreateDeviceFromSurface("IID_IDirect3DHALDevice", DD_Main, DS_Back, _
 D3DRMDEVICE_DEFAULT) 'Tell the Direct3D Device that we are using hardware rendering (HALDevice) instead
                                   'of software enumeration (RGBDevice).
 D3D_Device.SetBufferCount 2 'Set the number of buffers
 D3D_Device.SetQuality D3DRMRENDER_GOURAUD 'Set Rendering Quality. Can use Flat, or WireFrame, but
                                                                  'GOURAUD has the best rendering quality.
 D3D_Device.SetTextureQuality D3DRMTEXTURE_NEAREST 'Set the texture quality
 D3D_Device.SetRenderMode D3DRMRENDERMODE_BLENDEDTRANSPARENCY 'Set the render mode.

 Set DI_Main = DX_Main.DirectInputCreate() 'Create the DirectInput Device
 Set DI_Device = DI_Main.CreateDevice("GUID_SysKeyboard") 'Set it to use the keyboard.
 DI_Device.SetCommonDataFormat DIFORMAT_KEYBOARD 'Set the data format to the keyboard format.
 DI_Device.SetCooperativeLevel Me.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE 'Set Coperative Level.
 DI_Device.Acquire
 'The above block of code configures the DirectInput Device and starts it.
End Sub

'=====================================================================================

Private Sub DX_MakeObjects()
 Set FR_Root = D3D_Main.CreateFrame(Nothing) 'This will be the root frame of the 'tree'
 Set FR_Camera = D3D_Main.CreateFrame(FR_Root) 'Our Camera's Sub Frame. It goes under FR_Root in the 'Tree'.
 Set FR_Light = D3D_Main.CreateFrame(FR_Root) 'Our Light's Sub Frame
 Set FR_Building = D3D_Main.CreateFrame(FR_Root) 'Our Building (the 3D Thingy that will be placed in our world's)
                                                                      ' sub frame.
 'That above code set up the hierarchy of frames where FR_Root is the parent, and the other frames are all
 'owned by it.
 
 FR_Root.SetSceneBackgroundRGB 0, 0, 1 'Set the background color. Use decimals, not the standerd 255 = max.
                                                         'What I have here will make the background/sky 100% blue.
 FR_Camera.SetPosition Nothing, 1, 4, -35 'Set The Camera position; X=1, Y=1, z=0.
 Set D3D_ViewPort = D3D_Main.CreateViewport(D3D_Device, FR_Camera, 0, 0, 640, 480) 'Make our viewport and set
                                                                                                                  'it our camera to be it.
 D3D_ViewPort.SetBack 200 'How far back it will draw the image. (Kinda like a visibility limit)
  
 FR_Light.SetPosition Nothing, 1, 6, -20 'Set our 'point' light.
 Set LT_Spot = D3D_Main.CreateLightRGB(D3DRMLIGHT_POINT, 1, 1, 1) 'Set the light type and it's color.
 FR_Light.AddLight LT_Spot 'Add the light to it's frame.
  
 Set LT_Ambient = D3D_Main.CreateLightRGB(D3DRMLIGHT_AMBIENT, 0.5, 0.5, 0.5) 'Create our ambient light.
 FR_Root.AddLight LT_Ambient 'Add the ambient light to the root frame.
 
 Set BackGround = D3D_Main.LoadTexture(App.Path & "\background.bmp") 'Load our background texture and put it into BackGround
 FR_Root.SetSceneBackgroundImage BackGround 'Take our texture and make it the scene's background.
 
 Set MS_Building = D3D_Main.CreateMeshBuilder() 'Make the 3D Building Mesh
 MS_Building.LoadFromFile App.Path & "\building.x", 0, 0, Nothing, Nothing 'Load our building mesh from its .X file.
 MS_Building.ScaleMesh 0.5, 0.5, 0.5 'Set the it's scale. This is used to make the object smaller or bigger. 1 makes
                                          'it the same size as it was built in whatever program it was built in. .5 is half as big
 FR_Building.AddVisual MS_Building 'Add the 3D Building mesh to it's frame.

End Sub

'=====================================================================================

Private Sub DX_Render()
 'Lets put our main loop. Make it loop until esc = true (I'll explain later)
 Do While esc = False
   On Local Error Resume Next 'Incase there is an error
   DoEvents 'Give the computer time to do what it needs to do.
   DX_Input 'Call the input sub.
   FR_Camera.GetPosition Nothing, pos
   D3D_ViewPort.Clear D3DRMCLEAR_TARGET Or D3DRMCLEAR_ZBUFFER 'Clear your viewport.
   D3D_Device.Update 'Update the Direct3D Device.
   D3D_ViewPort.Render FR_Root 'Render the 3D Objects (lights, and your building!)
   DS_Back.DrawText 200, 0, "Direct3D Example X: " & pos.x & " Z: " & pos.z, False  'Draw some text!
   DS_Front.Flip Nothing, DDFLIP_WAIT 'Flip the back buffer with the front buffer.
 Loop
End Sub

'=====================================================================================

Private Sub DX_Input()
    Const Sin5 = 8.715574E-02!  ' Sin(5°)
    Const Cos5 = 0.9961947!     ' Cos(5°)
 DI_Device.GetDeviceStateKeyboard DI_State 'Get the array of keyboard keys and their current states
  
 If DI_State.Key(DIK_ESCAPE) <> 0 Then Call DX_Exit 'If user presses [esc] then exit end the program.
 
 If DI_State.Key(DIK_LEFT) <> 0 Then 'Quick Note: <> means 'does not'
   'FR_Camera.SetPosition FR_Camera, -1, 0, 0 'Move the viewport to the left
   FR_Camera.SetOrientation FR_Camera, -Sin5, 0, Cos5, 0, 1, 0
 End If
 
 If DI_State.Key(DIK_RIGHT) <> 0 Then
   'FR_Camera.SetPosition FR_Camera, 1, 0, 0 'Move the viewport to the right
   FR_Camera.SetOrientation FR_Camera, Sin5, 0, Cos5, 0, 1, 0
 End If
 
 If DI_State.Key(DIK_UP) <> 0 Then
   FR_Camera.SetPosition FR_Camera, 0, 0, 1 'Move the viewport forward
 End If

 If DI_State.Key(DIK_DOWN) <> 0 Then
   FR_Camera.SetPosition FR_Camera, 0, 0, -1 'Move the viewport back
 End If

End Sub

'=====================================================================================

Private Sub DX_Exit()
 Call DD_Main.RestoreDisplayMode
 Call DD_Main.SetCooperativeLevel(Me.hWnd, DDSCL_NORMAL)
  Call DI_Device.Unacquire
 'Restore all the devices

 End 'Ends the program.
End Sub

'=====================================================================================

Private Sub Form_Load()
 Me.Show 'Some computers do weird stuff if you don't show the form.
 DoEvents 'Give the computer time to do what it needs to do
 DX_Init 'Initialize DirectX
 DX_MakeObjects 'Make frames, lights, and mesh(es)
 DX_Render 'The Main Loop

End Sub


