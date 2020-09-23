VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLaunch 
      Caption         =   "Launch Form in DLL"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'NOTE: You must have the DLL checked in the refeneces tab or it will not work,
'if you have it grouped, it will be at the top. If you browse for it display the
'description in the project properties. Also you can't dim the library (e.x MyAddIn)
'because then it won't show Loader (the class module name) in the Code Hints and will
'not work. Also when you close this form in the Launcher, it will also close out the
'form in the DLL (you can prevent this by simply opening the Loader class moduler and
'changing frmDLL.Show to frmDLL.Show 1 or frmDLL.Show vbModal

'We must have this below or it will not work, even if you type the REAL name.
'Just make up a similar name for what it is. For Example:
Dim LoadDLL As Loader 'Loader is the Class Module name in the DLL.

'Private Sub cmdLaunch_Click()
'Click Event for the Command button, Launch Form in DLL. This is used to launch the DLL
'and the form in it.
Private Sub cmdLaunch_Click()
    'We must do this otherwise we will get an error:
    Set LoadDLL = New Loader
    'This is the command to load the DLL, note that you must put it in the Class module.
    'In this example it is called Loader
    LoadDLL.LoadMe
End Sub
