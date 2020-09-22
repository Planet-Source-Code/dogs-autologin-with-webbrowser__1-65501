VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   ScaleHeight     =   9465
   ScaleWidth      =   12030
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   9015
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   11775
      ExtentX         =   20770
      ExtentY         =   15901
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Andy Hughes (c)2006
' any queries pleaase feel free to contact me on andy@andythughes.co.uk
'
'assumes that you have a webbrowser on the form at form load

Dim Username As String
Dim Password As String
Dim Website As String

Private Sub Form_Load()

    Website = "[What Ever You Website Is]"
    Username = "[Valid Username]"
    Password = "[and Password]"

    WebBrowser1.Navigate (Website)

End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)


    ' the line below is dependant on what is
    ' returned as the web site url once you have
    ' entered you initail web address, note
    ' the return url is rarely the same as the passed url.
    
    If URL = Website Then
    
        WebBrowser1.Document.All.Item("user").Value = Username 'your Username text here
        WebBrowser1.Document.All.Item("pass").Value = Password 'your Password text here
        
        ' now we tell the browser to locate the
        ' login button and click it. all the variables stated User,Pass and buttonlogin
        ' are taken by viewing the source of the webpage you want to login to.
        
        WebBrowser1.Document.All.Item("buttonlogin").Click
    
        
    End If


End Sub

