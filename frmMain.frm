VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmMain 
   Caption         =   "Simple Web Browser using a Coolbar"
   ClientHeight    =   6360
   ClientLeft      =   675
   ClientTop       =   1800
   ClientWidth     =   9630
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   9630
   Begin ComctlLib.Toolbar tbrNav 
      Height          =   390
      Index           =   0
      Left            =   8400
      TabIndex        =   1
      Top             =   4560
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   688
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   327682
   End
   Begin ComctlLib.Toolbar tbrNav 
      Height          =   390
      Index           =   1
      Left            =   8400
      TabIndex        =   4
      Top             =   5040
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   688
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      _Version        =   327682
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   2535
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   5295
      ExtentX         =   9340
      ExtentY         =   4471
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
      Location        =   "http:///"
   End
   Begin PicClip.PictureClip ToolbarButtonsPicture 
      Left            =   3000
      Top             =   5640
      _ExtentX        =   11113
      _ExtentY        =   529
      _Version        =   393216
      Picture         =   "frmMain.frx":0442
   End
   Begin ComCtl3.CoolBar CoolBar 
      Height          =   765
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   1349
      BandCount       =   2
      MouseIcon       =   "frmMain.frx":1564
      EmbossPicture   =   -1  'True
      _CBWidth        =   9135
      _CBHeight       =   765
      _Version        =   "6.7.9782"
      MinHeight1      =   360
      Width1          =   9285
      Key1            =   "NavigationBand"
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Caption2        =   "Address"
      Child2          =   "cboURL"
      MinHeight2      =   315
      Width2          =   9675
      Key2            =   "AddressBand"
      NewRow2         =   -1  'True
      AllowVertical2  =   0   'False
      Begin VB.ComboBox cboURL 
         Height          =   315
         ItemData        =   "frmMain.frx":1580
         Left            =   855
         List            =   "frmMain.frx":158D
         TabIndex        =   2
         Top             =   420
         Visible         =   0   'False
         Width           =   8190
      End
   End
   Begin Browser.TransTBWrapper tbwNavWrapper 
      Height          =   315
      Left            =   6480
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
   End
   Begin ComctlLib.ImageList ToolbarImageList 
      Left            =   7560
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   16711935
      _Version        =   327682
   End
   Begin VB.Menu mnu_Toolbar 
      Caption         =   "Toolbar"
      Visible         =   0   'False
      Begin VB.Menu mnu_TBView 
         Caption         =   "Opaque"
         Checked         =   -1  'True
         Index           =   1
      End
      Begin VB.Menu mnu_TBView 
         Caption         =   "Transparent"
         Index           =   2
      End
      Begin VB.Menu mnu_Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_TBIconsWithText 
         Caption         =   "Icons With Te&xt"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_TBIconsOnly 
         Caption         =   "&Icons Only"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Save A&ll"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Propert&ies"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintSetup 
         Caption         =   "Print Set&up..."
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "Print Pre&view"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "Sen&d..."
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditPasteSpecial 
         Caption         =   "Paste &Special..."
      End
      Begin VB.Menu mnuEditBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditInvertSelection 
         Caption         =   "&Invert Selection"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
      End
      Begin VB.Menu mnuViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewLargeIcons 
         Caption         =   "Lar&ge Icons"
      End
      Begin VB.Menu mnuViewSmallIcons 
         Caption         =   "S&mall Icons"
      End
      Begin VB.Menu mnuViewList 
         Caption         =   "&List"
      End
      Begin VB.Menu mnuViewDetails 
         Caption         =   "&Details"
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewArrangeIcons 
         Caption         =   "Arrange &Icons"
         Begin VB.Menu mnuVAIByName 
            Caption         =   "by &Name"
         End
         Begin VB.Menu mnuVAIByType 
            Caption         =   "by &Type"
         End
         Begin VB.Menu mnuVAIBySize 
            Caption         =   "by Si&ze"
         End
         Begin VB.Menu mnuVAIByDate 
            Caption         =   "by &Date"
         End
      End
      Begin VB.Menu mnuViewLineUpIcons 
         Caption         =   "Li&ne Up Icons"
      End
      Begin VB.Menu mnuViewBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuGo 
      Caption         =   "&Go"
   End
   Begin VB.Menu mnuFavorites 
      Caption         =   "F&avorites"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "&Search For Help On..."
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About MyApp..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Enum ToolBarStyle
    tbrTextAndBMP = 0
    tbrBMPOnly = 1
End Enum

Enum ToolBarView
    tbrOpaque = 1
    tbrTransparent = 2
End Enum

Dim OriginalFormTitle As String
    
Private Sub Form_Load()
  
    OriginalFormTitle = Me.Caption
  
    ' _____ Create the Toolbars (with and without text).
    
    Call CreateToolbars
  
    ' _____ Setup the "NavigationBand" of the Coolbar.
  
    ' The container for the Toolbar Wrapper and the Toolbars is the Coolbar.
    Set tbwNavWrapper.Container = CoolBar(0)
    Set tbrNav(tbrBMPOnly).Container = CoolBar(0)
    Set tbrNav(tbrTextAndBMP).Container = CoolBar(0)
    
    ' The Child of this Band is the Toolbar Wrapper
    Set CoolBar(0).Bands.Item("NavigationBand").Child = tbwNavWrapper

    ' Known Issue: Setting transparency must be done after the Toolbar is visible.
    '              If Transparency is enabled while the toolbar is not visible, then
    '              Tool Tips will not work.  So, go ahead and show the form now.
    ' Me.Show
  
    ' Tell the Toolbar Wrapper we want the toolbar Transparent
    Call mnu_TBView_Click(tbrTransparent)
  
    ' Default the toolbar to Text and Bitmap
    Set tbwNavWrapper.Toolbar = tbrNav(tbrTextAndBMP)
                    
    ' _____ Load the web browser
    
    ' Initialize the combo box with the first entry (www.microsoft.com)
    ' and navigate there.
    cboURL.Text = cboURL.List(0)
    Call WebBrowser1.Navigate(cboURL.Text)
    
    DoEvents
    
End Sub

Public Sub CreateToolbars()
  
    Dim lCounter  As Long
    
    ' _____ Create the Image List from the PicClip Control
  
    ' Number of columns in the PicClip Control.
    Const CLIP_COLUMN_COUNT As Long = 14
     
    ' Setup the column count so the GraphicCell() call knows how big each picture is.
    ToolbarButtonsPicture.Cols = CLIP_COLUMN_COUNT
    
    ' Loop through the number of colums and add the image to the Image List
    For lCounter = 1 To CLIP_COLUMN_COUNT
        Call ToolbarImageList.ListImages.Add(lCounter, , _
            ToolbarButtonsPicture.GraphicCell(lCounter - 1))
    Next lCounter
      
    ' _____ Array to store names of Toolbar buttons and their properties
    '
    '       To add more buttons, just add a line to the array that contains
    '       the Key as the first string, and the button text as the second
    '       item.  Note we're only using the first six buttons, not all 14.
    '
    Dim maToolBarData As Variant
    maToolBarData = Array("BtnBack", "Back", _
                          "BtnForward", "Forward", _
                          "BtnStop", "Stop", _
                          "BtnRefresh", "Refresh", _
                          "BtnHome", "Home", _
                          "BtnSearch", "Search")
  
    Dim lUbound As Long
    lUbound = UBound(maToolBarData)
    
    ' This loop creates the two toolbars.
    ' One with text and bitmaps, one with bitmaps only.

    Dim lToolbar As Long
    For lToolbar = 0 To 1
        
        With tbrNav(lToolbar)
      
            ' This will eliminate flickering when the band is resized.
            .Wrappable = False
      
            .ImageList = ToolbarImageList
    
            Dim oButton As Button
      
            For lCounter = 0 To lUbound Step 2
                If lToolbar = tbrTextAndBMP Then
                    Set oButton = .Buttons.Add(, maToolBarData(lCounter), _
                                                 maToolBarData(lCounter + 1))
                Else  ' Assume BMP Only
                    Set oButton = .Buttons.Add(, maToolBarData(lCounter))
                End If
                oButton.ToolTipText = maToolBarData(lCounter + 1)
            Next lCounter
    
            Set oButton = Nothing
      
            ' _____ Add the bitmaps to the buttons
            
            For lCounter = 1 To (lUbound / 2)
              .Buttons(lCounter).Image = lCounter
            Next lCounter
    
        End With  ' tbrNav(lToolbar)
  
    Next lToolbar
  
End Sub

Private Sub cboURL_Click()

    On Error Resume Next
    Call WebBrowser1.Navigate(cboURL.Text)

End Sub

Private Sub cboURL_KeyPress(KeyAscii As Integer)

    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        Call WebBrowser1.Navigate(cboURL.Text)
    End If

End Sub



Private Sub tbrNav_ButtonClick(Index As Integer, ByVal Button As ComctlLib.Button)
            
    On Error Resume Next
    Select Case Button.Key
        
        Case "BtnBack"
            WebBrowser1.GoBack
        
        Case "BtnForward"
            WebBrowser1.GoForward
        
        Case "BtnStop"
            WebBrowser1.Stop
        
        Case "BtnRefresh"
            WebBrowser1.Refresh
        
        Case "BtnHome"
            WebBrowser1.GoHome
        
        Case "BtnSearch"
            WebBrowser1.GoSearch
    
    End Select

End Sub

Private Sub WebBrowser1_TitleChange(ByVal Text As String)
    
    Me.Caption = OriginalFormTitle & " - " & Text
    cboURL.Text = WebBrowser1.LocationURL

End Sub

Private Sub tbrNav_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  
    Select Case Button
        Case vbRightButton
            PopupMenu mnu_Toolbar
    End Select
  
End Sub

Private Sub mnu_TBView_Click(Index As Integer)
  
    If Index = tbrOpaque Then
        
        tbwNavWrapper.Transparent = False
        mnu_TBView(tbrOpaque).Checked = True
        mnu_TBView(tbrTransparent).Checked = False
    
    ElseIf Index = tbrTransparent Then
        
        tbwNavWrapper.Transparent = True
        mnu_TBView(tbrOpaque).Checked = False
        mnu_TBView(tbrTransparent).Checked = True
        
    End If

    Call Form_Resize

End Sub

Private Sub mnu_TBIconsOnly_Click()
  
    Set tbwNavWrapper.Toolbar = tbrNav(tbrBMPOnly)
    mnu_TBIconsOnly.Checked = True
    mnu_TBIconsWithText.Checked = False
    Call Form_Resize
  
End Sub

Private Sub mnu_TBIconsWithText_Click()
    
    Set tbwNavWrapper.Toolbar = tbrNav(tbrTextAndBMP)
    mnu_TBIconsOnly.Checked = False
    mnu_TBIconsWithText.Checked = True
    Call Form_Resize

End Sub

Private Sub Form_Resize()
  
    ' Get the width of the frame.
    Dim newScaleWidth As Long
    newScaleWidth = Me.ScaleWidth
    
    ' Move and stretch the Coolbar to be at the top and all the way across the form.
    Call CoolBar(0).Move(0, 0, newScaleWidth)
    
    ' Move and stretch the browser to be below the coolbar and all the way across the form.
    WebBrowser1.Left = 0
    WebBrowser1.Top = CoolBar(0).Top + CoolBar(0).Height
    WebBrowser1.Width = newScaleWidth
    WebBrowser1.Height = Me.ScaleHeight - WebBrowser1.Top
  
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  
    'Hiding the Coolbars before teardown reduces flickering.
    CoolBar(0).Visible = False
    
    'Be sure to set the TbwWrapper control's Toolbar properties
    'to Nothing in the form unload.  Otherwise, the application
    'will GPF on exit.
    Set tbwNavWrapper.Toolbar = Nothing

End Sub
