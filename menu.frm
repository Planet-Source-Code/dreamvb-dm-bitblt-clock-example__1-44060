VERSION 5.00
Begin VB.Form frmmenu 
   ClientHeight    =   675
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   1560
   LinkTopic       =   "Form2"
   ScaleHeight     =   675
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuclock 
         Caption         =   "View &Clock"
      End
      Begin VB.Menu mnuVDate 
         Caption         =   "&View &Date"
      End
      Begin VB.Menu mnublank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
    Set frmmenu = Nothing ' unload the form from memory
End Sub

Private Sub mnuabout_Click()
    MsgBox "DM Clock Example by Ben Jones", vbInformation, "About...."
End Sub

Private Sub mnuclock_Click()
    frmmain.ViewOption = 0 'Show the clock
    frmmain.lbltitle.Width = 144 ' resize the label size
    frmmain.lbltitle.Caption = "Clock View"
    frmmain.lblclose.Left = frmmain.lbltitle.Width - frmmain.lblclose.Width - 5 ' position the close button
    frmmain.imgupdown.Left = frmmain.lblclose.Left - frmmain.imgupdown.Width - 4 ' position the updown button
    frmmain.Width = 2160 ' resize the form
End Sub

Private Sub mnuexit_Click()
    frmmain.Timer1.Enabled = False ' Disable the timer
    Unload frmmenu  ' Unload the main form
    Unload frmmain  ' Unload this form
End Sub

Private Sub mnuVDate_Click()
    frmmain.ViewOption = 1 ' Show the date
    frmmain.lbltitle.Width = 160 ' resize the lable size
    frmmain.lbltitle.Caption = "Date View"
    frmmain.lblclose.Left = frmmain.lbltitle.Width - frmmain.lblclose.Width - 5 ' position the close button
    frmmain.imgupdown.Left = frmmain.lblclose.Left - frmmain.imgupdown.Width - 4 ' position the updown button
    frmmain.Width = 2400 ' resize the form
    
End Sub
