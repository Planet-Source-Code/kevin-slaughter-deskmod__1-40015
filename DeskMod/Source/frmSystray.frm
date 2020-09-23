VERSION 5.00
Begin VB.Form frmSysTray 
   BackColor       =   &H80000004&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DeskMod"
   ClientHeight    =   300
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   1515
   Icon            =   "frmSystray.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   300
   ScaleWidth      =   1515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "frmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*----------------------------------------*
'| Name: DeskMod                          |
'| Author: Kevin Slaughter                |
'| Date: 12/01/01 ~ 10/03/04              |
'| Webpage: www.Miryokuteki.com           |
'|----------------------------------------|
'| Purpose:  To be a simple app novelty   |
'|       app that sits in your systray    |
'|       for easy access. Also retains    |
'|       your settings, for re-use upon   |
'|       the next launch.                 |
'|                                        |
'| Demonstrates: Use of the system tray   |
'|               INI read/write           |
'|               SendMessage API          |
'|               Command line arg support |
'|                                        |
'| Comments: This was mainly just a       |
'|      theory at first, when I read that |
'|      the desktop was nothing but a big |
'|      ListView control. So I thought    |
'|      that if it is a LV, then I should |
'|      be able to change it's viewmode   |
'|      w/API, as I had seen done to a LV |
'|      a few days before.                |
'|                                        |
'|           Of course, I can't claim     |
'|      this as my idea entirely. It's    |
'|      been done numerous times. I just  |
'|      like having this option           |
'|      encapsulated into a nice-n-tiny   |
'|      systray app ^_^                   |
'*----------------------------------------*

'General needed stuff
Private gWnd As Long, ni As NOTIFYICONDATA
Private lCurrentView As Long, sCurrentView As String
Private bIsWinXP As Boolean





' Begin funx
Private Function ApplyLVM(ByVal lMode As Long)
    Select Case lMode
        Case 0
            'Normal
            Call SendMessage(gWnd, 4238, &H0, 0)
            lCurrentView = 0
            sCurrentView = "NORMAL" & Chr$(0)
        Case 1
            'Small
            Call SendMessage(gWnd, 4238, &H2, 0)
            lCurrentView = 1
            sCurrentView = "SMALLICONS" & Chr$(0)
        Case 2
            'List
            Call SendMessage(gWnd, 4238, &H3, 0)
            lCurrentView = 2
            sCurrentView = "LIST" & Chr$(0)
        Case 3
            'Details
            Call SendMessage(gWnd, 4238, &H1, 0)
            lCurrentView = 3
            sCurrentView = "DETAILS" & Chr$(0)
        Case 4
            'Tiles
            Call SendMessage(gWnd, 4238, &H4, 0)
            lCurrentView = 4
            sCurrentView = "TILES" & Chr$(0)
    End Select
End Function

'// Update notification icon
Private Function UpdateNI()
    ni.szTip = "Current view-mode: " & sCurrentView
    Call Shell_NotifyIcon(NIM_MODIFY, ni)
End Function

Private Sub Form_Initialize()
    If App.PrevInstance = True Then
        End 'THIS TRAIN STOPS HERE!
        '// Note: If "Unload Me" was used here, we'd
        '     fuck up the config file in our exit code,
        '     as no settings have been loaded yet. If
        '     this app is already running, and we mess
        '     with it's file, it won't be pretty :p.
    End If
    
    '// General setup
    bIsWinXP = IsWinXP
    Me.Move 0, 0, 0, 0
    
    '// Stupid way to have to get the listview, but it works..
    gWnd = FindWindow("Progman", "Program Manager")
    gWnd = GetWindow(gWnd, 5)
    gWnd = GetWindow(gWnd, 5)
    
    '// Create Notification icon
    With ni
        .cbSize = Len(ni)
        .hIcon = Me.Icon
        .hwnd = Me.hwnd
        .szTip = "Current view-mode: " & sCurrentView & Chr$(0)
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = &H200
    End With
    Call Shell_NotifyIcon(NIM_ADD, ni)
    Call Shell_NotifyIcon(NIM_MODIFY, ni)
    
    '// If no config is found, we default to the normal view
    If Len(Dir(App.Path & "\config.ini")) > 0 Then
        sCurrentView = UCase$(ReadINI(App.Path & "\config.ini", "Settings", "viewmode"))
        Select Case sCurrentView
            Case "NORMAL": Call ApplyLVM(0)
            Case "SMALLICONS": Call ApplyLVM(1)
            Case "LIST": Call ApplyLVM(2)
            Case "DETAILS": Call ApplyLVM(3)
            Case "TILES"
                If (bIsWinXP) Then
                    Call ApplyLVM(4)
                Else
                     Call ApplyLVM(0)   'Normal
                End If
            Case Else '// Someone fucked with the ini :-p
                Kill App.Path & "\config.ini"
                Call ApplyLVM(0)   'Normal is default
        End Select
    End If
    
    '// Moo!
    Call UpdateNI
End Sub

'// This lets us receive the notification icon info without
'     actually subclassing. Yay, no easy crashes! ^_^
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case x
        '// Right or left click
        Case 7710, 7755:
            Call ShowAPIMenu
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '// Reset to normal.. Write config file first, so when we change
    '//   back to normal mode, we don't always save as "NORMAL",
    '//   as it overrides the user's wishes.
    WriteINI App.Path & "\config.ini", "Settings", "viewmode", sCurrentView
    Call ApplyLVM(0)
    
    '// Kill notification icon and exit
    Call Shell_NotifyIcon(NIM_DELETE, ni)
End Sub

Private Function ShowAPIMenu()
    Dim hMen As Long, lChoice As Long, p As POINTAPI, r As RECT, bDie As Boolean
    hMen = CreatePopupMenu()
    bDie = False
    
    If (hMen) Then
        '// Add items to menu
        Call AppendMenu(hMen, IIf(lCurrentView = 0, MF_STRING Or MF_CHECKED, MF_STRING), &H10, "Normal")
        Call AppendMenu(hMen, IIf(lCurrentView = 1, MF_STRING Or MF_CHECKED, MF_STRING), &H20, "Small Icons")
        Call AppendMenu(hMen, IIf(lCurrentView = 2, MF_STRING Or MF_CHECKED, MF_STRING), &H30, "List")
        Call AppendMenu(hMen, IIf(lCurrentView = 3, MF_STRING Or MF_CHECKED, MF_STRING), &H40, "Details")
        If (bIsWinXP) Then
            Call AppendMenu(hMen, IIf(lCurrentView = 4, MF_STRING Or MF_CHECKED, MF_STRING), &H50, "Tiles")
        End If
        Call AppendMenu(hMen, MF_SEPARATOR, &H60, "-")
        Call AppendMenu(hMen, MF_STRING, &H70, "Exit")
        
        '// Display menu at mouse (return item id, not pos)
        Call GetCursorPos(p)
        
        '// If we don't set our dummy window as the foreground window,
        '     the menu won't go away if the user doesn't click on it.
        Call SetForegroundWindow(Me.hwnd)
        lChoice = TrackPopupMenu(hMen, &H100&, p.x, p.y, 0, Me.hwnd, r)
        
        '// Parse user's selection
        Select Case lChoice
            Case &H10: Call ApplyLVM(0)
            Case &H20: Call ApplyLVM(1)
            Case &H30: Call ApplyLVM(2)
            Case &H40: Call ApplyLVM(3)
            Case &H50: Call ApplyLVM(4)
            Case &H70: bDie = True
        End Select
        
        'Tidy up
        Call UpdateNI
    
        'Unload menu from memory
        Call DestroyMenu(hMen)
        If (bDie) Then Unload Me
    Else
        Call MsgBox("Error creating DeskMod menu", vbExclamation)
        Exit Function
    End If
End Function
