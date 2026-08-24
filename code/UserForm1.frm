VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Shape Properties"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   500
   ClientWidth     =   4780
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private IsMouseInside As Boolean

Private Type HeadingData
    Title As String
    RangeStart As Long
End Type

Dim Headings() As HeadingData
Public WithEvents NavListBox As MSForms.ListBox
Attribute NavListBox.VB_VarHelpID = -1
Public WithEvents ResizeGrip As MSForms.Label
Attribute ResizeGrip.VB_VarHelpID = -1

' Drag tracking variables
Dim IsResizing As Boolean
Dim StartX As Single
Dim StartY As Single
Dim StartWidth As Single
Dim StartHeight As Single

Dim gripRight As Long, gripBottom As Long

Private Sub UserForm_Initialize()
    Dim ctrl As Control
    Dim w As Long, h As Long: w = 300: h = 350
    gripRight = 35: gripBottom = 35

    ' Set up UserForm window size
    Me.Caption = "" ''"Heading Navigator"
    Me.Width = w: Me.Height = h

    ' Hide all static imported controls
    For Each ctrl In Me.Controls
        ctrl.Visible = False
    Next ctrl

    ' 1. Create ListBox (Sized SHORTER so it doesn't cover the bottom corner)
    Set NavListBox = Me.Controls.Add("Forms.ListBox.1", "NavListBox", True)
    With NavListBox
        .Left = -2
        .Top = -2
        .Width = w '- 15
        .Height = h - gripBottom '65 ' Leaves clear space at bottom
        .Font.Name = "Helvetica"
        .Font.Size = 14
    End With
    
    ' 2. Create Drag-Resize Grip in bottom-right corner
    Set ResizeGrip = Me.Controls.Add("Forms.Label.1", "ResizeGrip", True)
    With ResizeGrip
        .Caption = ": :" '"//"
        .Font.Name = "Arial"
        .Font.Size = 16
        .ForeColor = RGB(160, 160, 160)
        .Width = 20
        .Height = 20
        .Left = w - gripRight + 15 '35
        .Top = h - gripBottom - 10 '60
        .MousePointer = fmMousePointerSizeNWSE
    End With

    ' Populate ListBox with Heading 1s
    Call LoadHeadings
    If NavListBox.ListCount = 0 Then
        NavListBox.AddItem "No Heading 1 styles found."
    End If
    
    
    On Error Resume Next
    Dim dummyFrame As MSForms.Frame
    Set dummyFrame = Me.Controls.Add("Forms.Frame.1", "fraFocusHolder", True)
    With dummyFrame
        .Left = 0
        .Top = 0
        .Width = 1
        .Height = 1
        .BorderStyle = fmBorderStyleNone
        .TabStop = True
    End With
    On Error GoTo 0
    
    IsInternalCall = True
    Call z_help
End Sub

Private Sub NavListBox_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Detect if the mouse is near the outer boundary of the listbox (exiting)
    Dim bound As Long: bound = 20
    If X < bound Or Y < bound Or X > (NavListBox.Width - bound) Or Y > (NavListBox.Height - bound) Then
        IsMouseInside = False
    Else
        ' Cursor is inside the central area
        If Not IsMouseInside Then
            IsMouseInside = True
            Call LoadHeadings
        End If
    End If
End Sub
Public Sub LoadHeadings()
    Dim doc As Document
    Dim para As Paragraph
    Dim count As Long
    
    Set doc = ActiveDocument
    
    ' Clear existing listbox items
    Application.ScreenUpdating = False
    NavListBox.Clear
    
    count = 0
    ReDim Headings(0)
    
    ' Scan document for headings and populate array/listbox
    For Each para In doc.Paragraphs
        If Left(para.Style.NameLocal, 7) = "Heading" Or para.OutlineLevel <= wdOutlineLevel9 Then
            If Trim(para.Range.Text) <> "" Then
                ReDim Preserve Headings(count)
                Headings(count).RangeStart = para.Range.Start
                Headings(count).Title = Trim(Replace(para.Range.Text, vbCr, ""))
                
                NavListBox.AddItem Headings(count).Title
                count = count + 1
            End If
        End If
    Next para
    
    Application.ScreenUpdating = True
End Sub
Private Sub NavListBox_Click()
    Dim idx As Long
    idx = NavListBox.ListIndex

    If idx >= 0 Then
        If UBound(Headings) >= idx Then
            Dim targetPos As Long
            Dim targetRng As Range

            targetPos = Headings(idx).RangeStart
            Set targetRng = ActiveDocument.Range(targetPos, targetPos)

            targetRng.Select

            Application.ScreenUpdating = False
            ActiveWindow.ActivePane.LargeScroll Down:=1
            ActiveWindow.ScrollIntoView Selection.Range, True
            ActiveWindow.ActivePane.SmallScroll Down:=6 '12
            Application.ScreenUpdating = True
        End If

        IsInternalCall = True
        Application.OnTime Now + 5E-06, "z_help"
    End If
End Sub

' 1. Start dragging the resize grip
Private Sub ResizeGrip_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then ' Left mouse button
        IsResizing = True
        StartX = X
        StartY = Y
        StartWidth = Me.Width
        StartHeight = Me.Height
    End If
End Sub

Private Sub ResizeGrip_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If IsResizing Then
        Dim NewWidth As Single, NewHeight As Single
        
        ' Calculate target dimensions based on mouse offset
        NewWidth = StartWidth + (X - StartX)
        NewHeight = StartHeight + (Y - StartY)

        ' Apply minimum boundaries
        If NewWidth < 200 Then NewWidth = 200
        If NewHeight < 200 Then NewHeight = 200

        ' 1:1 resize update
        Me.Width = NewWidth
        Me.Height = NewHeight

        ' Re-bind ListBox and Grip to updated form bounds
        NavListBox.Width = Me.Width '- 15
        NavListBox.Height = Me.Height - gripBottom - 5 '65
        ResizeGrip.Left = Me.Width - gripRight + 15 '35
        ResizeGrip.Top = Me.Height - gripBottom - 10 'Me.Height - 60
        
        ' Update baseline tracking coordinates to match new position
        StartWidth = Me.Width
        StartHeight = Me.Height
    End If
End Sub

' 3. Release mouse to finish resizing
Private Sub ResizeGrip_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    IsResizing = False
End Sub
