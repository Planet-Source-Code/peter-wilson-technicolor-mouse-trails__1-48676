VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Mouse Trails by Peter Wilson (http://dev.midar.com/)"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6060
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   6060
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   900
      Top             =   840
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type mdrBallType
    ParentIndex As Integer   ' Link to parent. This is a really cool feature as you can create complex hierarchies!
    Size As Single
    Mass As Single
    Colour As OLE_COLOR
    
    DesiredPositionX As Single
    DesiredPositionY As Single
    
    CurrentPositionX As Single
    CurrentPositionY As Single
End Type

Private m_Ball() As mdrBallType

Private Sub InitBalls()

    Dim intN As Integer
    ReDim m_Ball(60)    ' How many balls do we want? Don't forget 0 to 50 is actually 51 balls!
    
    ' Create the "root" ball.
    With m_Ball(0)
        .ParentIndex = -1
        .Size = 30                                     ' <<< Change for fun!
        .Mass = 3                                      ' <<< Change for fun!
        .Colour = HSV(0, 1, 1)                         ' <<< Change for fun!
    End With
    
    For intN = 1 To 60
        With m_Ball(intN)
            .ParentIndex = intN - 1                     ' <<<  Make each ball link to the previous ball (or any other way you like!)
            .Size = Abs(30 - intN) * 2                  ' <<< Change for fun!
            .Mass = 4                                   ' <<< Change for fun!
            .Colour = HSV((intN / 25) * 360, 1, 1)      ' <<< Change for fun!
        End With
    Next intN
    
End Sub


Private Sub Form_Load()

    ' Set some basic properties
    Me.BackColor = RGB(0, 0, 0)
    Me.FillStyle = vbFSSolid
    
    ' Create the ball hierarchy.
    Call InitBalls
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' Set where you would like the "root ball" to be.
    m_Ball(0).DesiredPositionX = X
    m_Ball(0).DesiredPositionY = Y
        
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    ' Reset the width and height of our form, and also move the origin (0,0) into
    ' the centre of the form. This makes our life much easier.
    Dim sngAspectRatio As Single
    sngAspectRatio = Me.Width / Me.Height
    
    Me.ScaleLeft = -1000
    Me.ScaleWidth = 2000
           
    Me.ScaleHeight = 2000 / sngAspectRatio
    Me.ScaleTop = -Me.ScaleHeight / 2
    
End Sub


Private Sub DrawCrossHairs()

    ' Draws cross-hairs going through the origin of the 2D window.
    ' ============================================================
    Me.DrawWidth = 1
    
    ' Draw Horizontal line (slightly darker to compensate for CRT monitors)
    Me.ForeColor = RGB(0, 64, 64)
    Me.Line (Me.ScaleLeft, 0)-(Me.ScaleWidth, 0)
    
    ' Draw Vertical line
    Me.ForeColor = RGB(0, 96, 96)
    Me.Line (0, Me.ScaleTop)-(0, Me.ScaleHeight)
    
End Sub

Private Sub DrawBalls(ParentIndex As Long)

    ' ====================================================================
    ' This is a recursive procedure, this means it calls itself!
    ' If you are a slacker, and put in the wrong parent id's you might get
    ' stuck in an infinite loop and your comptuer will run out of memory.
    ' ====================================================================
    On Error Resume Next ' Ignore errors (which can occur if you use really small masses)
    
    Dim lngIndex As Long
    Dim lngNewParent As Long
    
    Dim sngDeltaX As Single
    Dim sngDeltaY As Single
    
    Dim sngBallX As Single
    Dim sngBallY As Single
    
    ' Loop through the balls from the Lower Boundry to the Upper Boundry of the array.
    For lngIndex = LBound(m_Ball) To UBound(m_Ball)
        If m_Ball(lngIndex).ParentIndex = ParentIndex Then
            
            With m_Ball(lngIndex)
                
                Me.ForeColor = .Colour
                Me.FillColor = .Colour
                
                If ParentIndex = -1 Then ' "root ball"
                    
                    ' Calculate the difference between where the ball currently is, and where we would like it to be.
                    sngDeltaX = (.CurrentPositionX - .DesiredPositionX)
                    sngDeltaY = (.CurrentPositionY - .DesiredPositionY)
                    
                    ' Then move the ball closer to where it should be, depending on it's mass.
                    .CurrentPositionX = .CurrentPositionX - (sngDeltaX / .Mass)
                    .CurrentPositionY = .CurrentPositionY - (sngDeltaY / .Mass)
                
                Else
                
                    ' Calculate the difference between where the ball currently is, and where we would like it to be.
                    ' Note: Each child ball, seeks it's parents current location.
                    sngDeltaX = (.CurrentPositionX - m_Ball(ParentIndex).CurrentPositionX)
                    sngDeltaY = (.CurrentPositionY - m_Ball(ParentIndex).CurrentPositionY)
                    
                    ' Then move the ball closer to where it should be, depending on it's mass.
                    .CurrentPositionX = .CurrentPositionX - (sngDeltaX / .Mass)
                    .CurrentPositionY = .CurrentPositionY - (sngDeltaY / .Mass)
                                    
                End If
                
                ' Draw a pretty circle on Me (ie. the form)
                Me.Circle (.CurrentPositionX, .CurrentPositionY), .Size
                
                ' Now go a draw my children's balls! (Gee - I never thought I would ever type that sentance! ha ha)
                Call DrawBalls(lngIndex)
                ' Do not place code after this point (since this is a recursive routine) and any code placed
                ' here could potentially get called many times (which you probably don't want!)
                
            End With
        End If
    Next lngIndex
    
End Sub


Private Sub Timer1_Timer()

    Me.Cls
    Call DrawCrossHairs
    
    ' Draw the Root Ball.
    ' Note: Once the root ball is drawn,
    '       it will then draw it's children,
    '       and they in turn will draw their children, etc.
    Call DrawBalls(-1)

End Sub
