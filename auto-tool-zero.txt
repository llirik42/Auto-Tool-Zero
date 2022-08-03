Const FEED_RATE As Integer = 400

Const EPSILON As Double = 0.001

Const CURRENT_X_OEM_CODE As Integer = 800

Const X_COORDINATE_NUMBER As Integer = 0
Const Y_COORDINATE_NUMBER As Integer = 1
Const Z_COORDINATE_NUMBER As Integer = 2

Const CURRENT_FEED_RATE_OEM_CODE As Integer = 818

Const SET_FEED_RATE_CODE As String = "F"
Const FAST_MODE_CODE As String = "G00"
Const MOVE_UNTILL_TOUCH_OR_REACH_CODE As String = "G31"

Const X_SIZE As Double = 30.0
Const Y_SIZE As Double = 30.0
Const X_THICKNESS As Double = 15.0
Const Y_THICKNESS As Double = 15.0
Const Z_THICKNESS As Double = 8.0

Const MAX_DEPTH As Double = 40.0
Const OFFSET As Double = 25.0
Const TOOL_RADIUS As Double = 1.5875


Call main()


Sub main()
	originFeedRate = getCurrentFeedRate()		
	setCurrentFeed(FEED_RATE)

	startXPosition = getCurrentX()
	startYPosition = getCurrentY()
	startZPosition = getCurrentZ()
	
	zTouch = autoZeroZ(startZPosition)
	If zTouch = False Then
		setCurrentFeed(originFeedRate)
		Return
	End If
	
	xTouch = autoZeroX(startXPosition)
	If xTouch = False Then
		setCurrentFeed(originFeedRate)
		Return
	End If
		
	yTouch = autoZeroY(startYPosition)
	
	While IsMoving
	Wend
	
	setCurrentZ(OFFSET + Z_THICKNESS)
	
	setCurrentFeed(originFeedRate)	
End Sub


Function autoZeroX(ByVal startXPosition As Double) As Boolean
	moveFastOnX(-X_SIZE)
	
	moveFastToZ(-Z_THICKNESS / 2)
		
	Sleep 4000
	
	moveUntillTouchOrReachX(startXPosition)
	
	If isZero(getCurrentX() - startXPosition) Then
		moveFastOnX(-X_SIZE)
		moveFastToZ(OFFSET)
		moveFastToX(startXPosition)
		printNoTouchError("X")
		autoZeroX = False
	Else
		setCurrentX(-X_THICKNESS - TOOL_RADIUS)
		moveFastOnX(-OFFSET)
		moveFastToZ(OFFSET)
		moveFastToX(0)
		autoZeroX = True
	End If
End Function

Function autoZeroY(ByVal startYPosition As Double) As Boolean
	moveFastOnY(-Y_SIZE)
	
	moveFastToZ(-Z_THICKNESS / 2)
		
	Sleep 6000
	
	moveUntillTouchOrReachY(startYPosition)
	
	If isZero(getCurrentY() - startYPosition) Then
		moveFastOnY(-Y_SIZE)
		moveFastToZ(OFFSET)
		moveFastToY(startYPosition)
		printNoTouchError("Y")	
		autoZeroY = False
	Else
		setCurrentY(-Y_THICKNESS - TOOL_RADIUS)
		moveFastOnY(-OFFSET)
		moveFastToZ(OFFSET)
		moveFastToY(0)
		autoZeroY = True
	End If

End Function

Function autoZeroZ(ByVal startZPosition As Double) As Boolean			
	moveUntillTouchOrReachZ(startZPosition - MAX_DEPTH)
		
	If reachedMaxDepth(startZPosition) Then
		moveFastToZ(startZPosition)
		printNoTouchError("Z")
		autoZeroZ = False
	Else 
		setCurrentZ(0)
		moveFastToZ(OFFSET)
		autoZeroZ = True
	End If
End Function


Function reachedMaxDepth(ByVal startZPosition As Double)
	reachedMaxDepth = isZero(getCurrentZ() - startPozition + MAX_DEPTH)
End Function


Sub moveFastOnX(dx)
	moveFastToX(getCurrentX() + dx)
End Sub

Sub moveFastOnY(dy)
	moveFastToY(getCurrentY() + dy)
End Sub

Sub moveFastOnZ(dz)
	moveFastToZ(getCurrentZ() + dz)
End Sub


Sub moveFastToX(xDestination)
	Code FAST_MODE_CODE & "X" & xDestination
End Sub

Sub moveFastToY(yDestination)
	Code FAST_MODE_CODE & "Y" & yDestination
End Sub

Sub moveFastToZ(zDestination)
	Code FAST_MODE_CODE & "Z" & zDestination
End Sub


Sub moveUntillTouchOrReachX(xDestination As Double)
	Code MOVE_UNTILL_TOUCH_OR_REACH_CODE & "X" & xDestination
	While IsMoving
	Wend
End Sub

Sub moveUntillTouchOrReachY(yDestination As Double)
	Code MOVE_UNTILL_TOUCH_OR_REACH_CODE & "Y" & yDestination
	While IsMoving
	Wend
End Sub

Sub moveUntillTouchOrReachZ(zDestination As Double)
	Code MOVE_UNTILL_TOUCH_OR_REACH_CODE & "Z" & zDestination
	While IsMoving
	Wend
End Sub


Function getCurrentFeedRate()
	getCurrentFeedRate = GetOEMDRO(CURRENT_FEED_RATE_OEM_CODE)
End Function

Sub setCurrentFeed(newFeed As Integer)
	Code SET_FEED_RATE_CODE & newFeed
End Sub


Function getCurrentX()
	getCurrentX = getCurrentCoordinate(X_COORDINATE_NUMBER)
End Function

Function getCurrentY()
	getCurrentY = getCurrentCoordinate(Y_COORDINATE_NUMBER)
End Function

Function getCurrentZ()
	getCurrentZ = getCurrentCoordinate(Z_COORDINATE_NUMBER)
End Function

Function getCurrentCoordinate(ByVal coordinateNumber As Integer)
	getCurrentCoordinate = GetOEMDRO(current_X_OEM_CODE + coordinateNumber)
End Function


Sub setCurrentX(ByVal newX As Double)
	SetDro(X_COORDINATE_NUMBER, newX)
End Sub

Sub setCurrentY(ByVal newY As Double)
	SetDro(Y_COORDINATE_NUMBER, newY)
End Sub

Sub setCurrentZ(ByVal newZ As Double)
	SetDro(Z_COORDINATE_NUMBER, newZ)
End Sub


Sub printNoTouchError(ByVal coordinate As String)
	printMessage("No " & coordinate & " touch")
End Sub

Sub printMessage(ByVal message As String)
	Code "(" & message & ")"
End Sub


Function isZero(ByVal number As Double) As Boolean
	isZero = (number < EPSILON) And (number > -EPSILON)
