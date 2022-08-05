Const X_THICKNESS As Double = 15.0 
Const Y_THICKNESS As Double = 15.0
Const Z_THICKNESS As Double = 8.0

Const X_WITHDRAWAL As Double = 30.0 

Const Y_WITHDRAWAL As Double = 30.0

Const MAX_DEPTH As Double = 40.0 

Const SAFE_DISTANCE As Double = 25.0 

Const TOOL_RADIUS As Double = 1.5875 

Const DELAY As Integer = 1000

Const FEED_RATE As Integer = 400 

Const EPSILON As Double = 0.001

Const X_COORDINATE_NUMBER As Integer = 0
Const Y_COORDINATE_NUMBER As Integer = 1
Const Z_COORDINATE_NUMBER As Integer = 2

Const CURRENT_X_OEM_CODE As Integer = 800

Const CURRENT_FEED_RATE_OEM_CODE As Integer = 818 

Const SET_FEED_RATE_CODE As String = "F" 

Const FAST_MODE_CODE As String = "G00" 

Const MOVE_UNTILL_TOUCH_OR_REACH_CODE As String = "G31" 


Call main()


Sub main()
	originFeedRate = getCurrentFeedRate()		
	setCurrentFeed(FEED_RATE)

	Call setZeroes()
	
	setCurrentFeed(originFeedRate)
End Sub

Sub setZeroes()
	startPositionX = getCurrentX()
	startPositionY = getCurrentY()
	startPositionZ = getCurrentZ()
	
	zTouch = autoZeroZ(startPositionZ)
	If zTouch = False Then
		Exit Sub
	End If
	
	xTouch = autoZeroX(startPositionX)
	If xTouch = False Then	
		Exit Sub
	End If
		
	yTouch = autoZeroY(startPositionY)
	If yTouch = False Then
		Exit Sub
	End If
	
	setCurrentZ(SAFE_DISTANCE + Z_THICKNESS)	
End Sub


Function autoZeroX(ByVal startPositionX As Double) As Boolean
	moveFastOnX(-X_WITHDRAWAL)
	
	moveFastToZ(-Z_THICKNESS / 2)
		
	Sleep(DELAY)
	
	moveUntillTouchOrReachX(startPositionX)
	
	If equal(getCurrentX(), startPositionX) Then
		moveFastOnX(-X_WITHDRAWAL)
		moveFastToZ(SAFE_DISTANCE)
		moveFastToX(startPositionX)
		Call printNoTouchErrorX()
		autoZeroX = False
	Else
		setCurrentX(-X_THICKNESS - TOOL_RADIUS)
		moveFastOnX(-SAFE_DISTANCE)
		moveFastToZ(SAFE_DISTANCE)
		moveFastToX(0)
		autoZeroX = True
	End If
End Function

Function autoZeroY(ByVal startPositionY As Double) As Boolean
	moveFastOnY(-Y_WITHDRAWAL)
	
	moveFastToZ(-Z_THICKNESS / 2)
		
	Sleep(DELAY)
	
	moveUntillTouchOrReachY(startPositionY)
	
	If equal(getCurrentY(), startPositionY) Then
		moveFastOnY(-Y_WITHDRAWAL)
		moveFastToZ(SAFE_DISTANCE)
		moveFastToY(startPositionY)
		Call printNoTouchErrorY()	
		autoZeroY = False
	Else
		setCurrentY(-Y_THICKNESS - TOOL_RADIUS)
		moveFastOnY(-SAFE_DISTANCE)
		moveFastToZ(SAFE_DISTANCE)
		moveFastToY(0)
		autoZeroY = True
	End If

End Function

Function autoZeroZ(ByVal startPositionZ As Double) As Boolean			
	moveUntillTouchOrReachZ(startPositionZ - MAX_DEPTH)
		
	If reachedMaxDepth(startPositionZ) Then
		moveFastToZ(startPositionZ)
		Call printNoTouchErrorZ()
		autoZeroZ = False
	Else 
		setCurrentZ(0)
		moveFastToZ(SAFE_DISTANCE)
		autoZeroZ = True
	End If
End Function


Function reachedMaxDepth(ByVal startPositionZ As Double) As Boolean
	reachedMaxDepth = isZero(getCurrentZ() - startPositionZ + MAX_DEPTH)
End Function


Sub moveFastOnX(dx)
	moveFastToX(getCurrentX() + dx)
End Sub

Sub moveFastOnY(dy)
	moveFastToY(getCurrentY() + dy)
End Sub


Sub moveFastToX(xDestination)
	Code FAST_MODE_CODE & "X" & xDestination
	While  IsMoving()
	Wend
End Sub

Sub moveFastToY(yDestination)
	Code FAST_MODE_CODE & "Y" & yDestination
	While IsMoving()
	Wend
End Sub

Sub moveFastToZ(zDestination)
	Code FAST_MODE_CODE & "Z" & zDestination
	While IsMoving()
	Wend
End Sub


Sub moveUntillTouchOrReachX(xDestination As Double)
	Code MOVE_UNTILL_TOUCH_OR_REACH_CODE & "X" & xDestination
	While IsMoving()
	Wend
End Sub

Sub moveUntillTouchOrReachY(yDestination As Double)
	Code MOVE_UNTILL_TOUCH_OR_REACH_CODE & "Y" & yDestination
	While IsMoving()
	Wend
End Sub

Sub moveUntillTouchOrReachZ(zDestination As Double)
	Code MOVE_UNTILL_TOUCH_OR_REACH_CODE & "Z" & zDestination
	While IsMoving()
	Wend
End Sub


Function getCurrentFeedRate() As Integer
	getCurrentFeedRate = GetOEMDRO(CURRENT_FEED_RATE_OEM_CODE)
End Function

Sub setCurrentFeed(newFeed As Integer)
	Code SET_FEED_RATE_CODE & newFeed
End Sub


Function getCurrentX() As Double
	getCurrentX = getCurrentCoordinate(X_COORDINATE_NUMBER)
End Function

Function getCurrentY() As Double
	getCurrentY = getCurrentCoordinate(Y_COORDINATE_NUMBER)
End Function

Function getCurrentZ() As Double
	getCurrentZ = getCurrentCoordinate(Z_COORDINATE_NUMBER)
End Function

Function getCurrentCoordinate(ByVal coordinateNumber As Integer) As Double
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


Sub printNoTouchErrorX()
	printNoTouchError("X")
End Sub

Sub printNoTouchErrorY()
	printNoTouchError("Y")
End Sub

Sub printNoTouchErrorZ()
	printNoTouchError("Z")
End Sub

Sub printNoTouchError(ByVal coordinate As String)
	Message("No " & coordinate & " touch")
End Sub


Function equal(ByVal a As Double, ByVal b As Double) As Boolean
	equal = isZero(a - b)
End Function

Function isZero(ByVal number As Double) As Boolean
	isZero = (number < EPSILON) And (number > -EPSILON)
End Function                                                   
