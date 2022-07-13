Const EPSILON As Double = 0.01

Const CURRENT_X_OEM_CODE As Integer = 800

Const X_COORDINATE_NUMBER As Integer = 0
Const Y_COORDINATE_NUMBER As Integer = 1
Const Z_COORDINATE_NUMBER As Integer = 2

Const CURRENT_FEED_RATE_OEM_CODE As Integer = 818

Const SET_FEED_RATE_CODE As String = "F"
Const FAST_MODE_CODE As String = "G00"
const MOVE_UNTILL_TOUCH_OR_REACH_CODE As String = "G31"

Const X_SIZE As Double = 50.0
const Y_SIZE As Double = 60.0

Const X_THICKNESS As Double = 5.0
Const Y_THICKNESS As Double = 6.0
Const Z_THICKNESS as Double = 10.0

Const MAX_DEPTH As Double = 50.0
Const OFFSET as Double = 25.0


Call main()


Sub main()
	originFeedRate = getCurrentFeedRate()
			
	setCurrentFeed(FEED_RATE)

	zTouch = autoZeroZ()
	If zTouch = False Then
		printNoTouchError("z")
		moveFastToZ(MAX_DEPTH)
		Return
	End If
	
	xTouch = autoZeroX()
	If xTouch = False Then
		printNoTouchError("x")
		moveFastOnX(X_SIZE)
		moveFastToZ(MAX_DEPTH)
		moveFastToX(0)
	End If

	yTouch = autoZeroY()
	If yTouch = False Then
		printNoTouchError("y")
		moveFastOnY(Y_SIZE)
		moveFastToZ(MAX_DEPTH)
		moveFastToY(0)
	End If

	setCurrentFeed(originFeedRate)	
End Sub


Function autoZeroX()
	moveFastToX(X_SIZE)

	moveFastToZ(-Z_THICKNESS / 2)

	MoveUntillTouchOrReachX(0)

	If isZero(getCurrentX()) Then
		Return False
	Else
		setCurrentX(X_THICKNESS)
		moveFastOnX(OFFSET)
		moveFastToZ(OFFSET)
		moveFastToX(0)
	End If
End Function

Function autoZeroY()
	moveFastToY(Y_SIZE)

	moveFastToZ(-Z_THICKNESS / 2)

	MoveUntillTouchOrReachY(0)

	If isZero(getCurrentY()) Then
		Return False
	Else
		setCurrentY(Y_THICKNESS)
		moveFastOnY(OFFSET)
		moveFastToZ(OFFSET)
		moveFastToY(0)
	End If
End Function

Function autoZeroZ() As Boolean
	MoveUntillTouchOrReachZ(-MAX_DEPTH)
		
	If reachedMaxDepth() Then
		moveFastToZ(0)
		Return False
	Else 
		setCurrentZ(0)
		moveFastToZ(LIFT)
		return True
	End If
End Function


Function reachedMaxDepth()
	Return isZero(getCurrentZ() - MAX_DEPTH)
End Function


Sub moveFastOnX(dx)
	moveFastToX(getCurrentX() + dx)
End Sub

Sub moveFastOnY(dy)
	moveFastToY(getCurrentY() + dy)
End Sub


Sub moveFastToX(to)
	Code FAST_MODE_CODE & "X" & to
End Sub

Sub moveFastToY(to)
	Code FAST_MODE_CODE & "Y" & to
End Sub

Sub moveFastToZ(to)
	Code FAST_MODE_CODE & "Z" & to
End Sub


Sub MoveUntillTouchOrReachX(to As Double)
	Code MOVE_UNTILL_TOUCH_OR_REACH_CODE & "X" & to
	While IsMoving
	Wend
End Sub

Sub MoveUntillTouchOrReachY(to as Double)
	Code MOVE_UNTILL_TOUCH_OR_REACH_CODE & "Y" & to
	While IsMoving
	Wend
End Sub

Sub MoveUntillTouchOrReachZ(to as Double)
	Code MOVE_UNTILL_TOUCH_OR_REACH_CODE & "Z" & to
	While IsMoving
	Wend
End Sub


Function getCurrentFeedRate()
	Return GetOEMDRO(CURRENT_FEED_RATE_OEM_CODE)
End Function

Sub setCurrentFeed(newFeed as Integer)
	Code SET_FEED_RATE_CODE & newFeed
End Sub


Function getCurrentX()
	Return getCurrentCoordinate(X_COORDINATE_NUMBER)
End Function

Function getCurrentY()
	Return getCurrentCoordinate(Y_COORDINATE_NUMBER)
End Function

Function getCurrentZ()
	Return getCurrentCoordinate(Z_COORDINATE_NUMBER)
End Function

Function getCurrentCoordinate(ByVal coordinateNumber as Integer)
	Return GetOEMDRO(current_X_OEM_CODE + coordinateNumber)
End Function


Sub setCurrentX(ByVal newX As Double)
	SetDro(X_COORDINATE_NUMBER, newX)
End Sub

Sub setCurrentY(ByVal newY As Double)
	SetDro(Y_COORDINATE_NUMBER, newZ)
End Sub

Sub setCurrentZ(ByVal newZ As Double)
	SetDro(Z_COORDINATE_NUMBER, newZ)
End Sub


Sub printNoTouchError(ByVal coordinate as String)
	printMessage("No" & coordinate & "touch")
End Sub

Sub printMessage(ByVal message As String)
	Code "(" & message & ")"
End Sub


Function isZero(ByVal number As Double)
	Return (number < EPSILON) and (number > -EPSILON)
End Function
