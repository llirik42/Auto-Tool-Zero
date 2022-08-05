Const Z_THICKNESS As Double = 8.0

Const MAX_DEPTH As Double = 40.0 

Const SAFE_DISTANCE As Double = 25.0 

Const FEED_RATE As Integer = 400 

Const DELAY = 500

Const EPSILON As Double = 0.001

Const Z_COORDINATE_NUMBER As Integer = 2

Const CURRENT_Z_OEM_CODE As Integer = 802

Const CURRENT_FEED_RATE_OEM_CODE As Integer = 818 

Const SET_FEED_RATE_CODE As String = "F" 

Const FAST_MODE_CODE As String = "G00" 

Const MOVE_UNTILL_TOUCH_OR_REACH_CODE As String = "G31" 


Call main()


Sub main()
	originFeedRate = getCurrentFeedRate()		
	setCurrentFeed(FEED_RATE)

	startPositionZ = getCurrentZ()
	
	zTouch = autoZeroZ(startPositionZ)
	If zTouch = False Then
		Exit Sub
	End If

	Sleep(DELAY)
	
	setCurrentZ(SAFE_DISTANCE + Z_THICKNESS)
End Sub

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


Function reachedMaxDepth(ByVal startPositionZ As Double) as Boolean
	reachedMaxDepth = isZero(getCurrentZ() - startPositionZ + MAX_DEPTH)
End Function


Sub moveFastToZ(zDestination)
	Code FAST_MODE_CODE & "Z" & zDestination
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


Function getCurrentZ() As Double
	getCurrentZ = getCurrentCoordinate(Z_COORDINATE_NUMBER)
End Function

Function getCurrentCoordinate(ByVal coordinateNumber As Integer) As Double
	getCurrentCoordinate = GetOEMDRO(current_Z_OEM_CODE)
End Function

Sub setCurrentZ(ByVal newZ As Double)
	SetDro(Z_COORDINATE_NUMBER, newZ)
End Sub


Sub printNoTouchErrorZ()
	Message("No Z touch")
End Sub


Function equal(ByVal a As Double, ByVal b As Double) As Boolean
	equal = isZero(a - b)
End Function

Function isZero(ByVal number As Double) As Boolean
	isZero = (number < EPSILON) And (number > -EPSILON)
End Function                                                  
