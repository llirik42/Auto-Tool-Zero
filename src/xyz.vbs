'Ширина для X, Y и Z соответственно. Нужна при выставления нулей координат
Const X_THICKNESS As Double = 15.0 
Const Y_THICKNESS As Double = 15.0
Const Z_THICKNESS As Double = 8.0

'Насколько далеко фреза уйдёт от начального положения по X, чтоб потом идти до касания 
Const X_WITHDRAWAL As Double = 30.0 

'Насколько далеко фреза уйдёт от начального положения по Y, чтоб потом идти до касания 
Const Y_WITHDRAWAL As Double = 30.0

'Насколько фреза может опуститься от начального положения при движении до касания
Const MAX_DEPTH As Double = 40.0 

'Безопасное расстояние, на котором фреза держится от детали, когда фрезе нужно быстро переместиться
Const SAFE_DISTANCE As Double = 25.0 

'Радиус инструмента. Мы его учитываем при выставление нулей координат по X и Y
Const TOOL_RADIUS As Double = 1.5875 

'Задержка, без которой сразу после быстрого движения фреза не сможет двигаться до касания
Const DELAY As Integer = 1000

'Скорость подачи во время время движения до касания
Const FEED_RATE As Integer = 400 

'Допуск. Нужен сразу после движений до касания.
'Если текущее положение мало (меньше чем на эпсилон) отличается от точки, в которую мы хотели прийти, то значит мы не коснулись.
'Он не влияет на остановку во время движения до касания (за это отвечает сам станок, можно задать лишь точку, в которую нужно двигаться)
Const EPSILON As Double = 0.001

'Координатные номера координат. Нужны для выставления и получения текущих координат
Const X_COORDINATE_NUMBER As Integer = 0
Const Y_COORDINATE_NUMBER As Integer = 1
Const Z_COORDINATE_NUMBER As Integer = 2

'OEM-код для получения текущей X-координаты. С помощью координатных номеров координат и него получаются текущие Y и Z. 
Const CURRENT_X_OEM_CODE As Integer = 800

'OEM-код для получения текущей скорости подачи
Const CURRENT_FEED_RATE_OEM_CODE As Integer = 818 

'Код для выставления скорости подачи
Const SET_FEED_RATE_CODE As String = "F" 

'Код для быстрого движения
Const FAST_MODE_CODE As String = "G00" 

'Код для движения с текущей скорости до точки или до касания
Const MOVE_UNTILL_TOUCH_OR_REACH_CODE As String = "G31" 


'Вызов главной функции
Call main() 


'Главная функция, в которой всё происходит
Sub main()
	originFeedRate = getCurrentFeedRate()		
	setCurrentFeed(FEED_RATE)

	Call setZeroes()
	
	setCurrentFeed(originFeedRate)
End Sub

'Функция, в которой происходит всё выставление
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
	
	'Изначально 0 по Z выставляется без учёта толщины, а после данного шага 0 по Z будет в нужном месте
	setCurrentZ(SAFE_DISTANCE + Z_THICKNESS)	
End Sub


'Все функция выставления возвращают True, если выставление прошло успешно и False если касания не произошло

'Выставление по X
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

'Выставление по Y
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

'Выставление по Z
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


'Достигла ли фреза максимальной глубины
Function reachedMaxDepth(ByVal startPositionZ As Double) As Boolean
	reachedMaxDepth = isZero(getCurrentZ() - startPositionZ + MAX_DEPTH)
End Function


'While IsMoving во всех функциях движения нужно, чтобы программа вышла из функции только при завершении движения.
'Иначе программа будет исполняться во время движения, и, например, координаты, которые нужно было получить после движения будут получены во время

'Эти 2 функции отвечают за быстрое движение по какой-либо из осей.
'Быстрое означает, что движение всегда с одной и той же высокой скоростью.
'При этом эти функции отвечают за движение на некоторое число. То есть на 13.0 вправо, а не в точку с координатой 13.0 по X (например)
Sub moveFastOnX(dx)
	moveFastToX(getCurrentX() + dx)
End Sub

Sub moveFastOnY(dy)
	moveFastToY(getCurrentY() + dy)
End Sub


'Эти 3 функции отвечают за быстрое движение по какой-либо из осей.
'Быстрое означает, что движение всегда с одной и той же высокой скоростью.
'При этом эти функции отвечают за движение в точку, то есть не на 5.0 вправо, а в точку с координатой 5.0 по X (например)
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


'Эти 3 функции отвечают за движение в точку до касания или достижения точки по какой-либо из осей
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


'Получить текущую скорость подачи
Function getCurrentFeedRate() As Integer
	getCurrentFeedRate = GetOEMDRO(CURRENT_FEED_RATE_OEM_CODE)
End Function

'Установить скорости подачи
Sub setCurrentFeed(newFeed As Integer)
	Code SET_FEED_RATE_CODE & newFeed
End Sub


'Эти 3 функции отвечают за получение текущей координаты по какой-то из осей
Function getCurrentX() As Double
	getCurrentX = getCurrentCoordinate(X_COORDINATE_NUMBER)
End Function

Function getCurrentY() As Double
	getCurrentY = getCurrentCoordinate(Y_COORDINATE_NUMBER)
End Function

Function getCurrentZ() As Double
	getCurrentZ = getCurrentCoordinate(Z_COORDINATE_NUMBER)
End Function

'Получение значения координаты по координатному номеру. 
'OEM-коды для X, Y и Z расположены подряд, так что <OEM-код для Z> = <OEM-код для Y> + 1 = <OEM-код для X> + 2
Function getCurrentCoordinate(ByVal coordinateNumber As Integer) As Double
	getCurrentCoordinate = GetOEMDRO(current_X_OEM_CODE + coordinateNumber)
End Function


'Эти 3 функции отвечают за установку текущей координаты по какой-то из осей
Sub setCurrentX(ByVal newX As Double) 
	SetDro(X_COORDINATE_NUMBER, newX)
End Sub

Sub setCurrentY(ByVal newY As Double)
	SetDro(Y_COORDINATE_NUMBER, newY)
End Sub

Sub setCurrentZ(ByVal newZ As Double)
	SetDro(Z_COORDINATE_NUMBER, newZ)
End Sub


'Вывести сообщение о том, что касания по X не произошло
Sub printNoTouchErrorX()
	printNoTouchError("X")
End Sub

'Вывести сообщение о том, что касания по Y не произошло
Sub printNoTouchErrorY()
	printNoTouchError("Y")
End Sub

'Вывести сообщение о том, что касания по Z не произошло
Sub printNoTouchErrorZ()
	printNoTouchError("Z")
End Sub

'Вывести сообщение о том, что касания по какой-то координате не произошло. Принимает на вход название координаты "X", "Y" или "Z"
Sub printNoTouchError(ByVal coordinate As String)
	Message("No " & coordinate & " touch")
End Sub


'Проверяет величины на равенство (с допуском)
Function equal(ByVal a As Double, ByVal b As Double) As Boolean
	equal = isZero(a - b)
End Function

'По модулю сравнивает число с нулём, используя "EPSILON". True - число по модулю меньше EPSILON, FALSE - иначе
Function isZero(ByVal number As Double) As Boolean
	isZero = (number < EPSILON) And (number > -EPSILON)
End Function                                                   
