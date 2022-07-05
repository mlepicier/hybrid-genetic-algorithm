Attribute VB_Name = "GA"
Option Explicit
Option Base 1

'****************************
'Variables globales
'****************************
Dim d() As Double 'Distance Matrix
Dim Population() As Integer  'All Population
Dim Solution() As Integer 'Final Solution
Dim dem() As Integer 'Demand List
Dim CostSolution As Double 'Cost of Final Solution
Dim capa As Double 'Truck Capacity
Dim n As Integer ' Graph Size
Dim m As Integer 'Population Size
Dim proba_mut As Double 'Probability of Mutation
Dim k As Integer 'Tournament selection size
Dim SizeNextGen As Integer 'Size of next generation
'****************************

Sub lecture()

Dim i As Integer
Dim j As Integer

n = Worksheets("AG").Range("B5").Value
m = Worksheets("AG").Range("B6").Value
proba_mut = 0.05
k = 2
SizeNextGen = 2 * m
CostSolution = 1000000

ReDim d(n, n)
ReDim Solution(n)
For i = 1 To n
    For j = 1 To n
        d(i, j) = Worksheets("DIST").Range("A1").Offset(i, j).Value
    Next j
Next i

ReDim dem(n)
For i = 1 To n
    dem(i) = Worksheets("DEMAND").Range("A2").Offset(0, i).Value
Next i
capa = Worksheets("DEMAND").Range("B3").Value

End Sub

Function randomized_two_closest(r As Integer, visited() As Boolean) As Integer

Dim i As Integer
Dim v1 As Integer
Dim v2 As Integer
Dim distref As Double

distref = 10000
For i = 1 To n
    If d(r, i) <= distref And visited(i) = False Then
        distref = d(r, i)
        v1 = i
    End If
Next i

distref = 10000
For i = 1 To n
    If d(r, i) <= distref And visited(i) = False And i <> v1 Then
        distref = d(r, i)
        v2 = i
    End If
Next i

i = Int(2 * Rnd) + 1
If i = 1 Then
    randomized_two_closest = v1
Else: randomized_two_closest = v2
End If

If v2 = 0 Then randomized_two_closest = v1

End Function


Sub Init(InitSolution() As Integer)

Dim i As Integer
Dim visited() As Boolean
ReDim visited(n)

visited(1) = True
For i = 2 To n
    visited(i) = False
Next i

InitSolution(1) = 1
For i = 1 To n - 1
    InitSolution(i + 1) = randomized_two_closest(i, visited)
    visited(InitSolution(i + 1)) = True
Next i

End Sub

Sub InitRandom(InitSolution() As Integer)

Dim i As Integer
Dim visited() As Boolean
ReDim visited(n)

visited(1) = True
For i = 2 To n
    visited(i) = False
Next i

InitSolution(1) = 1
For i = 1 To n - 1
    InitSolution(i + 1) = Int(Rnd() * n - 1) + 2
    While visited(InitSolution(i + 1))
        If InitSolution(i + 1) < n Then
            InitSolution(i + 1) = InitSolution(i + 1) + 1
        Else
            InitSolution(i + 1) = 2
        End If
    Wend
    visited(InitSolution(i + 1)) = True
Next i

End Sub

Function CostEval(InitSolution() As Integer) As Double 'Routing Cost Evaluation

Dim i As Integer
Dim cost As Double

cost = 0
For i = 1 To n - 1
    cost = cost + d(InitSolution(i), InitSolution(i + 1))
Next i
cost = cost + d(InitSolution(n), InitSolution(1))

CostEval = cost

End Function

Function Split(tg() As Integer) As Double

'*************************************
' We fill the SPLIT Function
' Breaking down optimally the route
' Returns the breakdown cost
'*************************************

Dim i, j As Integer
Dim V(), P() As Double
ReDim V(n)
ReDim P(n)
Dim Dm  As Integer
Dim cost As Double

V(1) = 0
For i = 2 To n
    V(i) = 9999
Next i

For i = 1 To n
    Dm = 0
    cost = 0
    j = i + 1
    While j <= n And Dm <= capa
        Dm = Dm + dem(tg(j))
        If j = i + 1 Then
            cost = d(tg(1), tg(j)) + d(tg(j), tg(1))
        Else
            cost = cost - d(tg(j - 1), tg(1)) + d(tg(j - 1), tg(j)) + d(tg(j), tg(1))
        End If
        If Dm <= capa Then
            If V(tg(i)) + cost < V(tg(j)) Then
                V(tg(j)) = V(tg(i)) + cost
                'P(i) = j - 1
                P(j) = i
                
            End If
            j = j + 1
        End If
    Wend
Next i

Split = V(tg(n))

End Function

Sub MutationInsertion(InitSolution() As Integer)

Dim Pos As Integer
Dim NewPos As Integer
Dim Memory As Integer
Dim i As Integer

Pos = Int((n - 1) * Rnd) + 2
NewPos = Pos
While NewPos = Pos
    NewPos = Int((n - 1) * Rnd) + 2
Wend

If NewPos < Pos Then
    Memory = Pos
    Pos = NewPos
    NewPos = Memory
End If

Memory = InitSolution(Pos)
For i = Pos To NewPos - 1
    InitSolution(i) = InitSolution(i + 1)
Next i
InitSolution(NewPos) = Memory

End Sub

Sub MutationChange(InitSolution() As Integer)

Dim Pos As Integer
Dim NewPos As Integer
Dim Memory As Integer

Pos = Int((n - 1) * Rnd) + 2
NewPos = Pos
While NewPos = Pos
    NewPos = Int((n - 1) * Rnd) + 2
Wend

Memory = InitSolution(NewPos)
InitSolution(NewPos) = InitSolution(Pos)
InitSolution(Pos) = Memory

End Sub

Sub MutationReverse(InitSolution() As Integer)

Dim PosStart As Integer
Dim PosStop As Integer
Dim Memory As Integer
Dim i As Integer

PosStart = Int((n - 1) * Rnd) + 2
PosStop = PosStart
While PosStop = PosStart
    PosStop = Int((n - 1) * Rnd) + 2
Wend

If PosStop < PosStart Then
    Memory = PosStart
    PosStart = PosStop
    PosStop = Memory
End If

For i = 0 To ((PosStop - PosStart) / 2)
    Memory = InitSolution(PosStart + i)
    InitSolution(PosStart + i) = InitSolution(PosStop - i)
    InitSolution(PosStop - i) = Memory
Next i

End Sub

Function TournamentSelection(Population() As Integer, cost() As Double, size As Integer) As Integer

Dim i As Integer
Dim Players As Integer
Dim Winner As Integer
Dim PlayerTab() As Integer

'CREATIVE PARAMETER

Players = k
ReDim PlayerTab(Players)

For i = 1 To Players
    PlayerTab(i) = Int((size) * Rnd) + 1
Next i

Winner = PlayerTab(1)
For i = 2 To Players
    If cost(PlayerTab(i)) < cost(Winner) Then
        Winner = PlayerTab(i)
    End If
Next i

TournamentSelection = Winner

End Function

Sub SimpleCrossing(ByVal Filled As Integer, Parent1 As Integer, Parent2 As Integer, Population() As Integer, cost() As Double)

Dim Child1() As Integer
Dim Child2() As Integer
Dim P As Integer 'Random point genetic cut
Dim i As Integer
Dim j As Integer
Dim Transmitted() As Boolean 'Store all the gen already transmitted to the current child

ReDim Child1(n)
ReDim Child2(n)
ReDim Transmitted(n)

'INIT

Filled = m + 2 * Filled  'State of generation of the next gen
P = Int((n) * Rnd) + 1   'Random point genetic cut

'CHILD1

For i = 1 To n 'Init
    Transmitted(i) = False
Next i

For i = 1 To P    'Genetic Transmission up to random point p
    Child1(i) = Population(Parent1, i)
    Transmitted(Population(Parent1, i)) = True
Next i

k = P + 1
For i = 1 To n    'Genetic Non-Transmission from p random point
        If Transmitted(Population(Parent2, i)) = False Then
            Child1(k) = Population(Parent2, i)
            k = k + 1
        End If
Next i
     
'CHILD2

For i = 1 To n 'Init
    Transmitted(i) = False
Next i
    
For i = 1 To P    'Genetic Transmission up torandom point p
    Child2(i) = Population(Parent2, i)
    Transmitted(Population(Parent2, i)) = True
Next i

k = P + 1
For i = 1 To n     'Genetic Non-Transmission from p random point
        If Transmitted(Population(Parent1, i)) = False Then
            Child2(k) = Population(Parent1, i)
            k = k + 1
        End If
Next i

'COST UPDATE

cost(Filled - 1) = Split(Child1)
cost(Filled) = Split(Child2)

'STORAGE

For i = 1 To n
    Population(Filled - 1, i) = Child1(i)
    Population(Filled, i) = Child2(i)
Next i

End Sub

Sub Selection(cost() As Double)

Dim TBS1 As Integer 'To Be Selected
Dim TBS2 As Integer 'To Be Selected
Dim Best As Integer
Dim Selected As Integer
Dim i As Integer
Dim PopulationTemp() As Double
Dim ToCostEval() As Integer

'INIT

ReDim PopulationTemp(m, n + 1)
ReDim ToCostEval(n)

Selected = 0
While Selected < m
    TBS1 = TournamentSelection(Population, cost, SizeNextGen)
    TBS2 = TBS1
    While TBS2 = TBS1
        TBS2 = TournamentSelection(Population, cost, SizeNextGen)
    Wend
    If cost(TBS1) <= cost(TBS2) Then
        Best = TBS1
    Else
        Best = TBS2
    End If
    For i = 1 To n
        PopulationTemp(Selected + 1, i) = Population(Best, i)
        ToCostEval(i) = Population(Best, i)
    Next i
    PopulationTemp(Selected + 1, n + 1) = Split(ToCostEval)
    Selected = Selected + 1
Wend

'OVERWRITE

Call Overwrite(PopulationTemp, cost)
Erase PopulationTemp

End Sub

Sub Overwrite(PopulationTemp() As Double, cost() As Double)

Dim i As Integer
Dim j As Integer

Erase Population
ReDim Population(SizeNextGen, n)  'Population Muted Storage

For i = 1 To m
    For j = 1 To n
        Population(i, j) = PopulationTemp(i, j)
    Next j
    cost(i) = PopulationTemp(i, n + 1)
Next i

Call SaveBest(Population, cost)

End Sub

Sub SaveBest(Population() As Integer, cost() As Double)

Dim i As Integer
Dim j As Integer

For i = 1 To SizeNextGen
    If cost(i) < CostSolution And Population(i, 1) <> 0 Then
        CostSolution = cost(i)
        For j = 1 To n
            Solution(j) = Population(i, j)
        Next j
    End If
Next i

End Sub

Sub WriteDown(Population() As Integer, cost() As Double)

Dim j As Integer
Dim i As Integer

Worksheets("AG").Activate
Range("B15", "BB113").ClearContents
For j = 1 To m
    For i = 1 To n
         Range("B15").Offset(j - 1, i) = Population(j, i)
    Next i
Range("B15").Offset(j - 1, 0) = cost(j)
Next j
End Sub

Sub WriteDownBest(Solution() As Integer, CostSolution As Double)

Dim i As Integer

Worksheets("AG").Activate
Range("N2", "BB2").ClearContents
For i = 1 To n
    Worksheets("AG").Range("A2").Offset(0, 2 + i - 1) = Solution(i)
Next i
Worksheets("AG").Range("H7") = CostSolution

Call PrintDecomposition(Solution)

End Sub

Sub PrintDecomposition(tg() As Integer)

Dim i, j As Integer
Dim V(), P() As Double
ReDim V(n)
ReDim P(n)
Dim Dm  As Integer
Dim cost As Double

Range("N5", "BA12").ClearContents

V(1) = 0
For i = 2 To n
    V(i) = 9999
Next i

For i = 1 To n
    Dm = 0
    cost = 0
    j = i + 1
    While j <= n And Dm <= capa
        Dm = Dm + dem(tg(j))
        If j = i + 1 Then
            cost = d(tg(1), tg(j)) + d(tg(j), tg(1))
        Else
            cost = cost - d(tg(j - 1), tg(1)) + d(tg(j - 1), tg(j)) + d(tg(j), tg(1))
        End If
        If Dm <= capa Then
            If V(tg(i)) + cost < V(tg(j)) Then
                V(tg(j)) = V(tg(i)) + cost
                'P(i) = j - 1
                P(j) = i
                
            End If
            j = j + 1
        End If
    Wend
Next i

Dim size, pred As Integer
Dim t, routingsize As Integer
Dim nbtournee As Integer

Worksheets("AG").Activate

size = n
nbtournee = 1
While size > 1
    pred = P(size) + 1
    Range("N4").Offset(nbtournee, 0) = 1
    routingsize = 1
    For t = pred To size
        Range("N4").Offset(nbtournee, routingsize) = tg(t)
        routingsize = routingsize + 1
    Next t
    Range("N4").Offset(nbtournee, routingsize) = 1
    nbtournee = nbtournee + 1
    size = P(size)
Wend
    nbtournee = nbtournee - 1
End Sub

Sub main()

Dim InitSolution() As Integer 'Routing Initial created by NearestNeighbours randomized
Dim People() As Integer
Dim cost() As Double
Dim Best() As Integer
Dim Parent1 As Integer
Dim Parent2 As Integer
Dim Iteration As Integer
Dim Trial As Integer
Dim RndChoice As Double
Dim i As Integer
Dim j As Integer
Randomize

'________________________READING_______________________'

Call lecture
ReDim InitSolution(n)  'Route length
ReDim People(n)
ReDim Population(SizeNextGen, n)  'Population Muted Storage
ReDim cost(SizeNextGen)
ReDim Best(n)

'________________________INITIALISATION_______________________'

For j = 1 To m  'Population of m initial solution by Nearest Neighbours

    RndChoice = Int(100 * Rnd) + 1
    If RndChoice < 50 Then
        Call Init(InitSolution) 'NN
    Else
        Call InitRandom(InitSolution)
    End If
    
    For i = 1 To n  'Population Write Down
        Population(j, i) = InitSolution(i)
    Next i
    cost(j) = Split(InitSolution) 'VRP
        
Next j

Worksheets("AG").Activate
Trial = Worksheets("AG").Range("B7").Value

'_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ '
'________________________AG PROCESS_______________________'
'_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ '

For Iteration = 1 To Trial

    '________________________MUTATION_______________________'
    
    'Increasing Mutation Rate To Avoid Early Convergence'
    If Iteration > Trial / 2 Then
        proba_mut = proba_mut * 2
    End If
    
    If Iteration > Trial / (1.3) Then
        proba_mut = proba_mut * 2
    End If
    
    If Iteration > Trial - (Trial / 5) Then
        proba_mut = 0.4
    End If
    
    'Mutation Process'
    For j = 1 To m
        RndChoice = Int(100 * Rnd) + 1
        If RndChoice < proba_mut * 100 Then
            For i = 1 To n
                People(i) = Population(j, i)
            Next i
            'Mutation Type Select CREATRIVITE' 'Just Random Choice'
            RndChoice = Int(100 * Rnd) + 1
            If RndChoice > 67 Then
                Call MutationReverse(People)
            ElseIf RndChoice < 33 Then
                Call MutationChange(People)
            Else
                Call MutationInsertion(People)
            End If
            cost(j) = Split(People)
            For i = 1 To n
                Population(j, i) = People(i)
            Next i
        End If
    Next j
    
    '________________________CROSSING_______________________'
    
    'PROCEED CROSSING ((SizeNextGen - m) / 2) TIMES
    
    For i = 1 To ((SizeNextGen - m) / 2)
    
        'PARENT TOURNAMENT SELECTION
        Parent1 = TournamentSelection(Population, cost, m)
        Parent2 = Parent1
        While Parent2 = Parent1     'Need two different parents to proceed crossing ((SizeNextGen - m) / 2) Times
            Parent2 = TournamentSelection(Population, cost, m)
        Wend
        
        'SIMPLE CROSSING
        Call SimpleCrossing(i, Parent1, Parent2, Population, cost)
    
    Next i
    
    '________________________SELECTION_______________________'
    
    Call Selection(cost)
    
Next Iteration

Call WriteDown(Population, cost)
Call WriteDownBest(Solution, CostSolution)

End Sub
