Attribute VB_Name = "GA"
Option Explicit
Option Base 1

'****************************
'Global Variables
'****************************
Dim d() As Double 'Distance Matrix
Dim t() As Double 'Travel Time Matrix'
Dim DD() As Integer 'Due Date'
Dim SD() As Integer 'Start Date'
Dim Population() As Integer  'All Population
Dim Solution() As Integer 'Final Solution
Dim CostSolution As Double ' Cost of Final Solution
Dim n As Integer ' Graph Size
Dim m As Integer ' Population Size
Dim proba_mut As Double ' Probability of mutation
Dim Speed As Double ' Average constant speed
Dim k As Integer ' Tournament Selection Size
Dim alpha As Double ' Lateness Cost
Dim SizeNextGen As Integer ' Size of next population
'****************************

Sub lecture()

Dim i As Integer
Dim j As Integer

n = Worksheets("AG").Range("B5").Value
m = Worksheets("AG").Range("B6").Value
alpha = Worksheets("AG").Range("B8").Value
Speed = Worksheets("AG").Range("B9").Value
proba_mut = 0.05
k = 2
SizeNextGen = 2 * m
CostSolution = 1000000

ReDim d(n, n)
ReDim t(n, n)
ReDim Solution(n)
ReDim DD(n)
ReDim SD(n)

For i = 1 To n
    DD(i) = Worksheets("DATA").Range("F2").Offset(i, 0).Value
    SD(i) = Worksheets("DATA").Range("E2").Offset(i, 0).Value
    For j = 1 To n
        t(i, j) = (Worksheets("DIST").Range("A1").Offset(i, j).Value / Speed) * 60
        d(i, j) = Worksheets("DIST").Range("A1").Offset(i, j).Value
    Next j
Next i

End Sub

Function randomized_two_closest_dist(r As Integer, visited() As Boolean) As Integer

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
    randomized_two_closest_dist = v1
Else: randomized_two_closest_dist = v2
End If

If v2 = 0 Then randomized_two_closest_dist = v1

End Function

Function randomized_two_closest_DD(r As Integer, visited() As Boolean) As Integer

Dim i As Integer
Dim v1 As Integer
Dim v2 As Integer
Dim distref As Double

distref = 10000
For i = 1 To n
    If DD(i) <= distref And visited(i) = False Then
        distref = DD(i)
        v1 = i
    End If
Next i

distref = 10000
For i = 1 To n
    If DD(i) <= distref And visited(i) = False And i <> v1 Then
        distref = DD(i)
        v2 = i
    End If
Next i

i = Int(2 * Rnd) + 1
If i = 1 Then
    randomized_two_closest_DD = v1
Else: randomized_two_closest_DD = v2
End If

If v2 = 0 Then randomized_two_closest_DD = v1

End Function

Function randomized_two_closest_distalphaDD(r As Integer, visited() As Boolean) As Integer

Dim i As Integer
Dim v1 As Integer
Dim v2 As Integer
Dim distref As Double

distref = 10000
For i = 1 To n
    If d(r, i) + alpha * DD(i) <= distref And visited(i) = False Then
        distref = d(r, i) + alpha * DD(i)
        v1 = i
    End If
Next i

distref = 10000
For i = 1 To n
    If d(r, i) + alpha * DD(i) <= distref And visited(i) = False And i <> v1 Then
        distref = d(r, i) + alpha * DD(i)
        v2 = i
    End If
Next i

i = Int(2 * Rnd) + 1
If i = 1 Then
    randomized_two_closest_distalphaDD = v1
Else: randomized_two_closest_distalphaDD = v2
End If

If v2 = 0 Then randomized_two_closest_distalphaDD = v1

End Function

Sub Init(InitSolution() As Integer)

Dim i As Integer
Dim RndChoice As Double
Dim visited() As Boolean
ReDim visited(n)

visited(1) = True
For i = 2 To n
    visited(i) = False
Next i

InitSolution(1) = 1
RndChoice = Int(100 * Rnd) + 1
If RndChoice < 40 Then
    For i = 1 To n - 1
        InitSolution(i + 1) = randomized_two_closest_distalphaDD(i, visited)
        visited(InitSolution(i + 1)) = True
    Next i
Else
    If RndChoice > 70 Then
        For i = 1 To n - 1
            InitSolution(i + 1) = randomized_two_closest_dist(i, visited)
            visited(InitSolution(i + 1)) = True
        Next i
    Else
        For i = 1 To n - 1
            InitSolution(i + 1) = randomized_two_closest_DD(i, visited)
            visited(InitSolution(i + 1)) = True
        Next i
    End If
End If

End Sub

Sub Init_Best_Insertion_Randomized(InitSolution() As Integer)

Dim i As Integer, j As Integer
Dim visited() As Boolean
ReDim visited(n)

Dim MaxGain As Integer
Dim NewPos As Integer
Dim MaxPos As Integer
Dim Gain As Integer
Dim Memory As Integer

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
    MaxGain = 0
    For NewPos = 2 To i + 1
        Gain = (d(InitSolution(NewPos - 1), InitSolution(NewPos)) + d(InitSolution(i), InitSolution(i + 1))) - (d(InitSolution(NewPos - 1), InitSolution(i + 1)) + d(InitSolution(i + 1), InitSolution(NewPos)))
        If Gain > MaxGain Then
            MaxGain = Gain
            MaxPos = NewPos
        End If
    Next NewPos
    
    Memory = InitSolution(i + 1)
    For j = i + 1 To MaxPos + 1 Step -1
        InitSolution(j) = InitSolution(j - 1)
    Next j
    InitSolution(MaxPos) = Memory
    visited(InitSolution(MaxPos)) = True
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

Function CostEval(InitSolution() As Integer) As Double   'Routing Cost Evaluation

Dim i As Integer
Dim Cost As Double
Dim Time As Double

Cost = 0
Time = 0
For i = 1 To n - 1
    Time = Time + t(InitSolution(i), InitSolution(i + 1))
    If Time < SD(InitSolution(i + 1)) Then
        Time = SD(InitSolution(i + 1))
    End If
    If (Time - DD(InitSolution(i + 1))) > 0 Then
        Cost = Cost + d(InitSolution(i), InitSolution(i + 1)) + alpha * (Time - DD(InitSolution(i + 1)))
    Else
        Cost = Cost + d(InitSolution(i), InitSolution(i + 1))
    End If
Next i

Time = Time + t(InitSolution(n), InitSolution(1))
If (Time - DD(1)) > 0 Then
    Cost = Cost + d(InitSolution(n), InitSolution(1)) + alpha * (Time - DD(InitSolution(1)))
Else
    Cost = Cost + d(InitSolution(n), InitSolution(1))
End If

CostEval = Cost

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

Function TournamentSelection(Population() As Integer, Cost() As Double, Size As Integer) As Integer

Dim i As Integer
Dim Players As Integer
Dim Winner As Integer
Dim PlayerTab() As Integer

'CREATIVE PARAMETER
Players = k
ReDim PlayerTab(Players)

For i = 1 To Players
    PlayerTab(i) = Int((Size) * Rnd) + 1
Next i

Winner = PlayerTab(1)
For i = 2 To Players
    If Cost(PlayerTab(i)) < Cost(Winner) Then
        Winner = PlayerTab(i)
    End If
Next i

TournamentSelection = Winner

End Function

Sub SimpleCrossing(ByVal Filled As Integer, Parent1 As Integer, Parent2 As Integer, Population() As Integer, Cost() As Double)

Dim Child1() As Integer
Dim Child2() As Integer
Dim p As Integer 'Random point genetic cut
Dim i As Integer
Dim j As Integer
Dim Transmitted() As Boolean 'Store all the gen already transmitted to the current child

ReDim Child1(n)
ReDim Child2(n)
ReDim Transmitted(n)

'INIT

Filled = m + 2 * Filled  'State of generation of the next gen
p = Int((n) * Rnd) + 1   'Random point genetic cut

'CHILD1

For i = 1 To n 'Init
    Transmitted(i) = False
Next i

For i = 1 To p    'Genetic Transmission up to random point p
    Child1(i) = Population(Parent1, i)
    Transmitted(Population(Parent1, i)) = True
Next i

k = p + 1
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
    
For i = 1 To p    'Genetic Transmission up torandom point p
    Child2(i) = Population(Parent2, i)
    Transmitted(Population(Parent2, i)) = True
Next i

k = p + 1
For i = 1 To n     'Genetic Non-Transmission from p random point
        If Transmitted(Population(Parent1, i)) = False Then
            Child2(k) = Population(Parent1, i)
            k = k + 1
        End If
Next i

'COST UPDATE

Cost(Filled - 1) = CostEval(Child1)
Cost(Filled) = CostEval(Child2)

'STORAGE

For i = 1 To n
    Population(Filled - 1, i) = Child1(i)
    Population(Filled, i) = Child2(i)
Next i

End Sub


Sub Selection(Cost() As Double)

Dim TBS1 As Integer 'To Be Selected
Dim TBS2 As Integer 'To Be Selected
Dim Best As Integer
Dim Selected As Integer
Dim RndChoice As Integer
Dim i As Integer
Dim PopulationTemp() As Double
Dim ToCostEval() As Integer

'INIT

ReDim PopulationTemp(m, n + 1)
ReDim ToCostEval(n)

Selected = 0
While Selected < m
    TBS1 = TournamentSelection(Population, Cost, SizeNextGen)
    TBS2 = TBS1
    While TBS2 = TBS1
        TBS2 = TournamentSelection(Population, Cost, SizeNextGen)
    Wend
    RndChoice = Int(100 * Rnd) + 1
    If Cost(TBS1) <= Cost(TBS2) Then
        If RndChoice < 5 Then
            Best = TBS2
        Else
            Best = TBS1
        End If
    Else
        If RndChoice < 5 Then
            Best = TBS1
        Else
            Best = TBS2
        End If
    End If
    For i = 1 To n
        PopulationTemp(Selected + 1, i) = Population(Best, i)
        ToCostEval(i) = Population(Best, i)
    Next i
    PopulationTemp(Selected + 1, n + 1) = CostEval(ToCostEval)
    Selected = Selected + 1
Wend

'OVERWRITE

Call Overwrite(PopulationTemp, Cost)
Erase PopulationTemp

End Sub

Sub Overwrite(PopulationTemp() As Double, Cost() As Double)

Dim i As Integer
Dim j As Integer

Erase Population
ReDim Population(SizeNextGen, n)  'Population Muted Storage

For i = 1 To m
    For j = 1 To n
        Population(i, j) = PopulationTemp(i, j)
    Next j
    Cost(i) = PopulationTemp(i, n + 1)
Next i

Call SaveBest(Population, Cost)

End Sub

Sub SaveBest(Population() As Integer, Cost() As Double)

Dim i As Integer
Dim j As Integer

For i = 1 To SizeNextGen
    If Cost(i) < CostSolution And Population(i, 1) <> 0 Then
        CostSolution = Cost(i)
        For j = 1 To n
            Solution(j) = Population(i, j)
        Next j
    End If
Next i

End Sub

Sub WriteDown(Population() As Integer, Cost() As Double)

Dim j As Integer
Dim i As Integer

Worksheets("AG").Activate
Range("B14", "CY113").ClearContents
For j = 1 To m
    For i = 1 To n
        Range("B14").Offset(j - 1, i) = Population(j, i)
    Next i
Range("B14").Offset(j - 1, 0) = Cost(j)
Next j

End Sub

Sub WriteDownFirst(Population() As Integer, Cost() As Double)

Dim j As Integer
Dim i As Integer

Worksheets("FIRST_POP").Activate
Range("B2", "CY101").ClearContents
For j = 1 To m
    For i = 1 To n
        Range("B2").Offset(j - 1, i) = Population(j, i)
    Next i
Range("B2").Offset(j - 1, 0) = Cost(j)
Next j

End Sub

Sub WriteDownBest(Solution() As Integer, CostSolution As Double)

Dim i As Integer

Worksheets("AG").Activate
Range("N2", "CY2").ClearContents
For i = 1 To n
    Worksheets("AG").Range("A2").Offset(0, 2 + i - 1) = Solution(i)
Next i
Worksheets("AG").Range("A2").Offset(0, 2 + n) = 1
Worksheets("AG").Range("H7") = CostSolution

End Sub

Sub TwoOptSwap(InitSolution() As Integer)

Dim Max_Improve As Double, Delta As Double, Deltatime As Double, Lateness1 As Double, Lateness2 As Double, d_ab_cd As Double, d_ac_bd As Double, Time As Double
Dim TimeVisit() As Double
Dim Temp As Integer, a As Integer, b As Integer, c As Integer, dback As Integer, i As Integer, j As Integer, k As Integer, aMax As Integer, cMax As Integer, Opt As Integer

Opt = 0
d_ab_cd = 0
d_ac_bd = 0

ReDim TimeVisit(n)
ReDim NewTimeVisit(n)
TimeVisit(1) = 0
'Current Time_Visit calculation for each node
For k = 2 To n
    TimeVisit(InitSolution(k)) = TimeVisit(InitSolution(k - 1)) + t(InitSolution(k - 1), InitSolution(k))
    If TimeVisit(InitSolution(k)) < SD(InitSolution(k)) Then
        TimeVisit(InitSolution(k)) = SD(InitSolution(k))
    End If
Next k

Do
    Opt = Opt + 1
    Max_Improve = 0
    For i = 1 To n - 2
        'Check to nodes next to each other
        a = InitSolution(i)
        b = InitSolution(i + 1)
        'Begin swap plus or mines value calculation for all nodes j (c) scheduled later
        For j = i + 2 To n - 1
            'Check to nodes next to each other
            c = InitSolution(j)
            dback = InitSolution(j + 1)
            'Distance calculation for the plus or mines distance calculation
            d_ab_cd = d(a, b) + d(c, dback)
            d_ac_bd = d(a, c) + d(b, dback)
            Delta = d_ac_bd - d_ab_cd
            'Current Lateness calculation
            Lateness1 = 0
            For k = i + 1 To n
                'Lateness penalization
                If (TimeVisit(InitSolution(k)) - DD(InitSolution(k))) > 0 Then
                    Lateness1 = Lateness1 + alpha * (TimeVisit(InitSolution(k)) - DD(InitSolution(k)))
                End If
            Next k
            If (TimeVisit(InitSolution(n)) + t(InitSolution(n), InitSolution(1)) - DD(InitSolution(1))) > 0 Then
                Lateness1 = Lateness1 + alpha * (TimeVisit(InitSolution(n)) + t(InitSolution(n), InitSolution(1)) - DD(InitSolution(1)))
            End If
            Time = TimeVisit(a)
            'Lateness calculation in case of a and c swap move
            Lateness2 = 0
            For k = j To i + 1 Step -1
                'Time implementation
                If k = j Then
                    Time = Time + t(a, InitSolution(k))
                Else
                    Time = Time + t(InitSolution(k + 1), InitSolution(k))
                End If
                If Time < SD(InitSolution(k)) Then
                    Time = SD(InitSolution(k))
                End If
                 'Lateness penalization
                If (Time - DD(InitSolution(k))) > 0 Then
                    Lateness2 = Lateness2 + alpha * (Time - DD(InitSolution(k)))
                End If
            Next k
            'Dont forget to calculate Latenesseven after the d point
            For k = j + 1 To n
                'Time implementation
                If k = j + 1 Then
                    Time = Time + t(InitSolution(i + 1), InitSolution(k))
                Else
                    Time = Time + t(InitSolution(k - 1), InitSolution(k))
                End If
                If Time < SD(InitSolution(k)) Then
                    Time = SD(InitSolution(k))
                End If
                'Lateness penalization
                If (Time - DD(InitSolution(k))) > 0 Then
                    Lateness2 = Lateness2 + alpha * (Time - DD(InitSolution(k)))
                End If
            Next k
            Time = Time + t(InitSolution(n), InitSolution(1))
            If (Time - DD(InitSolution(1))) > 0 Then
                Lateness2 = Lateness2 + alpha * (Time - DD(InitSolution(1)))
            End If
            'Plus or minus LATENESS calculation
            Deltatime = Lateness2 - Lateness1
            Delta = Delta + Deltatime
            'Maximum moves value research
            If -Delta > Max_Improve Then
                Max_Improve = -Delta
                aMax = i
                cMax = j
            End If
        Next j
    Next i
    'If we found a move that can make the objective function improve then :
    If Max_Improve > 0 Then
        'We make the move
        Do While cMax > (aMax + 1)
            Temp = InitSolution(aMax + 1)
            InitSolution(aMax + 1) = InitSolution(cMax)
            InitSolution(cMax) = Temp
            cMax = cMax - 1
            aMax = aMax + 1
        Loop
        TimeVisit(1) = 0
        For k = 2 To n
            TimeVisit(InitSolution(k)) = TimeVisit(InitSolution(k - 1)) + t(InitSolution(k - 1), InitSolution(k))
        Next k
    End If
'We repeat this processif we can still improve but not more than 10 times
Loop While Max_Improve > 0 And Opt < 10

End Sub

Sub main()

Dim InitSolution() As Integer 'Routing Initial created by NearestNeighbours randomized
Dim People() As Integer
Dim Cost() As Double
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
ReDim Cost(SizeNextGen)
ReDim Best(n)

'________________________INITIALISATION_______________________'

For j = 1 To m  'Population of m initial solution

    RndChoice = Int(100 * Rnd) + 1
    If RndChoice < 40 Then
        Call InitRandom(InitSolution)
    Else
        If RndChoice > 70 Then
            Call Init_Best_Insertion_Randomized(InitSolution)
        Else
            Call Init(InitSolution)
        End If
    End If
    
    For i = 1 To n  'Population Write Down
        Population(j, i) = InitSolution(i)
    Next i
    Cost(j) = CostEval(InitSolution)
        
Next j

Application.ScreenUpdating = False

Call WriteDownFirst(Population, Cost)

Worksheets("AG").Activate
Trial = Worksheets("AG").Range("B7").Value

'_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ '
'________________________AG PROCESS_______________________'
'_ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ '

For Iteration = 1 To Trial

    '________________________MUTATION_______________________'
    
    'Increasing Mutation Rate To Avoid Early Convergence'
    
    If Iteration > Trial - (Trial / 10) Then
        proba_mut = 0.4
    Else
        If Iteration > Trial / (1.3) Then
            proba_mut = 0.2
       Else
            If Iteration > Trial / 2 Then
                proba_mut = 0.1
            End If
        End If
      End If
    
    'Mutation Process'
    
    For j = 1 To m
        RndChoice = Int(100 * Rnd) + 1
        If RndChoice < proba_mut * 100 Then
            'Temp 1Dim Copy
            For i = 1 To n
                People(i) = Population(j, i)
            Next i
            'Mutation Type Select' 'Just Random Choice'
            RndChoice = Int(100 * Rnd) + 1
            If RndChoice > 90 Then
                Call MutationReverse(People)
            ElseIf RndChoice < 45 Then
                Call MutationChange(People)
            Else
                Call MutationInsertion(People)
            End If
            RndChoice = Int(100 * Rnd) + 1
            If RndChoice < 4 Then
                Call TwoOptSwap(People)
            End If
            Cost(j) = CostEval(People)
            For i = 1 To n
                Population(j, i) = People(i)
            Next i
        End If
    Next j
    
    '________________________CROSSING_______________________'
    
    'PROCEED CROSSING ((SizeNextGen - m) / 2) TIMES
    
    For i = 1 To ((SizeNextGen - m) / 2)
    
        'PARENT TOURNAMENT SELECTION
        Parent1 = TournamentSelection(Population, Cost, m)
        Parent2 = Parent1
        While Parent2 = Parent1     'Need two different parents to proceed crossing ((SizeNextGen - m) / 2) Times
            Parent2 = TournamentSelection(Population, Cost, m)
        Wend
        
        'SIMPLE CROSSING
        Call SimpleCrossing(i, Parent1, Parent2, Population, Cost)
    
    Next i
    
    '________________________SELECTION_______________________'
    
    Call Selection(Cost)
    
Next Iteration

Call WriteDown(Population, Cost)
Call TwoOptSwap(Solution)
CostSolution = CostEval(Solution)
Call WriteDownBest(Solution, CostSolution)

Application.ScreenUpdating = False

End Sub
