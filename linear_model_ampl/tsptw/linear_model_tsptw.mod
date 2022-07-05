# Parameters
param n;
param M;
param CoX{i in 1..n};
param CoY{i in 1..n};
param penalite{i in 1..n};
param a {i in 1..n}>=0;
param b {i in 1..n}>=0;

# Preprocessing - Distance Matrix 
param dist{i in 1..n, j in 1..n}:= if (i=j) then 100000  else sqrt((CoX[i]-CoX[j])^2 +(CoY[i]-CoY[j])^2 );

# Variables
var X{ i in 1..n, j in 1..n} binary;
var Y{i in 1..n} binary;
var t {i in 1..n},>=0;
var k {i in 1..n},>=0; 
var R,>=0;

# Objective Function 
minimize Distance : sum{ i in 1..n,j in 1..n} dist[i,j]*X[i,j]+ sum {i in 2..n} (penalite[i]*(k[i]-b[i])); 

# s.t.

# Customers visit constraints
Contrainte1 {j in 1..n}: (sum{i in 1..n} X[i,j])==1 ; 
Contrainte2 {j in 1..n}: sum{i in 1..n} X[i,j]- sum{i in 1..n} X[j,i]==0; 

# Time window constraints
Contrainte3{j in 1..n}: t[j]>=a[j]; 
Contrainte4{i in 1..n, j in 2..n}:t[j]>= t[i]+ dist[i,j]- M*(1-X[i,j]);

# Lateness penalties constraints 
Contrainte5{i in 2..n}: t[i]-b[i]<=M*Y[i];
Contrainte6{i in 2..n}: b[i]-t[i]<=M*(1-Y[i]);
Contrainte7{i in 2..n}: k[i]>=t[i];
Contrainte8 {i in 2..n}: k[i]>=(1-Y[i])*b[i];

solve;

printf :"\n\nThe optimal solution for this TSPTW has %s Km \n\n", sum{ i in 1..n,j in 1..n} dist[i,j]*X[i,j]+ sum {i in 2..n} (penalite[i]*(k[i]-b[i]));
printf :"The route found and the times of visit are share below: \n\n";
printf{i in 1..n, j in 1..n: X[i,j]==1} : "%s ->%s, t[%s]=%s\n",i,j,j,t[j];
printf:"\n\n";
printf {i in 2..n: Y[i]==0}: "Customer %s was visited on time\n",i;
printf {i in 2..n: Y[i]==1}: "\n\nCustomer %s was visited late by %s min \n\n",i,k[i]-b[i];
printf:"\n\nEND\n\n";

data;

param M:= 100000;
param n:=15;
param : CoX  CoY  a  b  penalite := 
1	17	33	0	10000	0
2	12	32	750	975	0.5
3	17	45	300	400	0.5
4	21	38	650	750	0.5
5	8	9	700	800	0.5
6	38	24	400	550	0.5
7	47	34	350	575	0.5
8	43	11	750	1000	0.5
9	31	28	600	875	0.5
10	10	44	550	700	0.5
11	44	18	400	625	0.5
12	19	21	350	450	0.5
13	23	32	600	875	0.5
14	6	2	250	375	0.5
15	47	3	550	825	0.5;
end;

#additional data
#16	40	24	400	700	0.5
#17	36	47	550	800	0.5
#18	34	19	450	725	0.5
#19	16	3	350	475	0.5
#20	31	42	700	925	0.5
#21	40	32	400	625	0.5
#22	44	39	500	800	0.5
#23	50	40	450	725	0.5
#24	42	12	550	650	0.5
#25	34	48	500	775	0.5
#26	28	11	500	600	0.5
#27	15	47	450	675	0.5
#28	17	11	50	300	0.5
#29	48	43	300	475	0.5
#30	1	12	150	450	0.5
#31	37	42	300	475	0.5
#32	17	13	50	250	0.5
#33	30	24	250	525	0.5
#34	33	47	50	225	0.5
#35	47	27	150	275	0.5
#36	24	18	450	750	0.5
#37	39	1	450	725	0.5
#38	25	8	450	650	0.5
#39	48	44	350	500	0.5
#40	24	19	625	750	0.5;