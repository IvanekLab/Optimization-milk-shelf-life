''' Created on Mon May 14 10:52:43 2018 @author: forough '''
import xlrd
from gurobipy import *
import math
from numpy import loadtxt
import numpy as np
import time
import heapq 


''' **************************************************************************** ''' 
''' ************************** Reading the data file *************************** '''
''' **************************************************************************** '''       

#startRead = time.time()
diff = 0.000 
PresolverOn = -1      # 0 off, -1 deafult/minimum, 2 minimum 
MyTimeLimit = 3600   # Time limit 
MinGap = 0.0001
ShowMILPStuff = 1   # 0 turns off display of Gurobi output when solving subproblems, 1 otherwise

print ('----------------------------------------------------------------------------')

file_location =                   "M6.xlsx"    
workbook = xlrd.open_workbook   ( "M6.xlsx" )   
print                           ( "M6" ) 
B = 479 # 274, 479, 548, 822 
NoIntervention = 0 # 1: do nothing
MinPR = -0.025 # minimum penalty for the fifth category milk - can change 

MaxPR    = 0.1   # 240 cent / CWT
MaxPRhat = 0.02  # 48 cents / CWT
M = 999999 

''' ************************* Parameters ************************* '''

sheet = workbook.sheet_by_name("Input")     # starts from 0 
Dsize =  int (sheet.cell_value(5,1))   
Drange = range (1, Dsize + 1)   
Wsize = math.ceil (Dsize/7)    
Wrange = range (1, Wsize + 1)    
RPR = {}
l = {}
RPR[1] = 0.2 # happens one time, since the begining
l[1] = Wsize + 1 # s = 1 unit: week
RPR[2] = 0.05 # happens one time every 4 weeks/1 month (after a months) - 5 years: 0.6, 20 yrs: 2.4   
l[2] = 52 # 1 month, 4 weeks  # s = 0
a = {}
a[1] = 100 
a[2] = 100 
#print ("RPR = ", RPR) 
#print ("a = ", a) #print ("l = ", l[2])  
s = {}
s[1] = 1 # 1 in the model (starts from the begining) 
s[2] = 0 # 0 in the model
alpha = 0.7
beta = 0.3

Csize = 5          
Crange = range (1, Csize + 1)     

Cmin    =  {}
Cmin[1] = -4
Cmin[2] = 1.728360
Cmin[3] = 2.30865
Cmin[4] = 2.809897
Cmin[5] = 3.39209    

Cmax    =  {}    
Cmax[1] = 1.72835
Cmax[2] = 2.30864   
Cmax[3] = 2.809896
Cmax[4] = 3.39208
Cmax[5] = 7 
#for c in Crange: #    print (Cmin[c], Cmax[c])
        
CSLsize =  8          
#print ("Category of milk based on its SCs =", CSLsize) 
CSLrange = range (1, CSLsize + 1)          
#print ("Crange =", CSLrange)

CSLmin = {} 
CSLmin[1]  = -14
CSLmin[2]  = -0.30112 
CSLmin[3]  =  0.17611
CSLmin[4]  =  0.81292
CSLmin[5]  =  1.72837
CSLmin[6]  =  2.30858
CSLmin[7]  =  2.80991
CSLmin[8]  =  3.39218

CSLmax = {} 
CSLmax[1]  = -0.30111
CSLmax[2]  =  0.17610
CSLmax[3]  =  0.81291
CSLmax[4]  =  1.72836
CSLmax[5]  =  2.30857
CSLmax[6]  =  2.80990
CSLmax[7]  =  3.39217
CSLmax[8]  =  7


SLmax = 34
SL    = {}
SL[1] = 34
SL[2] = 28
SL[3] = 27
SL[4] = 26
SL[5] = 25
SL[6] = 24
SL[7] = 22
SL[8] = 20

#for c in CSLrange:
#    print (CSLmin[c], CSLmax[c])
# Producer size
Psize =  int (sheet.cell_value(1,1))            
#print ("Producers =", Psize) 
Prange = range (1, Psize+1)   
#print ("Producers =", Prange) 
    
TNP = {}
AllMilkPackages = 0
for d in Drange:
    TNP[d] = round( float ( sheet.cell_value ( 17+3*Psize , d ) ) ) 
    AllMilkPackages += TNP[d]
#print ("TNP = ", TNP)
#print ("AllMilkPackages = ", AllMilkPackages)

# lease - %6 ineterst rate
if TNP[1] <= 3000: # Processing facility <= 4M 
    FCB1 = 86
    VCB1 = 0.00039
    FCB2 = 92
    VCB2 = 0.00078
    FCM = 310
    VCM = 0.00013
elif TNP[1] <= 6000: # Processing facility <= 8M 
    FCB1 = 86
    VCB1 = 0.00039
    FCB2 = 92
    VCB2 = 0.00078
    FCM = 310
    VCM = 0.00013
elif TNP[1] < 30000: # Processing facility = 40M
    FCB1 = 86
    VCB1 = 0.00039
    FCB2 = 107
    VCB2 = 0.00057
    FCM = 310
    VCM = 0.00013
elif TNP[1] < 50000: # Processing facility = 70M
    FCB1 = 99
    VCB1 = 0.00029
    FCB2 = 115
    VCB2 = 0.00038
    FCM = 551
    VCM = 0.00008
elif TNP[1] < 80000: # Processing facility <= 100M
    FCB1 = 99
    VCB1 = 0.00029
    FCB2 = 158
    VCB2 = 0.00026
    FCM = 551
    VCM = 0.00008
elif TNP[1] < 1100000: # Processing facility = 150M
    FCB1 = 107
    VCB1 = 0.00019
    FCB2 = 158
    VCB2 = 0.00026
    FCM = 551
    VCM = 0.00008
    

RB1 = 1.3 # 0.95
RB2 = 2.3 # 0.995
RM  = 3   # 0.9995

''' **************************************************************************** Change these to select an instance '''

# Producer category  
PC = {} 
for p in Prange:  
    PC[p] = int (sheet.cell(3, p).value)  


if Psize <= 10:
    TC = 21
elif Psize <= 20:
    TC = 16
elif Psize <= 30: 
    TC = 11
print ('MF', FCM, VCM, end ="\t")
print ('BF1', FCB1, VCB1, end ="\t")
print ('BF2', FCB2, VCB2, end ="\t")
print ('TC', TC)      
print ("Drange =", Drange)             
print ("Days =", Dsize)       
print ("Weeks =", Wsize)       
print ("Wrange =", Wrange) 

# Number of packages in days
NDP = {}
#print ("Number of packages = ")
for d in Drange: 
    NDP[d] = {} 
    for p in Prange:
        NDP[d,p]= round( float (sheet.cell_value(10+p,d)) , 2)
        #print ("day: ", d, "Producer: ", p , NDP[d,p])
        
# Number of packages in weeks 
NWP = {}
for w in Wrange:    
    NWP[w] = {} 
    for p in Prange:
        temp = 0
        for d in range ( 7*(w-1)+1, 7*w+1 ) :
            if d <= Dsize:
                #print ('day: ', d)
                temp += NDP[d,p]
        NWP[w,p]= round( temp , 2)
#print ("Number of weekly packages = ", NWP)

ISC = {}
for d in Drange: 
    ISC[d] = {} 
    for p in Prange: 
        ISC[d,p] = round( float ( sheet.cell_value ( 12+Psize+p , d ) ) , 4)

Wsize = math.ceil (Dsize/7)          
#print ("Weeks =", Wsize) 
Wrange = range (1, Wsize + 1)          
#print ("Wrange =", Wrange) 
ISCW = {}
for w in Wrange: 
    ISC[w] = {}
    for p in Prange: 
        '''
        temp = round( float ( sheet.cell_value ( 14+2*Psize+p , w ) ) , 4) # It is equal to a random day of the week count
        if temp >= 0: 
            ISCW[w,p] = temp
        else:   
            ISCW[w,p] = 0
        '''
        ISCW[w,p] = round( float ( sheet.cell_value ( 14+2*Psize+p , w ) ) , 4) # It is equal to a random day of the week count
        #print ("W: ", w, "Producer: ", p ,  ISCW[w,p])        
        
AveISC = {}
for d in Drange:
    AveISC[d] = round( float ( sheet.cell_value ( 16+3*Psize , d ) ) , 4)   
#print ("AveISC = ", AveISC)
    
TNP0 = {}
for d in Drange:
    TNP0[d] = round( float ( sheet.cell_value ( 18+3*Psize, d ) ) )  
    AllMilkPackages += TNP0[d]



print ('*************************************************')     
print (' ')
    
# The input has wrong TNP) values so I recalculate them here.
TotalNPDaywith0spores = {}
for d in Drange:
    TotalSporesInADay = 0
    for p in Prange:  
        TotalSporesInADay += NDP[d,p] * (10 ** (ISC[d,p]) ) 
    if math.floor ( TotalSporesInADay ) < TNP[d] :
        TotalNPDaywith0spores[d] = round (TNP[d] - int (math.floor ( TotalSporesInADay ) ),2)   
    else: 
        TotalNPDaywith0spores[d] = 0
    if TotalNPDaywith0spores[d] != TNP0[d]:
        print ('Day', d, TotalNPDaywith0spores[d], 'was: ', TNP0[d])
        TNP0[d] = TotalNPDaywith0spores[d]

#print ('Day', TotalNPDaywith0spores[104], AveISCDaynow[104], TotalNPDay[104] )
        
print (' ')
print ('*************************************************')     
#print ("TNP0 = ", TNP0)

EachC = {}
EachC[1] = 0
EachC[2] = 0
EachC[3] = 0
EachC[4] = 0
EachC[5] = 0
ProducerIC = {}
for d in Drange: 
    ProducerIC[d] = {} 
    for p in Prange: 
        ProducerIC[d,p] = sheet.cell_value ( 20+3*Psize+p , d )  
        if (ProducerIC[d,p]==1):
            EachC[1] += ProducerIC[d,p] * NDP[d,p]
        if (ProducerIC[d,p]==2):
            EachC[2] += ProducerIC[d,p] * NDP[d,p]
        if (ProducerIC[d,p]==3):
            EachC[3] += ProducerIC[d,p] * NDP[d,p]
        if (ProducerIC[d,p]==4):
            EachC[4] += ProducerIC[d,p] * NDP[d,p]
        if (ProducerIC[d,p]==5):
            EachC[5] += ProducerIC[d,p] * NDP[d,p]
percent = {}
percent[1] =   EachC[1]/ (EachC[1]+EachC[2]+EachC[3]+EachC[4]+EachC[5])  
percent[2] =   EachC[2]/ (EachC[1]+EachC[2]+EachC[3]+EachC[4]+EachC[5])  
percent[3] =   EachC[3]/ (EachC[1]+EachC[2]+EachC[3]+EachC[4]+EachC[5])  
percent[4] =   EachC[4]/ (EachC[1]+EachC[2]+EachC[3]+EachC[4]+EachC[5])  
percent[5] =   EachC[5]/ (EachC[1]+EachC[2]+EachC[3]+EachC[4]+EachC[5])    

print (round (percent[1],2))    
print (round (percent[2],2))  
print (round (percent[3],2))    
print (round (percent[4],2))    
print (round (percent[5],2))  
  
# ************************************************************************* 
# ********************************* MILP ********************************** 
# *************************************************************************     

startBuild = time.time() 
    
# Create a model  

MILP = Model("FFAR Model") 
 	
#MILP.params.logtoconsole = ShowMILPStuff  # turns off display of Gurobi output when solving subproblems
#MILP.params.presolve = PresolverOn 

# ********************************** Create variables **********************************     

# addVar ( lb=0.0, ub=GRB.INFINITY, vtype=GRB.CONTINUOUS, name = 'TT[%d,%d]' %(i,t), column=None )

xMF   = MILP.addVar ( vtype = GRB.BINARY, name = 'xMF' ) 
xBF1  = MILP.addVar ( vtype = GRB.BINARY, name = 'xBF1' )
xBF2  = MILP.addVar ( vtype = GRB.BINARY, name = 'xBF2' )
xPR   = MILP.addVar ( vtype = GRB.BINARY, name = 'xPR' )
AveSL = MILP.addVar ( lb = 0, vtype = GRB.CONTINUOUS, name = 'AveSL' )
TR    = MILP.addVar ( lb = 0, vtype = GRB.CONTINUOUS, name = 'TR' )
PRhat = MILP.addVar ( lb = 0, ub = MaxPRhat, vtype = GRB.CONTINUOUS, name = 'PRhat' )
#b = MILP.addVar ( lb = 0, ub = B, vtype = GRB.CONTINUOUS, name = 'b' )
    
PR = {}
for c in Crange:
    PR[c] = MILP.addVar (lb=MinPR, ub=MaxPR, vtype = GRB.CONTINUOUS, name = 'PR[%d]' %(c)) 
    
PRpaid = {}
for w in Wrange:
    PRpaid[w] = {} 
    for p in Prange:
        PRpaid[w,p] = MILP.addVar (lb = MinPR, ub = MaxPR, vtype = GRB.CONTINUOUS, name = 'PRpaid[%d,%d]' %(w,p))
    
z = {}
for d in Drange: 
    z[d] = {}
    for cc in CSLrange:
        z[d,cc] = MILP.addVar (vtype = GRB.BINARY, name = 'z[%d,%d]' %(d,cc))
            
zhat = {}
for w in Wrange: 
    zhat[d] = {}
    for p in Prange:  
        zhat[w,p] = {}
        for c in Crange:
            zhat[w,p,c] = MILP.addVar (vtype = GRB.BINARY, name = 'zhat[%d,%d,%d]' %(w,p,c)) 

 
''' ************************** Constraints ************************** '''     

    
if (NoIntervention == 1 ):
    MILP.addConstr ( xMF == 0 ) 
    MILP.addConstr ( xBF1 == 0 ) 
    MILP.addConstr ( xBF2 == 0 )     
    MILP.addConstr ( xPR == 0 )   
    MILP.addConstr ( PRhat == 0 )  
    MILP.addConstr ( PR[1] == 0 )   
    MILP.addConstr ( PR[2] == 0 )    
    MILP.addConstr ( PR[3] == 0 )   
    MILP.addConstr ( PR[4] == 0 )    
    MILP.addConstr ( PR[5] == 0 )   
    #MILP.addConstr ( b == B )
    for w in Wrange: 
        for p in Prange:
            MILP.addConstr ( PRpaid[w,p] == 0 )
            
# ************************** Constraint BFs  
MILP.addConstr ( xBF1 + xBF2 <= 1 )    

# ************************** Constraint PROrder                   
for c in Crange:
    if c <= 4:
        MILP.addConstr ( PR[c] >= PR[c+1] ) 
        #MILP.addConstr ( PR[c] >= PR[c+1] + diff * xPR)         
        MILP.addConstr ( PR[c] >= 0 )          

# ************************** Constraint xPR                   
for c in Crange:
        MILP.addConstr ( xPR >= PR[c] )
MILP.addConstr ( xPR >= - PR[5] )    # since PR5 can be negative also        

# ************************** Constraint PRhat Calculation     
TotalPackagesNotCategory1 = 0
for d in Drange:
    for p in Prange:
        if PC[p] >= 2:
            TotalPackagesNotCategory1 += NDP[d,p]
#print ('TotalPackagesNotCategory1', TotalPackagesNotCategory1)

'''
if TotalPackagesNotCategory1 == 0:
    MILP.addConstr ( PRhat == 0 )
     MILP.addConstr ( xPR == 0 )
else:
    MILP.addConstr ( TotalPackagesNotCategory1 * PRhat == quicksum ( quicksum ( NDP[d,p] * ( alpha * PR[PC[p]-1] + beta * PR[PC[p]-2] ) for p in Prange if PC[p] >= 3 ) +
                                          quicksum ( NDP[d,p] * PR[PC[p]-1] for p in Prange if PC[p] == 2 ) for d in Drange )  )
'''

MILP.addConstr ( TotalPackagesNotCategory1 * PRhat == quicksum ( quicksum ( NDP[d,p] * ( alpha * PR[PC[p]-1] + beta * PR[PC[p]-2] ) for p in Prange if PC[p] >= 3 ) + quicksum ( NDP[d,p] * PR[PC[p]-1] for p in Prange if PC[p] == 2 ) for d in Drange )  )

''' Newly added '''
TotalProducersNotinCategory1 = 0
for p in Prange:
    if PC[p] >= 2:
        TotalProducersNotinCategory1 += 1
            
MILP.addConstr ( PRhat <= TotalProducersNotinCategory1 )

TotalProducersNotinCategory1and2 = 0
for p in Prange:
    if PC[p] >= 3:
        TotalProducersNotinCategory1and2 += 1
            
MILP.addConstr ( PR[5] >= - TotalProducersNotinCategory1and2 )
''' till here'''

# ************************** Constraint OneZhatSelected   
for w in Wrange:
    for p in Prange:
        MILP.addConstr ( quicksum ( zhat[w,p,c] for c in Crange ) == 1 )  

# ************************** Constraint MilkCategoryofEachProducerforEachday 1-2  
for w in Wrange: 
    for p in Prange: 
        # The first change can occur at the beginning of the planning horizon (s=1 => ceil)
        MILP.addConstr ( ISCW[w,p] - math.ceil(w/l[1]) * RPR[1] * a[1] * PRhat - math.floor(w/l[2]) * RPR[2] * a[2] * PRhat >= quicksum ( Cmin[c] * zhat[w,p,c] for c in Crange ) )
        MILP.addConstr ( ISCW[w,p] - math.ceil(w/l[1]) * RPR[1] * a[1] * PRhat - math.floor(w/l[2]) * RPR[2] * a[2] * PRhat <= quicksum ( Cmax[c] * zhat[w,p,c] for c in Crange ) )
                
# ************************** PRpaidCalculation 1-3              
for w in Wrange: 
    for p in Prange:
        for c in Crange:
            MILP.addConstr ( PRpaid[w,p] <= PR[c] + M * ( 1 - zhat[w,p,c] ) )
            MILP.addConstr ( PRpaid[w,p] >= PR[c] - M * ( 1 - zhat[w,p,c] ) )            
 
# ************************** Constraint Budget      
MILP.addConstr (  quicksum ( FCM * xMF + quicksum ( VCM * NDP[d,p] * xMF for p in Prange ) for d in Drange) + 
                  quicksum ( FCB1 * xBF1 + quicksum ( VCB1 * NDP[d,p] * xBF1 for p in Prange ) for d in Drange) + 
                  quicksum ( FCB2 * xBF2 + quicksum ( VCB2 * NDP[d,p] * xBF2 for p in Prange ) for d in Drange) + 
                  quicksum ( quicksum ( TC * xPR  for p in Prange ) for w in Wrange ) +
                  quicksum ( quicksum ( NWP[w,p] * PRpaid[w,p]  for p in Prange ) for w in Wrange ) <= Dsize * B ) 
          
# ************************** Constraint TRCalculation 
MILP.addConstr ( TR == RM * xMF + RB1 * xBF1 + RB2 * xBF2  ) 

# ************************** Constraint OneZSelected  
for d in Drange:
    MILP.addConstr ( quicksum ( z[d,cc] for cc in CSLrange ) == 1 )  

# ************************** Constraint MilkCofDay 1-2  
for d in Drange: 
    MILP.addConstr ( AveISC[d] - math.ceil( math.ceil(d/7)/l[1]) * RPR[1] * a[1] * PRhat - math.floor( math.ceil(d/7)/l[2]) * RPR[2] * a[2] * PRhat - TR >= quicksum ( CSLmin[cc] * z[d,cc] for cc in CSLrange ) )
    MILP.addConstr ( AveISC[d] - math.ceil( math.ceil(d/7)/l[1]) * RPR[1] * a[1] * PRhat - math.floor( math.ceil(d/7)/l[2]) * RPR[2] * a[2] * PRhat - TR <= quicksum ( CSLmax[cc] * z[d,cc] for cc in CSLrange ) )
'''
for d in Drange: 
    MILP.addConstr ( AveISC[d]  >= quicksum ( CSLmin[cc] * z[d,cc] for cc in CSLrange ) )
    MILP.addConstr ( AveISC[d]  <= quicksum ( CSLmax[cc] * z[d,cc] for cc in CSLrange ) )
'''
# ************************** Constraint AveSL       
MILP.addConstr ( AllMilkPackages * AveSL == quicksum ( quicksum ( (TNP[d]-TNP0[d]) * SL[cc] * z[d,cc]  for cc in CSLrange )  for d in Drange ) + quicksum ( TNP0[d] * SLmax for d in Drange ) )
''' 
print (TNP)
print (TNP0) 
print (AllMilkPackages)
'''
''' *********************** Objective function *********************** '''  
    
#MILP.setObjective ( AveSL + b / MM) 
MILP.setObjective ( AveSL ) 
MILP.modelSense = GRB.MAXIMIZE         
endBuild = time.time()
BuildTime = round(endBuild - startBuild,1) 
#print("Build time", round (endBuild - startBuild,1)) 

''' ************************************************************************* '''
''' ****************************** Solve MILP ******************************** '''
''' ************************************************************************* '''      

MILP.params.logtoconsole = ShowMILPStuff  # turns off display of Gurobi output when solving subproblems 
MILP.params.presolve = PresolverOn 
MILP.write("MILP.lp")
MILP.update()
MILP.setParam(GRB.Param.TimeLimit, MyTimeLimit) 
MILP.setParam(GRB.Param.MIPGap, MinGap) 

startSolve = time.time()

MILP.optimize() 


endSolve = time.time()

RunTime = round (MILP.Runtime,1)
            
status = MILP.Status 
     

if status == GRB.Status.OPTIMAL:
    
    '''
    print (' ')
    #print('BuildTime = ', BuildTime)
    #print ('-----')     
    print ('B       ', "\t", B )
    print ('b       ', "\t", round(b.x,2))
    print ('MBB   ', "\t",  int(xMF.x),  "\t", int(xBF1.x),  "\t",  int(xBF2.x) )  
    print ('Phat    ', "\t",  int (round (PRhat.x * 24 * 100, 0) ) )   
    print ('PRs     ', "\t",  end= ' ')  
    for c in Crange:
        print ( int (round ( PR[c].x * 24 * 100, 0) ), end=  "\t" )  
    print (' ')
    print ('OFV     ', "\t",  round(MILP.objVal,1))  
    print('Time     ', "\t",  int (RunTime) )
    print ('----------------------------------------------------------------------------')    
    print ('PR_hat  ', "\t",  round (PRhat.x, 5) , end= "\t")             
    for c in Crange:
        print (round ( PR[c].x, 5) , end= "\t") 
    print (' ')
    print ('b/MM', "\t", round(b.x / MM,4) ) 
    print ('OFV    ', "\t", round(MILP.objVal,3))  
    print ('MinGap ', "\t", round(MinGap,4))  
    print ('----------------------------------------------------------------------------')  
    '''
    print (' ')     
    print (B )
    #print (round(b.x,2))
    print (int(xMF.x),  ' ', int(xBF1.x),  ' ',  int(xBF2.x) )  
    print (int (round (PRhat.x * 24 * 100, 0) ) , ' => ', end = '  ')   
    for c in Crange: 
        print ( int (round ( PR[c].x * 24 * 100, 0) ), end=  '   ' )  
    print (' ')
    print (round(MILP.objVal,4))  
    print ('------')
    print (int (RunTime) )    
    print ('------')
    print (round (PRhat.x, 5) , end= '     ')             
    for c in Crange:
        print (round ( PR[c].x, 4) , end = '     ' ) 
    print (' ')
    print ("Producers =", Psize) 
    #print ('----------------------------------------------------------------------------')  
    #print (' ')
    
    
else:
    print ("Optimization stopped with status", status)     
