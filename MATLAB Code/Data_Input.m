%{
    Newton Raphson Powerflow 
    EE 454 Prof. Christie
    Josh Reidt, Devin Pegues, Jacob Madian

    This program reads in data from excel sheets, then performs
    the Newton Raphson Method to see if the model converges.

%}

clear all;
clc

format short

%Reading in the data from the excel sheets
P_Active = xlsread('Active_Power_Production.xlsx');
Line_Data = xlsread('Line_Data.xlsx');
V_Ref = xlsread('PV_Bus_Reference_Voltages.xlsx');
Bus_Loads = xlsread('Bus_Loads.xlsx');
Admittance = zeros(size(Bus_Loads,2)+ 1,size(Bus_Loads,2) + 1);


%Using this for the final code, leaving commented until final
%implementation

%prompt = 'cold start?'
%response = input(prompt,'s')
response = 'yes';

if response == 'yes'
    %V will be the voltage p.u. of the given and initial guesses
    V = ones(size(Admittance,1),1);
    %Theta will be all zeros initially
    Theta = zeros(1,12);
    
    for i = (1:size(V_Ref,1))
        V(V_Ref(i,1)) = V_Ref(i,2);
    end
end

%Forming the initial Y matrix before incorporating diagonal and effect of B
R = Line_Data(:,3);
X = (1j*Line_Data(:,4));
Z = plus(R,X);
Y = -power(Z,-1);

% Since B effects only certain lines, we use a matrix of zeros and modify
% it
B = zeros(1,size(Bus_Loads,2) + 1);
B_in = Line_Data(:,5);

%This for loop populates the Y matrix, and forms the B matrix to add to Y
for i = (1:size(Line_Data,1))
    row_coor = Line_Data(i,1);
    col_coor = Line_Data(i,2);
    Admittance(row_coor, col_coor) = Y(i);
    Admittance(col_coor, row_coor) = Y(i);
    
    for k = (1:size(Line_Data,1))
        if (Line_Data(k,1) == i)
          B(i) = B(i) + B_in(k)/2;
        end
    end
end

% Formation of B and Y diagonal matrices
B_diag = diag((-1j*B)');
Y_diag = diag(-1 * sum(Admittance));

% This is the final Admittance matrix
Y = Admittance + B_diag + Y_diag;

mismatch = mismatch(V,Theta,Y,P_Active,Bus_Loads);
