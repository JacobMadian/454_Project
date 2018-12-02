clear all;
clc

format long
P_Active = xlsread('Active_Power_Production.xlsx');
Line_Data = xlsread('Line_Data.xlsx');
V_Ref = xlsread('PV_Bus_Reference_Voltages.xlsx');
Bus_Loads = xlsread('Bus_Loads.xlsx');
[m,n] = size(Line_Data);
Admittance = zeros(12,12);

%prompt = 'cold start?'
%response = input(prompt,'s')
response = 'yes';

if response == 'yes'
    V = ones(size(Admittance,1),1);
    Theta = zeros(1,12);
    
    for i = (1:size(V_Ref,1))
        V(V_Ref(i,1)) = V_Ref(i,2);
    end
end

B = zeros(1,12);
R = Line_Data(:,3);
X = (1j*Line_Data(:,4));
Z = plus(R,X);
Y = -power(Z,-1);
B_in = Line_Data(:,5);

for i = (1:m)
    row_coor = Line_Data(i,1);
    col_coor = Line_Data(i,2);
    Admittance(row_coor, col_coor) = Y(i);
    Admittance(col_coor, row_coor) = Y(i);
    for k = (1:m)
        if (Line_Data(k,1) == i)
          B(i) = B(i) + B_in(k)/2;
        end
    end
end
B_diag = diag((-1j*B)');
Y_diag = diag(sum(Admittance));
Admittance_Complete = Admittance + B_diag + Y_diag;

%%
Y = Admittance_Complete;


P = zeros(size(Y,1),1);
for i = (1:size(P_Active,1))
   P(P_Active(i,1)) = P_Active(i,2); 
end

P_temp = zeros(12,12);
Q_temp = zeros(12,12);
P_curr = [];
Q_curr = [];

for k = 1:size(Y,1)
    for i = 1:size(Y,1)
       P_temp(k,i) = V(k)*V(i)*(real(Y(k,i))*cos(Theta(k) - Theta(i)) + imag(Y(k,i))*sin(Theta(k) - Theta(i)));
       if ~(ismember(i,P_Active(:,1)))
          Q_temp(k,i) = V(k)*V(i)*(real(Y(k,i))*sin(Theta(k) - Theta(i)) - imag(Y(k,i))*cos(Theta(k) - Theta(i)));
       end
    end
end

P_curr = sum(P_temp);
Q_curr = sum(Q_temp);
MisM = zeros(1,(size(P_curr,2) + (size(Q_curr,2))));

for i = (1:size(P_curr,2))
    MisM(i) = P_curr(i);
end
for i = (1:size(Q_curr,2))
   MisM(size(P_curr,2) + i) = (Q_curr(i)); 
end

