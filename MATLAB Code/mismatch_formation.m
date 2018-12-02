function [] = mismatch_formation(V_Ref, Admittance_Complete, P_Active)
Y = Admittance_Complete;
V = ones(size(Admittance_Complete,1),1);
theta = zeros(1,12);

for i = (1:size(V_Ref,1))
    V(V_Ref(i,1)) = V_Ref(i,2);
end

P = zeros(size(Admittance_Complete,1),1);
for i = (1:size(P_Active,1))
   P(P_Active(i,1)) = P_Active(i,2); 
end

P_temp = [];
Q_temp = [];
for k = 1:size(Y,1)
    for i = 1:size(Y,1)
       %P_temp(k) = V(k)*V(i)*(Y(k,i).real*cos(theta(k) - theta(i)) + Y(k,i).imag*sin(theta(k) - theta(i)));
       %Q_temp(k) = V(k)*V(i)*(Y(k,i).real*sin(theta(k) - theta(i)) - Y(k,i).imag*cos(theta(k) - theta(i)));
    end
end

