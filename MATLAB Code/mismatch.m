function [MisM] = mismatch(V, Theta, Y, P_Active, Bus_Loads)

%Active power vector
P = zeros(size(Y,1),1);

%P_temp will hold the mismatch values before being added to the active and
%load power, I'm doing a 12x12 so that I can see each value of the sum for
%Pk and Qk, to make debugging easier.

P_temp = zeros(12,12);
Q_temp = zeros(12,12);

%Column vectors to be the final form, and added to active and load power
P_col = [];
Q_col = [];

MisM = zeros(1,(size(P_col,2) + (size(Q_col,2))));

for i = (1:size(P_Active,1))
    P(P_Active(i,1)) = P_Active(i,2);
end

for k = 1:(size(Y,1))
    for i = 1:(size(Y,1))
        P_temp(k,i) = V(k)*V(i)*(real(Y(k,i))*cos(Theta(k) - Theta(i)) + imag(Y(k,i))*sin(Theta(k) - Theta(i)));
       if ~(ismember(i,P_Active(:,1)))
          Q_temp(k,i) = V(k)*V(i)*(real(Y(k,i))*sin(Theta(k) - Theta(i)) - imag(Y(k,i))*cos(Theta(k) - Theta(i)));
       end
    end
end

% Representing the summation in the original formula
P_col = sum(P_temp);
Q_col = sum(Q_temp);

%Populating the returned mismatch equation
for i = (1:size(P_col,2))
    MisM(i) = P_col(i);
    for i = (1:size(Q_col,2))
        MisM(size(Bus_Loads,2) + 1 + i) = Q_col(i);
    end
end

