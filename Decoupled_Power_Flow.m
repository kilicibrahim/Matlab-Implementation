clc
clear
% returns numeric and text data in num and txt 
%, and string content in cell array raw
[num] = xlsread('ee303FinalProjectTemplate.xlsx','BUS');

%Identify Load P (real) and Q (imag) Values in per-unit
Sbase=100;
pd2=num(1,2)/Sbase;%real power of bus2
qd2=num(2,2)/Sbase;%reactive power of bus2
pd3=num(1,3)/Sbase;%real power of bus3
qd3=num(2,3)/Sbase;%reactive power of bus3
pg2=0;
qg2=0;
pg3=0;
qg3=0;
%calculate for each bus Real and Reactive Power
P2=pd2-pg2;%bus2 PQ bus, there is no generator real power
Q2=qd2-qg2;%and generator reactive power
P3=pd3-pg3;%bus3 PQ bus, there is no generator real power
Q3=qd3-qg3;%and generator reactive power


v1=num(3,1)*cosd(num(4,1))+num(3,1)*sind(num(4,1))*1i;
%initial guess for bus2 and bus3
v2=1+0i;
v3=1+0i;

%reading excel sheet BRANCH
[num,txt,raw] = xlsread('ee303FinalProjectTemplate.xlsx','BRANCH');

%---------- impedance of transmission lines-------%
%because in the raw we have string ,
%we convert the data into the double
z12=str2double(raw(2,3));
z13=str2double(raw(2,4));
z23=str2double(raw(4,3));
z21=z12;
z31=z13;
z32=z23;


%-------converting impedances to Admitances-----%
y12=1/z12;
y13=1/z13;
y23=1/z23;

%----------the Bus Admittance Matrix------%
Y= [(y12+y13) -(y12) -(y13); 
      -(y12) (y12+y23) -(y23); 
      -(y13) -(y23) (y23+y13)];
  
  
%----------Conductance Values------------%
G(1,1)=real(Y(1,1)); 
G(1,2)=real(Y(1,2)); 
G(1,3)=real(Y(1,3)); 
G(2,1)=real(Y(2,1)); 
G(2,2)=real(Y(2,2)); 
G(2,3)=real(Y(2,3)); 
G(3,1)=real(Y(3,1)); 
G(3,2)=real(Y(3,2)); 
G(3,3)=real(Y(3,3));
%--------Susceptance Values----------%
B(1,1)=imag(Y(1,1)); 
B(1,2)=imag(Y(1,2)); 
B(1,3)=imag(Y(1,3));
B(2,1)=imag(Y(2,1));
B(2,2)=imag(Y(2,2));
B(2,3)=imag(Y(2,3));
B(3,1)=imag(Y(3,1));
B(3,2)=imag(Y(3,2)); 
B(3,3)=imag(Y(3,3));

%--------- Given Specifications in pu (Known)----------%
V1MAG=abs(v1);   
ANG1=angle(v1); 
%--------unknowns initial guess-----------%
V2MAG=abs(v2);
ANG2=angle(v2);
V3MAG=abs(v3);
ANG3=angle(v3);

%we have 4 unknowns we will name them as x1,x2,x3,x4 as follows
x1=ANG2;
x2=ANG3;
x3=V2MAG;
x4=V3MAG;
k=1;%iteration number
epsilon=0.01;
x = [x1;x2;x3;x4];
f(1,1)=pd2-pg2;
f(2,1)=pd3-pg3;
f(3,1)=qd2-qg2;
f(4,1)=qd3-qg3;
        
        
        
        
      
for i=1:6
        
    %----------Calculating Jacobian------------%
    %J(1,1)=delP2/delANG2
    J(1,1)=x3*(V1MAG*(B(2,1)*cos(x1-ANG1)-G(2,1)*sin(x1-ANG1))+x4*(B(2,3)*cos(x1-x2)-G(2,3)*sin(x1-x2)));
    %J(1,2)=delP2/delx2 = J(1,2)
    J(1,2)=x3*x4*(G(2,3)*sin(x1-x2)-B(2,3)*cos(x1-x2));
    %J(1,3)=delP2/delx3
    J(1,3) =V1MAG*(G(2,1)*cos(x1- ANG1)+B(2,1)*sin(x1-ANG1))+2*x3*(G(2,2)*cos(x1-x1)+B(2,2)*sin(x1-x1))+x4*(G(2,3)*cos(x1-x2)+B(2,3)*sin(x1-x2));
    %J(1,4)=delP2/delx4
    %J(1,4)=x3*(G(2,3)*cos(x1-x2)+B(2,3)*sin(x1-x2));
    %we take it 0 because its value is verry small, so we do not take it
    J(1,4)=0;
    %J(2,1)=delP3/delx1
    J(2,1)=x4*x3*(G(3,2)*sin(x2-x1)-B(3,2)*cos(x2-x1));
    %J(2,2)=delP3/delx2
    J(2,2)=x4*(V1MAG*(B(3,1)*cos(x2-ANG1)-G(3,1)*sin(x2-ANG1))+x3*(B(3,2)*cos(x2-x1)-G(3,2)*sin(x2-x1)));
    %J(2,3)=delP3/delx3
    %J(2,3)=x4*(G(3,2)*cos(x2-x1)+B(3,2)*sin(x2-x1));
    %we take it 0 because its value is verry small, so we do not take it
    J(2,3)=0;
    %J(2,4)=delP3/delx4
    J(2,4)=V1MAG*(G(3,1)*cos(x2-ANG1))+x3*(G(3,2)*cos(x2-x1))+2*x4*(G(3,3)*cos(x2-x2));
    %J(3,1)delQ2/delx1
    J(3,1)=x3*V1MAG*(B(2,1)*sin(x1-ANG1)+G(2,1)*cos(x1-ANG1))+x3*x4*(B(2,3)*sin(x1-x2)+G(2,3)*cos(x1-x2));
    %J(3,2)=delQ2/delx2
    %J(3,2)=x3*x4*(G(2,3)*cos(x1-x2)+B(2,3)*sin(x1-x2));
    %we take it 0 because its value is verry small, so we do not take it
    J(3,2)=0;
    %J(3,3)=delQ2/delx3
    J(3,3)=V1MAG*(G(2,1)*sin(x1-ANG1)-B(2,1)*cos(x1-ANG1))+2*x3*(G(2,2)*sin(x1-x1)-B(2,2)*cos(x1-x1))+x4*(G(2,3)*sin(x1-x2)-B(2,3)*cos(x1-x2));
    %J(3,4)=delQ2/delx4
    J(3,4)=x3*(G(2,3)*sin(x1-x2)-B(2,3)*cos(x1-x2));
    %J(4,1)=delQ3/delx1
    %J(4,1)=x4*V1MAG*(G(3,1)*cos(x2-ANG1)+B(3,1)*sin(x2-ANG1))+x4*x3*(B(3,2)*sin(x2-x1)+G(3,2)*cos(x2-x1));
    %we take it 0 because its value is verry small, so we do not take it
    J(4,1)=0;
    %J(4,2)=delQ3/delx2
    J(4,2)=x4*V1MAG*(B(3,1)*sin(x2-ANG1)+G(3,1)*cos(x2-ANG1))+x4*x3*(B(3,2)*sin(x2-x1)+G(3,2)*cos(x2-x1))+x4*x4*(G(3,3)*cos(x2-x2)+B(3,3)*sin(x2-x2));
    %J(4,3)=delQ3/delx3
    J(4,3)=x4*(G(3,2)*cos(x2-x1)+B(3,2)*sin(x2-x1));
    %J(4,4)=delQ3/delx4
    J(4,4)=V1MAG*(G(3,1)*cos(x2-ANG1)+B(3,1)*sin(x2-ANG1))+x3*(G(3,2)*cos(x2-x1)+B(3,2)*sin(x2-x1))+2*x4*(G(3,3)*cos(x2-x2)+B(3,3)*sin(x2-x2));
    
    
        
        %we calculated the variables of vectors one by one and then assign them to
        %f whcih is our f(x)
        %--------f(x)-------------%
        P2=x3*V1MAG*(G(2,1)*cos(x1-ANG1)+B(2,1)*sin(x1-ANG1))+(x3*x3*(G(2,2)*cos(x1-x1)+B(2,2)*sin(x1-x1)))+(x3*x4*(G(2,3)*cos(x1-x2)+B(2,3)*sin(x1-x2)));
        P3=x4*V1MAG*(G(3,1)*cos(x2-ANG1)+B(3,1)*sin(x2-ANG1))+(x4*x3*(G(3,2)*cos(x2-x1)+B(3,2)*sin(x2-x1)))+(x4*x4*(G(3,3)*cos(x2-x2)+B(3,3)*sin(x2-x2)));
        Q2=x3*V1MAG*(G(2,1)*sin(x1-ANG1)-B(2,1)*cos(x1-ANG1))+(x3*x3*(G(2,2)*sin(x1-x1)-B(2,2)*cos(x1-x1)))+(x3*x4*(G(2,3)*sin(x1-x2)-B(2,3)*cos(x1-x2)));
        Q3=x4*V1MAG*(G(3,1)*sin(x2-ANG1)-B(3,1)*cos(x2-ANG1))+(x4*x3*(G(3,2)*sin(x2-x1)-B(3,2)*cos(x2-x1)))+(x4*x4*(G(3,3)*sin(x2-x2)-B(3,3)*cos(x2-x2)));
        %Calculate delta P and Q
        delta_P(1,1)=P2;
        delta_P(2,1)=P3;
        delta_Q(1,1)=Q2;
        delta_Q(2,1)=Q3;
        
        f(1,1)=P2+f(1,1);
        f(2,1)=P3+f(2,1);
        f(3,1)=Q2+f(3,1);
        f(4,1)=Q3+f(4,1);
        
        
        f=-J*x;
       
        %Calculate the next point of x
            
            k=k+1; 
end
 x1=f(1,1);
            x2=f(2,1);
            x3=f(3,1);
            x4=f(4,1);
