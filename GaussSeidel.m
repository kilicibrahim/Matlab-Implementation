clc
clear

% returns numeric and text data in num and txt 
%, and unprocessed cell content in cell array raw
[num] = xlsread('ee303FinalProjectTemplate.xlsx','BUS');
%Identify Load P (real) and Q (imag) Values in per-unit
Sbase=100;
%retrieve the data from  1.row 2nd column for real power P2
%retrieve the data from 2nd row 22nd column for imag of S2
%retrieve the data from  1.row 3rd column for real power P3
%retrieve the data from 2nd row 3rd column for imag of S3
p2=num(1,2)/Sbase;
q2=num(2,2)/Sbase;

p3=num(1,3)/Sbase;
q3=num(2,3)/Sbase;
v1=num(3,1)*cosd(num(4,1))+num(3,1)*sind(num(4,1))*1i;
[num,txt,raw] = xlsread('ee303FinalProjectTemplate.xlsx','BRANCH');


%v1=1+0i;

%initial guess
v2=1+0i;
v3=1+0i;


% impedance of transmission lines
%because in the raw we have string ,
%we convert the data into the double
z12=str2double(raw(2,3));
z13=str2double(raw(2,4));
z23=str2double(raw(4,3));
z21=z12;
z31=z13;
z32=z23;


%converting impedances to Admitances
y12=1/z12;
y13=1/z13;
y23=1/z23;

%the Bus Admittance Matrix
Y= [(y12+y13) -(y12) -(y13); 
      -(y12) (y12+y23) -(y23); 
      -(y13) -(y23) (y23+y13)];
  
 n=1;
 eps=0.01;
 %Vi=1/Yii[(Si*/Vi*)-Sum(Yik*Vk)]
 %e2 and e3 are the diference between the new value of v2 and v3
 %and and old value of v2 and v3 
 %em variable is assigned as the  difference or error value
ite=1;

 while n > eps                                                        
    
    v2e = ((conj(p2-(q2*1i))/conj(v2))-y12*v1+y23*v3)/(y12+y23);                 
    v3e = ((conj(p3-(q3*1i))/conj(v3))-y13*v1+y23*v2)/(y13+y23);                 
    e2 = abs(v2e - v2);                                                     
    e3 = abs(v3e - v3);
    v2 = v2e;                                                               
    v3 = v3e;                                                               
    em = [e2,e3];                                                           
    n = max(em);
    v2array(ite)=abs(v2);
    v3array(ite)=abs(v3);
                
    ite=ite+1;
 end
v2array(ite)=abs(v2);
v3array(ite)=abs(v3);
plot(v2array,v3array)
disp(ite) 
disp('V2 Value =')
ang2=180/pi*angle(v2);
ang3=180/pi*angle(v3);
disp(v2)
disp('V3 Value =')
disp(v3)
abs(v2)

a = num2str(v2);
% b=num2str(v3)
sprintf(a)


