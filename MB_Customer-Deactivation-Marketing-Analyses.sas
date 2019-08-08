Proc import datafile='C:\Users\bmehdi\Desktop\Advanced SAS project\New_Wireless_Pipe.txt'
            Out=df
            DBMS=DLM
            REPLACE;

    DELIMITER="|";
    GETNAMES=YES;
    GUESSINGROWS=10;

Run;

Data df2;
    retain  RatePlan
            Acctno1
			Actdt
			Deactdt
			DeactReason
			GoodCredit			
			DealerType
			AGE
			Province
			Sales;

	Set df;
	acctno1 = PUT(acctno, BEST13.);
	drop acctno;
Run;

Data Df5;
	Length Region $17 Acctstatus $11 AgeGroup $16 SalesGroup $16;
	Format Actdt Deactdt DATE9.;
	Retain 	Province
			AcctStatus;
 	Set Df2;
	If Province in ('BC' 'SK' 'AB') Then Region      = '  West Provinces';
    If Province in ('PE' 'NS' 'NB' 'NL') Then Region = ' Ocean Provinces';
    If Province in ('MT','QC','ON') Then Region      = 'Central Provinces';

    If Deactdt = . Then 
        AcctStatus = 'Active';
    Else
        AcctStatus = 'Deactivated';

    If AGE = . Then AgeGroup = "Age: Unknown";
    Else if AGE <=21 Then AgeGroup = ' <= 20 yrs';
    Else if 21 < AGE <= 41 Then AgeGroup = '21 - 40 yrs';
	Else if 41 < AGE <= 61 Then AgeGroup = '41 - 60 yrs';
    Else if AGE > 61 Then AgeGroup = '> 60 yrs';

	If Sales = . Then SalesGroup = "Sale: Unknown";
    Else if Sales <=100 Then SalesGroup = ' <= $100';
    Else if 100 < Sales <= 500 Then SalesGroup = '$101 - $500';
	Else if 500 < Sales <= 800 Then SalesGroup = '$501 - $800';
    Else if Sales > 800 Then SalesGroup = '> $800';
	
	If AcctStatus = 'Deactivated' Then 
        Do;
            Tenure = Deactdt - Actdt;
            Deact_month = Month(deactdt);
            Deact_year = Year(deactdt);
        End;
    Else
        Do;
            Tenure = '31Mar2001'd - Actdt; 
            Active_month = Month(Actdt);
            Active_year = Year(Actdt);
        End;    
Run;


/*Is the acctno unique? or is there any duplicates in acctno? */
Proc Sql;
title 'How many duplicate are there for Account number (Acctno1)?';
	Select Acctno1, count(*) AS counter 
	From Work.Df5
	Group by Acctno1
	Having counter > 1 
	;
Quit;

Proc SQL;
title 'How many duplicate are there for Account number?';
    Select count(acctno1) as Count,
           count(distinct acctno1) as 'Distinct Count'n
    From Df5
    ;
Quit;

/*
======= Exploratory Data Analysis ======
======= Exploratory Data Analysis ======
======= Exploratory Data Analysis ======
*/
Proc Means data=Df5 N MEAN MEDIAN MODE MIN MAX CLM ALPHA= .05 MAXDEC=2;
	Var Sales Age Tenure;
    Title "Summary of Descriptive Statistics";
Run;

Proc Sql;
title 'Earliest and Lastest Activation and Deactivation';
	Select min(Actdt) As 'Earliest Activation Date'n format=yymmdd10., max(Actdt) As 'Lastest Activation Date'n format=yymmdd10. , min(Deactdt) As 'Earliest Deactivation Date'n format=yymmdd10., max(Deactdt) As 'Lastest Deactivation Date'n format=yymmdd10.
	From Work.Df5
;    
Quit;

/*What are the counts by Provinces, Regions and Age Group ? */
Proc Tabulate data=Df5;
title 'Frequency of Each Province';
    Class Province;    
    Table Province;
Run;

Proc sgplot data=Df5;
    Vbar Province;
    Label ;
    Title2 'Frequency by Province';
Run;

Proc Tabulate data=Df5;
    Class Region;    
    Table Region;
Run;

Proc Gchart data=Df5;
    Pie Region/Discrete value=inside
             Percent=Outside slice=outside;
	Title 'Frequency by Region';
Run;
Title;

Proc Tabulate data=Df5;
    Class AgeGroup;    
    Table Agegroup;
Run;

Proc sgplot data=Df5 noborder;
    Vbar AgeGroup;
    Title2 'Frequency by AgeGroup';
Run;

/*Univariate*/
Proc Univariate Data=Df5;
    Var Sales;
    Histogram Sales / normal;
    Probplot Sales;
    Title;
Run;

proc univariate data=Df5;
  class Region;
  var Sales;
  histogram Sales / kernel overlay; 
run;

proc sgplot data=Df5;
  where Acctstatus;
  histogram Sales / group=Acctstatus transparency=0.5;       
  density Sales / type=kernel group=Acctstatus; 
run;

proc univariate data=Df5;
  class GoodCredit;
  var Sales;
  histogram Sales / kernel overlay; 
run;

Proc Univariate Data=Df5;
    Var Tenure;
    Histogram Tenure / normal;
    Title;
Run;

proc univariate data=Df5;
  class Region;
  var Tenure;
  histogram Tenure / kernel overlay; 
run;

Proc Sgplot data=Df5;
    Vbox Tenure/Category=Region;
    Title2 "Tenure by Region";
Run;

Proc Sgplot data=Df5;
    Vbox Sales/Category=Region;
    Title2 "Sales by Region";
Run;

Proc Sgplot data=Df5;
    Vbox Tenure/Category=Province;
    Title2 "Tenure by Province";
Run;

Proc Sgplot data=Df5;
    Vbox Sales/Category=Province;
    Title2 "Sales by Province";
Run;
title;

proc chart data=Df5; 
hbar Province / discrete type=Sum           
sumvar=Sales mean; 
run; 
quit; 

Proc Chart data=Df5; 
hbar Province / discrete type=mean sumvar=Sales mean 
group=AcctStatus;
format Sales DOLLAR8.2;
run; 
quit; 

Proc Chart data=Df5; 
hbar Region / discrete type=sum sumvar=Sales mean 
subgroup=AcctStatus;
run; 
quit; 

Proc Sgplot data=Df5;
    Vbox Tenure/Category=AgeGroup;
    Title2 "Tenure by Age Group";
Run;

Proc Sgplot data=Df5;
    Vbox Sales/Category=AgeGroup;
    Title2 "Sales by Age Group";
Run;

Proc Sgplot data=Df5;
    Vbox Tenure/Category=AcctStatus;
    Title2 "Tenure by Account Status";
Run;

Proc Sgplot data=Df5;
    Vbox Sales/Category=AcctStatus;
    Title2 "Sales by Account Status";
Run;

Proc Sgplot data=Df5;
    Vbox Tenure/Category=GoodCredit;
    Title2 "Tenure by GoodCredit";
Run;

Proc Sgplot data=Df5;
    Vbox Sales/Category=GoodCredit;
    Title2 "Sales by GoodCredit";
Run;

/* Determine Total Activated Account and Total Deactivated Account by Province */
Proc Tabulate data=Df5;
    Class Province AcctStatus;    
    Table Province*AcctStatus='Account Status'*f=COMMA15.;
Run;

Proc Freq Data=Df5;
    Tables Region * Acctstatus / CHISQ NOROW NOCOL NOPERCENT;
    Title;
Run; 

Proc Freq Data=Df5;
    Tables AgeGroup * Acctstatus / CHISQ NOROW NOCOL NOPERCENT;
    Title;
Run; 

Proc Freq Data=Df5;
    Tables SalesGroup * Acctstatus / CHISQ NOROW NOCOL NOPERCENT;
    Title;
Run; 

Proc Freq Data=Df5;
    Tables Acctstatus * GoodCredit / CHISQ NOROW NOCOL NOPERCENT;
    Title;
Run; 

Proc Ttest Data=Df5;
	Class Acctstatus;
	Var Sales;
	Title "Test difference between two levels of Account Status";
Run;

Proc Sgplot data=Df5;
    Scatter x =  Age y =  Tenure / Group=Region;
    Title "Comparison of Age vs Tenure";
Run;

Proc Sgplot data=Df5;
    Scatter x =  Sales y =  Tenure / Group=Region;
    Title "Comparison of Sales vs Tenure";
Run;

Proc Sgplot data=Df5;
    Scatter x =  Age y =  Sales / Group=Region;
    Title "Comparison of Age vs Sales";
Run;

*=====Correlation;
Proc Corr data=Df5;
    Var Sales Tenure;
    With Age;
Run;

Proc Corr data=Df5;
    Var Tenure;
    With Sales;
Run;


/*
======= Churn RATE ======
======= Churn RATE ======
======= Churn RATE ======
*/

Proc SQL;
    Create Table Monthly_Deact_Acct as
    Select Deact_year,
           Deact_month, 
           Case Deact_month
               When 1 Then 'Jan'
               When 2 Then 'Feb'
               When 3 Then 'Mar'
               When 4 Then 'Apr'
               When 5 Then 'May'
               When 6 Then 'Jun'
               When 7 Then 'Jul'
               When 8 Then 'Aug'
               When 9 Then 'Sep'
               When 10 Then 'Oct'
               When 11 Then 'Nov'
               When 12 Then 'Dec'
               Else 'UNK'
           End as Month_Abbr,
           Count(AcctNo1) as Total_DeActivated           
    From Work.Df5
    Where AcctStatus = 'Deactivated'
    Group by Deact_year, Deact_month
    ;
    Create Table Monthly_Active_Acct as
    Select Active_year,
           Active_month, 
           Case Active_month
               When 1 Then 'Jan'
               When 2 Then 'Feb'
               When 3 Then 'Mar'
               When 4 Then 'Apr'
               When 5 Then 'May'
               When 6 Then 'Jun'
               When 7 Then 'Jul'
               When 8 Then 'Aug'
               When 9 Then 'Sep'
               When 10 Then 'Oct'
               When 11 Then 'Nov'
               When 12 Then 'Dec'
               Else 'UNK'
           End as Month_Abbr,
           Count(AcctNo1) as Total_Active           
    From Work.Df5
    Where AcctStatus = 'Active'
    Group by Active_year, Active_month
    ;
    Create Table Monthly_Churn as
    Select t1.deact_year,
           t1.deact_month,
           Case t1.deact_Month
               When 1 Then 'Jan'
               When 2 Then 'Feb'
               When 3 Then 'Mar'
               When 4 Then 'Apr'
               When 5 Then 'May'
               When 6 Then 'Jun'
               When 7 Then 'Jul'
               When 8 Then 'Aug'
               When 9 Then 'Sep'
               When 10 Then 'Oct'
               When 11 Then 'Nov'
               When 12 Then 'Dec'
               Else 'UNK'
           End as Month_name,
           t1.Total_DeActivated,
           t2.Total_active,
           (t1.Total_Deactivated / (t1.Total_Deactivated + t2.Total_active)) as churn_rate format=PERCENT9.2
    From Monthly_Deact_Acct t1 inner join Monthly_Active_Acct t2
    on t1.Deact_Year = t2.active_year and
       t1.Deact_month = t2.active_month
    order by t1.deact_year, t1.deact_month
    ;
    
Quit;


Data Monthly_Churn (keep=Month_no Total_Active Total_Deactivated Churn_rate);
    Retain Month_no Churn_rate Total_Active Total_Deactivated;
    Set Monthly_Churn;
    Month_no = _N_;
Run;

Proc Corr data=Monthly_Churn;
    Var Churn_rate Total_Active Total_Deactivated;
    With Month_no;
Run;

Proc Corr data=Monthly_Churn;
    Var Total_Active;
    With  Total_Deactivated;
Run;

symbol1 color=blue interpol=join; 
proc gplot data=Monthly_Churn;  
plot churn_rate*month_no / hminor=0;  
run; 
quit;


symbol1 color=blue interpol=join; symbol2 color=red interpol=join;
proc gplot data=Monthly_Churn; 
plot churn_rate*month_no / caxis=blue hminor=0          noframe;  
plot2 Total_Deactivated*month_no / caxis=red;  
run; 
quit;

proc sgplot data=Monthly_Churn;
	reg x=month_no y=churn_rate / transparency=0.3 
                                  markerattrs=(symbol=CircleFilled color=black size=12px) 
                                  lineattrs=(color=green thickness=2);
  
run;

title1  ’Linear Regression of Total_Deactivated and month_no; title2 ’(with 95% Confidence Limits)’;
symbol ci=red cv=blue co=gray value=dot       interpol=rlclm95 ;
proc gplot data=Monthly_Churn;   
plot Total_Deactivated*month_no / regeqn; 
run; 
quit;


/*
======= CREATE GOOD/BAD CREDIT TABLE ======
======= CREATE GOOD/BAD CREDIT TABLE ======
======= CREATE GOOD/BAD CREDIT TABLE ======
*/

Proc SQL;
Create Table Df5_Good as
Select * from Df5
Where GoodCredit = 1
;
Quit;

Proc SQL;
Create Table Df5_Bad as
Select * from Df5
Where GoodCredit = 0
;
Quit;

/*
=======CHURN RATE FOR GOOD CREDIT ====
=======CHURN RATE FOR GOOD CREDIT ====
=======CHURN RATE FOR GOOD CREDIT ====
*/

Proc SQL;
    Create Table Monthly_Deact_Acct_Good as
    Select Deact_year,
           Deact_month, 
           Case Deact_month
               When 1 Then 'Jan'
               When 2 Then 'Feb'
               When 3 Then 'Mar'
               When 4 Then 'Apr'
               When 5 Then 'May'
               When 6 Then 'Jun'
               When 7 Then 'Jul'
               When 8 Then 'Aug'
               When 9 Then 'Sep'
               When 10 Then 'Oct'
               When 11 Then 'Nov'
               When 12 Then 'Dec'
               Else 'UNK'
           End as Month_Abbr,
           Count(AcctNo1) as Total_DeActivated           
    From Work.Df5_Good
    Where AcctStatus = 'Deactivated'
    Group by Deact_year, Deact_month
    ;
    Create Table Monthly_Active_Acct_Good as
    Select Active_year,
           Active_month, 
           Case Active_month
               When 1 Then 'Jan'
               When 2 Then 'Feb'
               When 3 Then 'Mar'
               When 4 Then 'Apr'
               When 5 Then 'May'
               When 6 Then 'Jun'
               When 7 Then 'Jul'
               When 8 Then 'Aug'
               When 9 Then 'Sep'
               When 10 Then 'Oct'
               When 11 Then 'Nov'
               When 12 Then 'Dec'
               Else 'UNK'
           End as Month_Abbr,
           Count(AcctNo1) as Total_Active           
    From Work.Df5_Good
    Where AcctStatus = 'Active'
    Group by Active_year, Active_month
    ;
    Create Table Monthly_Churn_Good as
    Select t1.deact_year,
           t1.deact_month,
           Case t1.deact_Month
               When 1 Then 'Jan'
               When 2 Then 'Feb'
               When 3 Then 'Mar'
               When 4 Then 'Apr'
               When 5 Then 'May'
               When 6 Then 'Jun'
               When 7 Then 'Jul'
               When 8 Then 'Aug'
               When 9 Then 'Sep'
               When 10 Then 'Oct'
               When 11 Then 'Nov'
               When 12 Then 'Dec'
               Else 'UNK'
           End as Month_name,
           t1.Total_DeActivated,
           t2.Total_active,
           (t1.Total_Deactivated / (t1.Total_Deactivated + t2.Total_active)) as churn_rate format=PERCENT9.2
    From Monthly_Deact_Acct_Good t1 inner join Monthly_Active_Acct_Good t2
    on t1.Deact_Year = t2.active_year and
       t1.Deact_month = t2.active_month
    order by t1.deact_year, t1.deact_month
    ;
    
Quit;

Data Monthly_Churn_Good (keep=Month_no Total_Active Total_Deactivated Churn_rate);
    Retain Month_no Churn_rate Total_Active Total_Deactivated;
    Set Monthly_Churn_Good;
    Month_no = _N_;
Run;

/*
=======CHURN RATE FOR BAD CREDIT ====
=======CHURN RATE FOR BAD CREDIT ====
=======CHURN RATE FOR BAD CREDIT ====
*/

Proc SQL;
    Create Table Monthly_Deact_Acct_Bad as
    Select Deact_year,
           Deact_month, 
           Case Deact_month
               When 1 Then 'Jan'
               When 2 Then 'Feb'
               When 3 Then 'Mar'
               When 4 Then 'Apr'
               When 5 Then 'May'
               When 6 Then 'Jun'
               When 7 Then 'Jul'
               When 8 Then 'Aug'
               When 9 Then 'Sep'
               When 10 Then 'Oct'
               When 11 Then 'Nov'
               When 12 Then 'Dec'
               Else 'UNK'
           End as Month_Abbr,
           Count(AcctNo1) as Total_DeActivated           
    From Work.Df5_Bad
    Where AcctStatus = 'Deactivated'
    Group by Deact_year, Deact_month
    ;
    Create Table Monthly_Active_Acct_Bad as
    Select Active_year,
           Active_month, 
           Case Active_month
               When 1 Then 'Jan'
               When 2 Then 'Feb'
               When 3 Then 'Mar'
               When 4 Then 'Apr'
               When 5 Then 'May'
               When 6 Then 'Jun'
               When 7 Then 'Jul'
               When 8 Then 'Aug'
               When 9 Then 'Sep'
               When 10 Then 'Oct'
               When 11 Then 'Nov'
               When 12 Then 'Dec'
               Else 'UNK'
           End as Month_Abbr,
           Count(AcctNo1) as Total_Active           
    From Work.Df5_Bad
    Where AcctStatus = 'Active'
    Group by Active_year, Active_month
    ;
    Create Table Monthly_Churn_Bad as
    Select t1.deact_year,
           t1.deact_month,
           Case t1.deact_Month
               When 1 Then 'Jan'
               When 2 Then 'Feb'
               When 3 Then 'Mar'
               When 4 Then 'Apr'
               When 5 Then 'May'
               When 6 Then 'Jun'
               When 7 Then 'Jul'
               When 8 Then 'Aug'
               When 9 Then 'Sep'
               When 10 Then 'Oct'
               When 11 Then 'Nov'
               When 12 Then 'Dec'
               Else 'UNK'
           End as Month_name,
           t1.Total_DeActivated,
           t2.Total_active,
           (t1.Total_Deactivated / (t1.Total_Deactivated + t2.Total_active)) as churn_rate format=PERCENT9.2
    From Monthly_Deact_Acct_Bad t1 inner join Monthly_Active_Acct_Bad t2
    on t1.Deact_Year = t2.active_year and
       t1.Deact_month = t2.active_month
    order by t1.deact_year, t1.deact_month
    ;
    
Quit;

Data Monthly_Churn_Bad (keep=Month_no Total_Active Total_Deactivated Churn_rate);
    Retain Month_no Churn_rate Total_Active Total_Deactivated;
    Set Monthly_Churn_Bad;
    Month_no = _N_;
Run;


*=======CORRELATION BETWEEN MONTH NO AND CHURN RATE FOR GOOD CREDIT AND BAD CREDIT ====;
Proc Corr data=Monthly_Churn_Good;
    Var Churn_rate Total_Active Total_Deactivated;
    With Month_no;
	Title "Correlation Between Month Number and Churn Rate, Active and Deactivated Accounts For Good Credit Class";
Run;

Proc Corr data=Monthly_Churn_Bad;
    Var Churn_rate Total_Active Total_Deactivated;
    With Month_no;
	Title "Correlation Between Month Number and Churn Rate, Active and Deactivated Accounts For Bad Credit Class";
Run;
title;

symbol1 color=blue interpol=join; symbol2 color=red interpol=join;
proc gplot data=Monthly_Churn_Good; 
plot churn_rate*month_no / caxis=blue hminor=0          noframe;  
plot2 Total_Deactivated*month_no / caxis=red;  
Title "Good Credit: A comparison between Churn Rate trend and number of Deactivated Accounts over time";
run; 
quit;

symbol1 color=blue interpol=join; symbol2 color=red interpol=join;
proc gplot data=Monthly_Churn_Bad; 
plot churn_rate*month_no / caxis=blue hminor=0          noframe;  
plot2 Total_Deactivated*month_no / caxis=red;  
Title "Bad Credit: A comparison between Churn Rate trend and number of Deactivated Accounts over time";
run; 
quit;


proc sgplot data=Monthly_Churn_Good;
	reg x=month_no y=churn_rate / transparency=0.3 
                                  markerattrs=(symbol=CircleFilled color=black size=12px) 
                                  lineattrs=(color=green thickness=2);
  Title "Good Credit: Churn Rate Regression Trend Line over Time";
run;

proc sgplot data=Monthly_Churn_Bad;
	reg x=month_no y=churn_rate / transparency=0.3 
                                  markerattrs=(symbol=CircleFilled color=black size=12px) 
                                  lineattrs=(color=green thickness=2);
  Title "Bad Credit: Churn Rate Regression Trend Line over Time";
run;
title;

title1  ’Linear Regression of Total Deactivated Accounts and Month No for Good Credit Class; title2 ’(with 95% Confidence Limits)’;
symbol ci=red cv=blue co=gray value=dot       interpol=rlclm95 ;
proc gplot data=Monthly_Churn_Good;   
plot Total_Deactivated*month_no / regeqn; 
run; 
quit;

title1  ’Linear Regression of Total Deactivated Accounts and Month No for Bad Credit Class; title2 ’(with 95% Confidence Limits)’;
symbol ci=red cv=blue co=gray value=dot       interpol=rlclm95 ;
proc gplot data=Monthly_Churn_Bad;   
plot Total_Deactivated*month_no / regeqn; 
run; 
quit;

*======= Regression Models ====;
Proc Reg data=Monthly_Churn;
    Model Churn_Rate = Month_no;
    Title "Results of Regression Model of Monthly Churn (Overall)";
Run;

Proc Reg data=Monthly_Churn_Good;
    Model Churn_Rate = Month_no;
    Title "Results of Regression Model of Monthly Churn (Good Credit)";
Run;

Proc Reg data=Monthly_Churn_Bad;
    Model Churn_Rate = Month_no;
    Title "Results of Regression Model of Monthly Churn (Bad Credit)";
Run;



*What are the main reasons of deactivation? ;
Proc Tabulate data=Df5;
    Class DeactReason;    
    Table DeactReason;
Run;

Proc Tabulate data=Df5_Good;
    Class DeactReason;    
    Table DeactReason;
Run;

Proc Tabulate data=Df5_Bad;
    Class DeactReason;    
    Table DeactReason;
Run;

Proc sgplot data=Df5;
    Vbar DeactReason;
    Label ;
    Title2 'Frequency by Deactivation Reasons';
Run;

Proc sgplot data=Df5_Good;
    Vbar DeactReason;
    Label ;
    Title2 'Frequency by Deactivation Reasons of Good Credit Class';
Run;

Proc sgplot data=Df5_Bad;
    Vbar DeactReason;
    Label ;
    Title2 'Frequency by Deactivation Reasons of Bad Credit Class';
Run;


*=====Logistic Regrestion to predict deactivation;
DATA Df6 ;
  SET Df5 ;

	IF AcctStatus = 'Deactivated' THEN AcctStatus1 = 1; 
	    ELSE AcctStatus1 = 0;

	IF Region = '  West Provinces' THEN Region1 = 1; 
	    ELSE Region1 = 0;
	IF Region = ' Ocean Provinces' THEN Region2 = 1; 
	    ELSE Region2 = 0;
	IF Region = 'Central Provinces' THEN Region3 = 1; 
	    ELSE Region3 = 0;
RUN;

proc logistic data=Df6  descending; 
    model AcctStatus1 = Sales Tenure Age Region1 Region2 Region3 RatePlan GoodCredit;
run;


DATA Df6_Good ;
  SET Df5_Good ;

	IF AcctStatus = 'Deactivated' THEN AcctStatus1 = 1; 
	    ELSE AcctStatus1 = 0;

	IF Region = '  West Provinces' THEN Region1 = 1; 
	    ELSE Region1 = 0;
	IF Region = ' Ocean Provinces' THEN Region2 = 1; 
	    ELSE Region2 = 0;
	IF Region = 'Central Provinces' THEN Region3 = 1; 
	    ELSE Region3 = 0;
RUN;

proc logistic data=Df6_Good  descending; 
    model AcctStatus1 = Sales Tenure Age Region1 Region2 Region3 RatePlan;
run;

DATA Df6_Bad ;
  SET Df5_Bad ;

	IF AcctStatus = 'Deactivated' THEN AcctStatus1 = 1; 
	    ELSE AcctStatus1 = 0;

	IF Region = '  West Provinces' THEN Region1 = 1; 
	    ELSE Region1 = 0;
	IF Region = ' Ocean Provinces' THEN Region2 = 1; 
	    ELSE Region2 = 0;
	IF Region = 'Central Provinces' THEN Region3 = 1; 
	    ELSE Region3 = 0;
RUN;
 
proc logistic data=Df6_Bad  descending; 
    model AcctStatus1 = Sales Tenure Age Region1 Region2 Region3 RatePlan;
run;

title;
/*
======= EXCEL TEMPLATE ======
======= EXCEL TEMPLATE ======
======= EXCEL TEMPLATE ======
*/

/* output to word document with landscape orientation */
Title "Active Customers";
Proc tabulate data=df5 (where=(AcctStatus='Active')) order=formatted;
    class region province agegroup salesgroup acctno1/missing;    
    Table region=' '*province=' ' ALL, agegroup=' '*salesgroup='Number of Acct'*N=' ';
Run;


/* output to word document with landscape orientation */
Title "Deactivated Customers";
Proc tabulate data=Df5 (where=(AcctStatus='Deactivated')) order=formatted;
    class region province agegroup salesgroup acctno1/missing;    
    Table region=' '*province=' ' ALL, agegroup=' '*salesgroup='Number of Acct'*N=' ';
Run;



