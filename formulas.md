#### calendar
> =IF(MONTH(DATE(YEAR(A1),MONTH(A1),1))<>MONTH(DATE(YEAR(A1),MONTH(A1),1)-WEEKDAY(DATE(YEAR(A1),MONTH(A1),1),2)+{1,2,3,4,5,6,7}+{0;1;2;3;4;5}*7),"",DATE(YEAR(A1),MONTH(A1),1)-WEEKDAY(DATE(YEAR(A1),MONTH(A1),1),2)+{1,2,3,4,5,6,7}+{0;1;2;3;4;5}*7)

##### format setting in MailMerge field (insert before })
> \\@ YYYY-M-DD
> \\#,##0.00

#### max nd min in vba
> data_max = WorksheetFunction.Max(a, b, c, d)
> data_min = WorksheetFunction.Min(a, b, c, d)

#### rank -continues
> {=SUM(IF(A$1:A$17>=A1,1/COUNTIF(A$1:A$17,A$1:A$17),""))}

##### round to even integer
> =IF((A1*10-INT(A1)*10)<=4,INT(A1),IF((A1*10-INT(A1)*10)>=6,INT(A1)+1,INT((INT(A1)+1)/2)*2))

##### show 1 to 10 then 10 to 1 over and over in column
> =ABS(TRUNC(9.5-MOD(ROW()+9,20)))+1

##### show financial Chinese of number in cell A1
> =TEXT(INT(RC[-1]),"[DBNum2]")&"元"&IFERROR(IF(MID(RC[-1],FIND(".",RC[-1])+1,1)*1,TEXT(MID(RC[-1],FIND(".",RC[-1])+1,1),"[DBNum2]")&"角","零"),"")&IF(RIGHT(TEXT(RC[-1],"0.00"),1)*1,TEXT(RIGHT(RC[-1],1),"[DBNum2]")&"分","整")

> =TEXT(INT(A1),"[DBNum2]")&"元"&IFERROR(IF(MID(A1,FIND(".",A1)+1,1)*1,TEXT(MID(A1,FIND(".",A1)+1,1),"[DBNum2]")&"角","零"),"")&IF(RIGHT(TEXT(A1,"0.00"),1)*1,TEXT(RIGHT(A1,1),"[DBNum2]")&"分","整")

#### Rounding
> =ROUND(A1,1)-(RIGHT(INT(A1*100),1)*1=5)*NOT(MOD(RIGHT(INT(A1*10),1),2))/10
