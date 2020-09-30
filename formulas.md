##### format setting in MailMerge field (insert before })
> \\@ YYYY-M-DD
> \\#,##0.00

#### max nd min in vba
> data_max = WorksheetFunction.Max(a, b, c, d)
> data_min = WorksheetFunction.Min(a, b, c, d)

##### round to even integer
> =IF((A1*10-INT(A1)*10)<=4,INT(A1),IF((A1*10-INT(A1)*10)>=6,INT(A1)+1,INT((INT(A1)+1)/2)*2))

##### show 1 to 10 then 10 to 1 over and over in column
> =ABS(TRUNC(9.5-MOD(ROW()+9,20)))+1

##### show financial Chinese of number in cell A1
> =TEXT(INT(RC[-1]),"[DBNum2]")&"元"&IFERROR(IF(MID(RC[-1],FIND(".",RC[-1])+1,1)*1,TEXT(MID(RC[-1],FIND(".",RC[-1])+1,1),"[DBNum2]")&"角","零"),"")&IF(RIGHT(TEXT(RC[-1],"0.00"),1)*1,TEXT(RIGHT(RC[-1],1),"[DBNum2]")&"分","整")
> =TEXT(INT(A1),"[DBNum2]")&"元"&IFERROR(IF(MID(A1,FIND(".",A1)+1,1)*1,TEXT(MID(A1,FIND(".",A1)+1,1),"[DBNum2]")&"角","零"),"")&IF(RIGHT(TEXT(A1,"0.00"),1)*1,TEXT(RIGHT(A1,1),"[DBNum2]")&"分","整")
