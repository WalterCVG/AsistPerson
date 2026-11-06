Option Explicit

dim genano
dim genmes 
dim gdia
dim i 
dim unafecha
 

For i=1 to 7
  gdia=right("0"+trim(cstr(day(date+i-1))),2)    
  genano = trim(cstr(Year(date+i-1)))
  genmes =right("0"+ trim(cstr(Month(Date+i-1))),2)
  unafecha = gdia + "/" + genmes + "/" + genano
  msgbox(unafecha) 
  msgbox(WeekdayName(Weekday(unafecha )) + " " & gdia)
next

