<div align="center">

## Calculate the last day of the month


</div>

### Description

Determines number of days in a month, with leap year check.
 
### More Info
 
Month and year requested

Number of days in that month


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mike Collins](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mike-collins.md)
**Level**          |Beginner
**User Rating**    |4.3 (17 globes from 4 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Algorithims](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/algorithims__4-29.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mike-collins-calculate-the-last-day-of-the-month__4-6115/archive/master.zip)

### API Declarations

```
Copyright (c) 2000 Mike Collins.
You are hearby granted the right of non-exclusive use for any reason provided that you provide this copyright message in any work that uses this code.
```


### Source Code

Here is the VBScript Sample
----------------------------------------------
function GetLastDay( month, year )
 month = month + 1
 if month > 12 then
  month = month - 12
  year = year + 1
 end if
 Dim x
 x = DateAdd("d", -1, month&"/01/"&year)
 GetLastDay = Day( x )
end function
Response.Write GetLastDay( 3, 1999 )&"<br>"
Response.Write GetLastDay( 2, 1999 )&"<br>"
Response.Write GetLastDay( 2, 2000 )&"<br>"
----------------------------------------------
Here is the JavaScript Sample
----------------------------------------------
function GetLastDay( month, year )
{
 month++;
 if( month > 12 )
 {
  month -= 12;
  year++;
 }
 var x = new Date( month+"/1/"+year );
 x = new Date( x.valueOf()-(1000*60*60*24) );
 //1000*60*60*24=number of millisecs in 1 day
 return x.getDate();
}
Response.Write( GetLastDay( 3, 1999 )+"<br>" );
Response.Write( GetLastDay( 2, 1999 )+"<br>" );
Response.Write( GetLastDay( 2, 2000 )+"<br>" );

