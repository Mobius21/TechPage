## Welcome to My Technical Sample Pages

Here are some technical part from the following project I have participated before. 
- Barcode Transformation Project using Excel VBA
- College Raptor Group Capstone Project using R
- Iowa City Safest County Project using SQL

### Barcode Transformation Project

This project is created by [American Bear Logistics Inc.](https://www.americanbearlogistics.com/), the international logistic company located in Itasca, IL. It provides a one-stop service for both traditional and E-Commerce customers. Its business covers ocean, air, land shipping, and warehousing. It operates multiple warehouse facilities, bonded warehouses, and trailers. 
In the summer 2021, I was able to join the team as the intern of the Business Analyst. My first project was to find a way to help operators converting date, batch number and package number as a serial of Barcode using Code 128. I work a week and create a Macro file with VBA code that can use for all Barcode convertion projects.


```markdown
// **Excel VBA** code

Option Explicit
Dim lenList As Integer
Dim i As Integer, LastRow As Integer

Sub DataToBarCode()

    Call createHorizon
    
    Worksheets("Horizon").Activate
    Dim m As Integer, n As Integer, ct As Integer
    Dim LastCol As Integer
    Dim curr As String
    Dim arr() As String
    ct = 0
    With Range("A2")
        LastCol = lenList
        Debug.Print "Last Column is " & LastCol
        For i = 0 To LastCol
            LastRow = .Offset(1, i) - 1
            Debug.Print LastRow
            For m = 0 To LastRow
                ct = ct + 1
                ReDim Preserve arr(1 To ct)
                curr = .Offset(2 + m, i)
                arr(ct) = curr
            Next m
        Next i
    End With
    
    Worksheets("DASHBOARD").Activate
    
    Range("G:G").ClearContents
    Range("G1") = "Barcode"
    Range("G1").Font.Bold = True
    With Range("G1")
        For n = 1 To ct
            .Offset(n, 0) = arr(n)
        Next n
    End With
    
    LastRow = UBound(arr)
    Debug.Print LastRow
    
    'Stop
    
    Worksheets("TEMPLET").Activate
    Range("A:C").ClearContents
    
    Dim Remain As Integer, irow As Integer
    Remain = LastRow * 1
    ct = 0
    With Range("A1")
        Do While Remain > 0
            irow = LastRow - Remain
            For i = 0 To WorksheetFunction.Min(2, Remain - 1)
                .Offset(irow + 0, i) = "Originated U.S.A."
                .Offset(irow + 1, i) = Code128(arr(i + irow + 1))
                .Offset(irow + 2, i) = arr(i + irow + 1)
                ct = ct + 1
            Next
            With Range(.Offset(irow + 0, 0), .Offset(irow + 0, 2))
                .Font.Name = "OCRB"
                .Font.Size = 9
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
            End With
            With Range(.Offset(irow + 1, 0), .Offset(irow + 1, 2))
                .Font.Name = "Code 128"
                .Font.Size = 36
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
            End With
            With Range(.Offset(irow + 2, 0), .Offset(irow + 2, 2))
                .NumberFormat = "0"
                .Font.Name = "OCRB"
                .Font.Size = 9
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlTop
            End With
            
            Remain = Remain - WorksheetFunction.Min(3, Remain)
        Loop
    End With
    
    Call LastFormat
    
End Sub
   
Private Sub LastFormat()
    
    Worksheets("TEMPLET").Activate
    Columns("A:C").Select
    
    With Selection
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .Interior.Pattern = xlNone
        .Interior.TintAndShade = 0
        .Interior.PatternTintAndShade = 0
    End With
    
End Sub

Private Sub setupHorizon()
    
    Worksheets("Horizon").Activate
    Cells.Clear

    Worksheets("DASHBOARD").Activate
    lenList = Cells(Rows.Count, 1).End(xlUp).Row - 1
    'Debug.Print lenList
    
    Dim sourceRange As Range
    Dim targetRange As Range
    Set sourceRange = Worksheets("DASHBOARD").Range("A2", Range("A2").Offset(lenList - 1, 2))
    Set targetRange = Worksheets("Horizon").Range("A1")
    sourceRange.Copy
    targetRange.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:=False, Transpose:=True

End Sub


Private Sub createHorizon()
    
    Call setupHorizon
    
    Worksheets("Horizon").Activate
    Dim icol As Integer, irow As Integer
    Dim currDate As String, currBatch As Integer, currQuan As Integer, currVal As String
    
    With Range("A1")
        For icol = 0 To lenList
            currDate = .Offset(0, icol)
            currBatch = .Offset(1, icol)
            currQuan = .Offset(2, icol)
            
            For irow = 1 To currQuan
                currVal = CStr(Format(currDate, "YYYYMMDD")) & Format(currBatch, "00") & Format(irow, "000")
                'Debug.Print currVal
                .Offset(2 + irow, icol) = currVal
                .Offset(2 + irow, icol).NumberFormat = "0"
            Next irow
            
        Next icol
    End With

End Sub


```




### College Raptor Group Capstone Project 

The project is sponsored by [College Raptor Inc.](https://www.collegeraptor.com/), a company focused on helping high school students make smarter decisions on choosing their college. College Raptor helps students to find the colleges that fit their needs based on student preferences and personal information. In order to create a more accurate model for College Raptor, the team used data from Melissa. The Melissa API code listed below is part of the work used for the data preparation that gets 10 addresses based on one unique Zip+4.

```markdown
// **Rscript** code with syntax highlighting.

for (n in which(dffull$zpfprocess == "N")){
  tryCatch({
  URLZip = dffull$inputzpf[n]
  jd = 0
  
  while (jd < 10) {
    URL = paste(URLbase, URLZip, URLform,  sep = "")
    # This is the call to the Melissa API
    json_data = suppressWarnings(fromJSON(paste(readLines(URL), collapse="")))
    
    if (length(json_data)==0){
      zpfour[rc,]= c(dffull$inputzpf[n], NA, NA, NA, NA, NA)
      rc = rc+1
      break
    }
    else{
      limit = max(jd+length(json_data), 10)
      if (limit > 10) { a = 10 - jd } else { a = length(json_data) }
      for (i in (1:a)){
        zaddress = json_data[[i]]$Address
        zcityzip = paste(json_data[[i]]$City, json_data[[i]]$State, dffull$inputzpf[n], sep = ', ')
        
        ### add result to a temp data frame
        zpfour[rc,]= c(dffull$sourcezpf[n],
                       dffull$inputzpf[n],
                       json_data[[i]]$Address,
                       json_data[[i]]$City,
                       json_data[[i]]$State)
        ### update the row count
        print(c(rc, dffull$sourcezpf[n]))
        rc = rc+1
      }
    }
    
    jd = jd + length(json_data)
    zipfive <- substr(URLZip, 0, 6)
    zipfour <- as.numeric(substr(URLZip,7,10)) 
    URLZip = paste(zipfive, as.character(zipfour + 1), sep = "")
    dffull$inputzpf[n] = URLZip
  }
    
  dffull$zpfprocess[n] = "Y"
  zpfprocess[n] = "Y"
  
  date_time <- Sys.time()
  while((as.numeric(Sys.time()) - as.numeric(date_time)) < 2){} 
  }, error=function(e){})
}

```

### Iowa City Safest County Project ###

In this project, we assume to provide multiple choice for shelter for families who want to live in the safest county in Iowa.  According to the [Iowa Demographics by Cubic](https://www.iowa-demographics.com/cities_by_population), there are 1004 cities in Iowa, containing approximately 3.2 million residents. Our analysis will be driven by criteria including population, homicides, drug crime, traffic fatalities, income, etc. Our resulting database application will be posted on the website generated by **Oracle Apex** and will be useful for real estate agencies, travel agencies, school districts, police departments, and the public. The following SQL code is what we used to create one of the tables.

```markdown
\\SQL code with syntax highlighting.

CREATE TABLE "NON_SEX_DISEASE"
 (  "NON_SEXID" CHAR(50) NOT NULL ENABLE,
    "COUNTYID" CHAR(50) NOT NULL ENABLE,
    "TYPE_OF_NONSEX_DISEASE" VARCHAR2(50) NOT NULL ENABLE,
    "RATE_OF_NONSEXDISEASE" NUMBER,
    CONSTRAINT "NON_SEX_DISEASE_CON" PRIMARY KEY ("NON_SEXID")
    USING INDEX ENABLE
 )
/
ALTER TABLE "NON_SEX_DISEASE" ADD CONSTRAINT "NON_SEX_DISEASE_CON1"
FOREIGN KEY ("COUNTYID")
 REFERENCES "COUNTY" ("COUNTYID") ENABLE
/
```
One of our problem is to identify the counties with the lower mortality. Potential residents looking for a safe Iowa county to live in will prefer to find a location with a low mortality. However, the ratio of intentional to accidental deaths should be considered when we make a comparison. This information would also be helpful to government officials and police officers in deciding what areas have less need for police presence and examining successful departments and systems within counties with low mortality.

To increase the accuracy, we used a complex formula to get the mortality rate in each county and excluded counties that have mortality rate higher than average. Since the mortality rate we get is small and difficult to represent in percentage form, we used it to order the table from low to high by mortality rate.

```markdown
\\SQL code with syntax highlighting.

SELECT County_Name, Total_Death, TO_CHAR(Population,'999,999,999') as Population
FROM(
  Select County_Name, sum(Number_Of_Deaths) as Total_Death, Population
  FROM Death Join County on Death.CountyID = County.CountyID
  GROUP BY Death.countyID, County_Name, Population
  Having sum(Number_Of_Deaths)/Population <= (
    SELECT avg(Rate)
    FROM (
      SELECT Death.countyID, County_Name, Population, sum(Number_Of_Deaths)/Population as Rate
      FROM Death Join County on Death.CountyID = County.CountyID
      GROUP BY Death.countyID, County_Name, Population
      )
  )
  ORDER BY sum(Number_Of_Deaths)/Population
);


```

