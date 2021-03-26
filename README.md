## Stock-Analysis

# Green Energy Stocks

## Overview Of Project

### Purpose
We are helping Steve and his environmentally concious parents decide where to best invest in the green energy stock market

### Analysis
By creating Subroutines inside of a "macro", we are able to run code to simplify and speed up our findings on data/information that benefits our purpose

## Results

### Year by Year Analysis
![Screenshot (5)](https://user-images.githubusercontent.com/79877349/112691607-ccca7600-8e3a-11eb-8aae-fafe6a69d23d.png)
You can see that DAQO (DQ) was extremely successful in 2017, almost doubling what was invested. On the other side of things, Terraform Power (TERP) was incredibly unsuccessful at the same time, losing almost all investments.

![Screenshot (4)](https://user-images.githubusercontent.com/79877349/112691904-3cd8fc00-8e3b-11eb-8d2b-1d34c13ea083.png)
As for 2018, stocks were down across the board in Green Energy with the exception of Sunrun(RUN) and Enphase Enaergy (ENPH), who were thriving.

### Examples Of Code
This is the forLoop we used on the "AllStocksAnalysis" subroutine

   RowCount = Cells(Rows.Count, "A").End(xlUp).Row
   
   '4) Loop through tickers
   
      For i = 0 To 11
   
       ticker = tickers(i)
       
       totalVolume = 0
       
       '5) loop through rows in the data
       
       Worksheets(yearValue).Activate
       
       For j = 2 To RowCount
       
           '5a) Get total volume for current ticker
           
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           
           '5b) get starting price for current ticker
           
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If
           
This is the output arrays and the varibles we initialized in the "refactored" version

    Dim tickerVolumes(12) As Long
    
    Dim tickerStartingPrices(12) As Single
    
    Dim tickerEndingPrices(12) As Single
    
        For i = 0 To 11
        
      tickerVolumes(i) = 0
      
      tickerStartingPrices(i) = 0
      
      tickerEndingPrices(i) = 0
      
### Stock Comparison 2017 & 2018
When we look at our data for both years, we are able to see some major differences....There is a lot of success for investors in Green Energy as a whole in 2017, while in 2018, only a couple stocks were in the "black". (RUN) and (ENPH)

### Execution Times of Original and Refactored Subroutines

![VBA_Challenge_2018 png](https://user-images.githubusercontent.com/79877349/112693351-b671e980-8e3d-11eb-910d-968add455b84.png)
![VBA_Challenge_2017 png](https://user-images.githubusercontent.com/79877349/112693447-e5885b00-8e3d-11eb-90d9-deab6e954ac6.png)

## Summary
With all data considered, there is definately a downward trend forming as an overall stock category. That said, there is some promising stocks to be looked at in more detail. I reccomend doing more reseaerch on the two ((RUN) and (ENPH)) most recent successful stocks. Find out exactly what made them shine in a descending market. And then, make a more informed decision on the best place to invest.

### Advantages and Disadvantages
As far as refactoring the code, there is definately an advantage of automating the "macro" to do the analysis faster and more efficiently.
The disadvantage of creating  the "macro" is that some things might not run to your needs because of one misspelled word or misplaced varible/etc. It can get taxing on the brain trying to figure out where you went wrong.

### What are the pro's and con's of refactoring the original subroutine?
The original script ran just fine without refactoring. The refactored routine shaved just milliseconds off of the original. But on the other hand, we are able to access both sheets without having to dive in too much. 


