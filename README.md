# Stock Analysis

## Summary
### The purpose of this project is to use VBA assist our friend with quickly identifying profitable stocks in 2017 and 2018.

## Results
## Stock Performance by Year
### The selected stocks had a much better year in 2017 than they did in 2018. In 2017, the selected stocks had an average return of 67.3%. These stocks were performing incredibly well and would have seemed like a sound investment. However, the average return in 2018 was an awful -8.5% loss. Despite the largely poor perfomance of these stocks in 2018, two stocks did have a high retun: RUN and ENPH. As we can see in the photos below, RUN and ENPH had a respective return of 5.5% and 129.5% in 2017, and 84% and 81.9% in 2018. These would be great stocks to recommend to our friend.

### 2017 Stock Gains
![2017 Stock Gain](https://github.com/BrianWegemann/stock-analysis/blob/main/Stock_Gains_2017.PNG)

### 2018 Stock Gains
![2018 Stock Gain](https://github.com/BrianWegemann/stock-analysis/blob/main/Stock_Gains_2018.PNG)

## Code and Execution
### The refactored code greatly increased the speed at which VBA ran through each ticker and calculate totalVolume, starting price and endind price. This was done by creating a loop that pulled everything together at once instead of slowly calculating the relevant data for each individual ticker, then repeating over and over for the others. The final refactored code that decreased the run time can be seen below. 

     '2a) Create a for loop to initialize the tickerVolumes to zero.
       
    For i = 0 To 11
    tickerVolumes(i) = 0
    
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
        For i = 2 To RowCount

        '3a) Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
         tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
         
           '3d Increase the tickerIndex.
            tickerIndex = tickerIndex + 1
              
               End If
    
    Next i

### This refactored code reduced the output times significantly and will really show its value on larger datasets as our friend plans to use it for. 
### 2017 Output Time
![2017 Output](https://github.com/BrianWegemann/stock-analysis/blob/main/VBA_Challenge_2017.PNG)

### 2018 Output Time
![2018 Output](https://github.com/BrianWegemann/stock-analysis/blob/main/VBA_Challenge_2018.PNG)

## Summary
### The refactored code allowed us to decrease the amount of time it took our code to run. This will be very useful for larger datasets that could take quite some time to run using the older code. A potential issue that could arise from the refactored code would be the issue of an improperly formatted dataset. Because this code is looking in specific columns and cells, a dataset with different column names or values in different cells could cause an error with this refactored code or return invalid results. 
