# Homework_Assignment_2_Notes

# Methodolgoy

- For Sub Stock_Loop_2016, 2015, and 2014()

- For i = 2 To 797711 

Stock_Close = Cells(i, 6)    
        If Stock_Open = 0 Then
            Yearly_Change = 0
            Percent_Change = 0
        Else
            Yearly_Change = Stock_Close - Stock_Open
            Percent_Change = Round((Yearly_Change / Stock_Open),4)

- For year and percent change, I was having trouble with the (Year Change = Stock_Close - Stock_Open) and (Percent Change = (Stock_Close - Stock_Open)/Stock_Open * 100%)formulas. 
Particularly, when I tried executing the string of formulas as was originally, all the numbers came out as 0 and 0%s. 
So I asked around stockoverflow.com and was suggested to try: 

- If Stock_Open = 0 Then 
Yearly_Change = 0
Percent_Change = 0
  Else
Yearly_Change = Stock_Close - Stock_Open"

This "If everything equals zero statement" says "ok if all this equals zero, then it's zero. Unless if Yearly change Does not equal zero, then it equals Stock_Close non-zero value minus Stock open Non-zero value; and I end up with the answer.       

- Each 2016, 2015, and 2014 VBA Scripts co-respond to each respective year bucket in excel.

- Open Excel, click on any year bucket, open its respective VBA year script (2016 VBA scripts should only be executed in the 2016 bucket on Excel, and 2015 for 2015, and for 2014 as well.   
