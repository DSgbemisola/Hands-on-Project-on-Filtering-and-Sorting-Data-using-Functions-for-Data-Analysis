# Hands-on-Project-on-Filtering-and-Sorting-Data-using-Functions-for-Data-Analysis

This is a Hands-on Project on Filtering and Sorting Data using Excel Functions.

# Software Used in this Project
Excel Desktop and free ‘Excel for the web’ version.

# Datasets Used in this Lab
The first dataset used in this lab comes from the following source: https://dataplatform.cloud.ibm.com/exchange/public/entry/view/f8ccaf607372882403a37d9019b3abf4. This dataset is published by IBM, and includes fictitious customer demographics and sales data.

The second dataset used in this lab comes from the following source: https://www.kaggle.com/sudalairajkumar/indian-startup-funding under a CC0: Public Domain license.
Acknowledgement and thanks also goes to https://trak.in who were generous enough to share the data publicly for free.

The third dataset used in this lab is an internal dataset from IBM.

# Objectives

The objectives of this project are to:

1. Use the Filter and Sort tools
2. Use IF, IFS, COUNTIF, and SUMIF functions for data analysis
3. Use the VLOOKUP and HLOOKUP reference functions

# Task 1: Filtering Data

Using Auto Filters to filter data required th following steps.

1. Download the file Customer_demographics_and_sales_Lab6.xlsx. Upload and open it using Excel for the web.
   Dataset is available here: https://docs.google.com/spreadsheets/d/1mkQFyS_owXD873lcu0zo094e6cxxMXN7/edit?usp=sharing&ouid=101032132621933397345&rtpof=true&sd=true
2. Select any cell in the data, and click the Data tab, then click Filter.
3. Click the filter drop-down in column AG (Purchase_Status), and select Filter….
4. In the list, only select Frequent and click OK.
5. Click the filter drop-down in the column AG, and click Clear Filter From “Purchase_Status”.
6. Click the filter drop-down in column AE (T_Type), and select Filter….
7. In the list, only select Cancelled and click OK.
8. Click the filter drop-down in column AF (Purchase_Touchpoint), and select Filter….
9. In the list, only select Desktop and click OK.
10. On the Data tab, click Clear.

![image](https://github.com/user-attachments/assets/d12d11eb-c389-42b0-b2e4-479aa9a5ae7d)

- To use Custom Filters to filter data:

1. Click the filter drop-down in column AD (Order_Value), then Number Filters>Top 10….
2. Change the value from 10 to 50 and Click OK.
3. Click the filter drop-down in the column AD, and click Clear Filter From “Order_Value”.

# Task 2: Sorting data
1. On the Data tab, click Custom Sort to open a dialog box like below.
2. Click the Column drop-down of row Sort By, select Order_Ship_Date.
3. Click the Order drop-down of row Sort By, select Sort Ascending.
4. Click Add.
5. Click the Column drop-down of row Then By, select Order_Value.
6. Click the Order drop-down of row Then By, select Sort Descending.
7. Click OK.

![image](https://github.com/user-attachments/assets/60e6e3c4-956f-4376-92be-5aebc75c82d5)

# Task 3: Using IF to apply one condition

1. Select column AF, right-click, Insert.
2. In cell AF1, type Complete?.
3. In cell AF2, type =IF(AE2="Complete","Yes","No") and press Enter.
4. Double-click the Fill Handle of AF2 to copy down the column.

# Task 4: Using Nested IF to apply multiple conditions

1. Select column AE, right-click, Insert.
2. In cell AE1, type Order Size (IF).
3. In cell AE2, type =IF(AD2>300,"Large",IF(AD2>100,"Medium",IF(AD2>0,"Small"))) and press Enter.
4. Double-click the Fill Handle of AE2 to copy down the column.

# Task 5: Using IFS to apply multiple conditions (alternative of Nested IF)

1. Select column AE, right-click, Insert.
2. In cell AE1, type Order Size (IFS).
3. In cell AE2, type =IFS(AD2>300,"Large",AD2>100,"Medium",AD2>0,"Small") and press Enter.
4. Double-click the Fill Handle of AE2 to copy down the column.

# Task 6: Using COUNTIF to count the number of cells that meet a specified criterion

1. Select cell BX2 and type count VISA card.
2. Select cell BY2 and type:
 =COUNTIF(N2:N195,"VISA") and press Enter.

# Task 7: Using SUMIF function to sum the values within a specified range that meet a specified criterion

1. Select cell BX3 and type sum Large order.
2. Select cell BY3 and type =SUMIF(AE2:AE195,"Large", AD2:AD195) and press Enter.
   Formula: =SUMIF(range, criteria, [sum range]).

# Task 8: Using SUMIFS function to sum the values within a specified range that meet multiple specified criteria

1. Select cell BX4 and type sum Large order with Baby Gen.
2. Select cell BY4 and type =SUMIFS(AD2:AD195, AE2:AE195,"Large", AL2:AL195,"BABY_BOOMERS") and press Enter.
3. Formula: =SUMIFS ([sum range], range1, criteria1, range2, criteria2, …).

# Task 9: Using the VLOOKUP Function

1. Download the file indian_startup_funding_Lab6.xlsx. Upload and open it using Excel for the web.
   Dataset is available here: https://docs.google.com/spreadsheets/d/17g3dVM7hYr5msmfkwEtvc9lTEPrkC5YF/edit?usp=sharing&ouid=101032132621933397345&rtpof=true&sd=true
2. In cell K2,L2,M2, type VLOOKUP, Startup Name, Amount in USD respectively.
3. Select and copy cells from C9 to C15 and paste in cell L3.
4. In cell M3, type =VLOOKUP(L3, C2:I113, 7, FALSE) and press Enter.
5. Formula: =VLOOKUP (value, table, col_index, [range_lookup]).
6. Hover over the bottom-right corner of cell M3, and drag the Fill Handle down to the cell M9.
7. Select cells from M3 to M9 and select Number Format>Currency.

![image](https://github.com/user-attachments/assets/21029920-378a-40bb-bb6a-aa1e7a23027d)

# Task 10: Using the HLOOKUP Function

1. Download the file Personal_Monthly_Expenditure_Lab6.xlsx. Upload and open it using Excel for the web.
2. In cell J2,K2,L2,M2, type HLOOKUP, Month, Food & Dining, Health & Fitness respectively.
3. Select and copy cells from A10 to A12 and paste in cell K3.
4. In cell L3, type =HLOOKUP(D1, A1:H14, 10, FALSE) and press Enter.
   Formula: =HLOOKUP (value, table, row_index, [range_lookup]).
5. Hover over the bottom-right corner of cell L3, and drag the Fill Handle down to the cell L5.
6. Select cells from L3 to L5 and select Number Format>Currency.
7. In cell M3, type =HLOOKUP(G1, A1:H14, 10, FALSE) and press Enter.
8. Hover over the bottom-right corner of cell M3, and drag the Fill Handle down to the cell M5.
9. Select cells from M3 to M5 and select Number Format>Currency.

![image](https://github.com/user-attachments/assets/271c8426-0fc8-4f54-9fee-5433792b1f0c)

