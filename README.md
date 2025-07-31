Excel-21

Project Description

Excel-21 is a guide to Analysis ToolPak in Microsoft Excel. Learn practical tips, step-by-step instructions, and see illustrated examples for: Histogram, Descriptive Statistics, Moving Average, Exponential Smoothing, Correlation Analysis, Regression Analysis.

Table of Contents

The Analysis ToolPak is an Excel add-in program that provides data analysis tools for financial, statistical and engineering data analysis.

To load the Analysis ToolPak add-in, execute the following steps.

1. On the File tab, click Options.
2. Under Add-ins, select Analysis ToolPak and click on the Go button.

![screenshot](Screenshots/Analysis1.png)

3. Check Analysis ToolPak and click on OK.

![screenshot](Screenshots/Analysis2.png)

4. On the Data tab, in the Analysis group, you can now click on Data Analysis.

![screenshot](Screenshots/Analysis3.png)

Histogram

1. First, enter the bin numbers (upper levels) in the range C4:C8.
2. On the Data tab, in the Analysis group, click Data Analysis.
3. Select Histogram and click OK.

![screenshot](Screenshots/Histogram1.png)

4. Select the range A2:A20
5. Click in the Bin Range box and select the range C4:C8.
6. Click the Output Range option button, click in the Output Range box and select cell F6.
7. Check Chart Output.

![screenshot](Screenshots/Histogram2.png)

8. Click OK.

Result:

![screenshot](Screenshots/Histogram3.png)

9. Click the legend on the right side and press Delete.
10. Properly label your bins.
11. To remove the space between the bars, right click a bar, click Format Data Series and change the Gap Width to 0%.
12. To add borders, right click a bar, click Format Data Series, click the Fill & Line icon, click Border and select a color.

![screenshot](Screenshots/Histogram4.png)

(Note: I changed the color to make it look more clear.)

(Note: If you have Excel 2016 or later, you can use the Histogram chart type.)

Descriptive Statistics

You can use the Analysis Toolpak add-in to generate descriptive statistics.

1. On the Data tab, in the Analysis group, click Data Analysis.
2. Select Descriptive Statistics and click OK.

![screenshot](Screenshots/Desc1.png)

3. Select the range A1:A15 as the Input Range.
4. Select cell C4 as the Output Range.
5. Make sure Summary statistics is checked.

![screenshot](Screenshots/Desc2.png)

6. Click OK.

Result:

![screenshot](Screenshots/Desc3.png)

Moving Average

A moving average is used to smooth out irregularities (peaks and valleys) to easily recognize trends.

1. On the Data tab, in the Analysis group, click Data Analysis.
2. Select Moving Average and click OK.

![screenshot](Screenshots/Moving1.png)

3. Click in the Input Range box and select the range B2:M2.
4. Click in the Interval box and type 6.
5. Click in the Output Range box and select cell B3.
6. Click OK.

![screenshot](Screenshots/Moving2.png)

Result:

![screenshot](Screenshots/Moving3.png)

7. Repeat steps 1-6 for interval = 2 and interval = 4
8. Plot a graph of these values.

Result:

![screenshot](Screenshots/Moving4.png)

Conclusion: The larger the interval, the more the peaks and valleys are smoothed out. The smaller the interval, the closer the moving averages are to the actual data points.

Exponential Smoothing

Exponential smoothing is used to smooth out irregularities (peaks and valleys) to easily recognize trends.

1. On the Data tab, in the Analysis group, click Data Analysis.
2. Select Exponential Smoothing and click OK.

![screenshot](Screenshots/Exp1.png)

3. Click in the Input Range box and select the range B2:M2.
4. Click in the Damping factor box and type 0.9. Literature often talks about the smoothing constant α (alpha). The value (1- α) is called the damping factor.
5. Click in the Output Range box and select cell B3.
6. Click OK.

![screenshot](Screenshots/Exp2.png)

Result:

![screenshot](Screenshots/Exp3.png)

7. Repeat steps 1 to 6 for alpha = 0.3 and alpha = 0.8.
8. Plot a graph of these values.

Result:

![screenshot](Screenshots/Exp4.png)

Conclusion: The smaller the alpha (and the larger the damping factor), the more the peaks and valleys are smoothed out. A larger alpha (and a smaller damping factor) results in the smoothed values being closer to the actual data points.

Correlation Analysis

The correlation coefficient (a value between -1 and +1) tells you how strongly two variables are related to each other.

1. On the Data tab, in the Analysis group, click Data Analysis.
2. Select Correlation and click OK.

![screenshot](Screenshots/Correlation1.png)

3. Select the range A1:C6 as the Input Range.
4. Check Labels in first row.
5. Select cell A8 as the Output Range.
6. Click OK.

![screenshot](Screenshots/Correlation2.png)

Result: 

![screenshot](Screenshots/Correlation3.png)

(Note: I plotted a graph of these values.)

Conclusion: variables A and C are positively correlated (0.91). Variables A and B are not correlated (0.19). Variables B and C are also not correlated (0.11). You can verify these conclusions by looking at the graph.

(Note: You can also use the built-in CORREL function.)

Regression Analysis

Is there a relationship between Quantity Sold (Output) and Price and Advertising (Input) in our example? Lets find out.

1. On the Data tab, in the Analysis group, click Data Analysis.
2. Select Regression and click OK.

![screenshot](Screenshots/Regression1.png)

3. Select the Y Range (A1:A8). This is the predictor variable (also called dependent variable).
4. Select the X Range(B1:C8). These are the explanatory variables (also called independent variables). These columns must be adjacent to each other.
5. Check Labels.
6. Click in the Output Range box and select cell A11.
7. Check Residuals.
8. Click OK.

![screenshot](Screenshots/Regression2.png)

Result with plotted graph:

![screenshot](Screenshots/Regression3.png)

R Square equals 0.962, which is a very good fit. 96% of the variation in Quantity Sold is explained by the independent variables Price and Advertising. The closer to 1, the better the regression line fits the data.

Significance F and P-values

To check if your results are reliable (statistically significant), look at Significance F (0.001). If this value is less than 0.05, you're OK.
Most or all P-values should be below 0.05. In our example this is the case. (0.000, 0.001 and 0.005).

Coefficients

The regression line is: y = Quantity Sold = 8536.214 -835.722 * Price + 0.592 * Advertising. In other words, for each unit increase in price, Quantity Sold decreases by 835.722 units.

Residuals

The residuals show you how far away the actual data points are from the predicted data points (using the equation).

   
