# module-2-challenge-assignment

Overview of Project: Explain the purpose of this analysis.

This project entailed refactoring the VBA Script from the Green Stocks Analysis spreadsheet, using arrays instead of variables. The purpose of this project was to compare the elapsed time of completing the analysis of stocks, between the refactored script versus the original one.



Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

There were a couple of key differences between the refactored script versus the original script. We computed Total Volumes as well as Starting & Ending Prices for each stock ticker, using arrays rather than variables. Additionally, we used an index variable called tickerIndex, in order to call the correct stock ticker.
![image](https://user-images.githubusercontent.com/93381221/140629796-d3951a42-9495-4691-b2af-27cdb94debaa.png)

As a result of this method, the For loops which were used to compute the Volume, Starting Price, and Ending Price contained arrays instead, and each array used the index variable to call the stock ticker. ![image](https://user-images.githubusercontent.com/93381221/140629806-26880efe-dfb2-497f-b8bc-d10373e952ec.png)
![image](https://user-images.githubusercontent.com/93381221/140629813-079be449-7508-402c-b6c0-7453489980ff.png)

This is different from the original script, which contained variables in the For loops and called the stock ticker using the iterator variable itself.![image](https://user-images.githubusercontent.com/93381221/140629825-987d0352-62ff-441f-9241-0632ab3d86b8.png)

From the completed analysis, we can see in 2017 that all but one of the stocks had a positive return on value. The only negative return was from TERP, while the rest of the stocks showed a significant increase in value – a few even greater than 100%.![image](https://user-images.githubusercontent.com/93381221/140629830-a396a698-7cb2-48e6-b28a-76091516c784.png)

In 2018 on the other hand, we can see almost the exact opposite result. All but two of the stocks yield a negative return, and many of these decreases in value are over 20%. The two positive returns are very significant increases – over 80% for both.![image](https://user-images.githubusercontent.com/93381221/140629832-c28cf5bf-fdb5-469d-b8f7-0c72198df294.png)

Although the end result was the same for both years from both methods of VBA Script, there was a noticeable difference between the two methods in elapsed time for completing the analysis.
![image](https://user-images.githubusercontent.com/93381221/140629840-73117285-031c-46b3-af57-737310a76707.png)
![image](https://user-images.githubusercontent.com/93381221/140629843-0420f75f-164a-4c70-acde-561da7f729c7.png)
![image](https://user-images.githubusercontent.com/93381221/140629846-cef9b106-1dff-43c6-b197-468705f4ab49.png)
![image](https://user-images.githubusercontent.com/93381221/140629847-3f1e4b13-352f-4736-b280-ee88ebdef8c7.png)
The analysis appears to take slightly longer using the refactored script, as opposed to the original one.




Summary: In a summary statement, address the following questions.

1.	What are the advantages or disadvantages of refactoring code?
One key advantage of refactoring code may include shortening the code to execute the same exact task using better logic, less memory, and fewer lines. However, one potential disadvantage is that the original code might actually be faster, so refactoring the code may result in a slight or significant increase in elapsed time.

2.	How do these pros and cons apply to refactoring the original VBA script?
The refactored script uses better logic, because each of the key metrics being computed – volume, start price, end price – are stored in arrays. This is a better approach than using variables, because now we can reference a specific metric in our code directly. For example, if we wanted to call the Total Volume for DQ stocks in our code, we can now do that easily by referencing the array using its appropriate index, which for DQ would be tickerVolumes(2). This would not be possible using variables as done in the original script.
However, the refactored code does take additional elapsed time to run, by approximately one tenth of a second. This is likely because of the Index variable used to call the stock ticker. Because tickerIndex begins with a value of zero, we must add another if-then statement in our inner For loop, in order to increase the index value by 1, before the next iteration.
![image](https://user-images.githubusercontent.com/93381221/140629884-512f4a14-7523-4926-b5ae-bde09e42b506.png)
This step is necessary because our 4 arrays within both For loops use this index variable to call a stock ticker, and without this step, the index would remain at zero for the entire For loop. However, if we were to use the iterator as the index variable, as done in the original script, this step would be covered automatically and there would be no need for this additional if-then statement.
