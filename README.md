# Stock Analysis Report - Module 2/VBA

## Overview of Project
This challenge involved refactoring a VBA script using additional arrays and loops to see if it would run faster than before. 

## Results
The code before refactoring had the script first identify the ticker, then find the volume, then find the start, then find the end, then output the data. Then it would go to the next ticker. See screenshot below.
![Screenshot](https://user-images.githubusercontent.com/72076683/95664015-a1143780-0b09-11eb-8e82-478ff64337e9.png)

After refactoring, the script looped through all the tickers and found the volume, start and end price at the same time. Then it did another loop to out the data and then another to format the cells/color code them.

Prior to running the refactored code, the scripts ran in .742 and .753 seconds respectively for 2018 and 2017. After running the refactored code, the scripts ran in .234 and .140 seconds respectively for 2018 and 2017. This was an average (for both years) of a 75% reduction in the time it took to run the code. 
Here are sreenshots of both the 2018 and 2017 analysis and time it took to run the scripts.
![Screenshot](https://user-images.githubusercontent.com/72076683/95662039-9e114b00-0af9-11eb-8ffd-8bb3901bfc8d.png)
![Screenshot](https://user-images.githubusercontent.com/72076683/95662063-d749bb00-0af9-11eb-959f-ca69d9664d15.png)

## Summary
### What are the advantages or disadvantages of refactoring code?**
- **Advantages:** If you have a very long script, and it will be used multiple times, the advantage is time savings and processing savings. You don't want to send the spreadsheet to someone and they have a lot of files open and the script causes their computer to crash. You also want the script to run as quickly as possible so you can see the results faster.
- **Cons:** It takes time to refactor the script. It took many attempts for me to figure out how to do this. This would be especially difficult if I was not the original author of the code. It is best to write the original script as efficiently as possible.

### How do these pros and cons apply to refactoring the original VBA script?
I can imagine that this spreadsheet could be used by Steve in future years, all he would have to do now is add another sheet for 2019 performance and so on without many performance impacts. It could also be refactored to not hard code the names of the indexes but to look those up. That would be an added benefit of refactoring, if he wanted to look at other stocks in future years.
