# stock-analysis

Module 2 | Assignment - Wall Street
Background Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.

Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.

_______________________________________________

Overview of Project: 
Steve oversees his parent's investing strategy. We (the company) were approached by Steve to make a program to automate his analysis. As Steve uses excel daily, a program(macro) in VBA made the most sense to our team. VBA is powerful enough to do simple calculations without needing additional software that may come with additional cost. As Steve may or may not be super excel savvy we created a macro that could be activated with the click of the button.

Results:

We made two separate versions of the macro to compare their performance against each other.

The first version of the macro (not refactored) took on average for both years about 0.50 seconds. The 2nd version (refactored) took on average about 0.10 seconds. Please see actual screenshots in repo for exact numbers. These are approximating and should only be used for reference.

before_refactor_2017.png and before_refactor_2018.png (see attached repo)
VBA_Challenge2017.png and VBA_Challenge_2018.png (see attached repo) for the refactored code
The main difference between the two macros is that one uses a "nested for loop" versus multiple "for loops" on the refactored code. There are pros and cons to each of course. Both macros rely on the use of "if" statements to loop through rows or columns. For loops are great for repetitive tasks such as the ones in this challenge assignment.

Summary:

In a summary statement, address the following questions. What are the advantages or disadvantages of refactoring code? How do these pros and cons apply to refactoring the original VBA script?

Based on our analysis, it seems like "nested for loops" can take more time for computer to run. It seems that using arrays will make a progam run faster. However, if computer performance is not a customer requirement, and concise code is one then the original code works better. In most cases, customers will likely be looking more for performance over visually appealing code. In nested for loops, you lose a control variable every time you use another nest another for loop. In our application, this prevented me from nesting for loops in the refactored code.

Pros Original / Cons Original Concise Code / Slow ------------/ lose control variables for each new nested for loop

Pros Refactor/ Cons Refactor Fast/ Lengthy Code Uses arrays / all control variables are available
