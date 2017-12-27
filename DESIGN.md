1. Use of Python
Using Python was much simpler than using C. The use of dictionaries was prominent in my code, so it was necessary to use a language whereby
the use of dictionaries is simple (instead of having to build various dictionaries in C using hash tables etc...)

2. Forgoing SQL
Instead of setting up a database to store the financial statements, I thought it was easier to simply create dictionaries for each financial
statement. Each time you reference a database using SQL, you have to create a dictionary to store the contents of what you retrieve. By storing
each financial statement in a list of dictionaries, I could directly pull information by indexing within.

3. Excel
Using Excel was the most efficient way of creating the DCF. Performing the calculations within Python would have been quite difficult. Excel has
certain built-in functions such as "Net Present Value" that simplified the process of making a DCF.
