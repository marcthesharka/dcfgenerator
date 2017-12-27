A discounted cash flow (DCF) is a valuation method used to estimate the attractiveness of an investment opportunity.
DCF analysis uses future free cash flow projections and discounts them to arrive at a present value estimate,
which is used to evaluate the potential for investment.

My program aims to automate the process of generating a DCF model in Excel. Powered fully by Python, it retrieves financial
information from the MorningStar API and from user inputted metrics, and then builds an Excel DCF model with those inputs.

Code Walkthrough:
The first part of the code is aimed at creating a form for users to provide their inputs. The HTML of the form can be found in
form.html. This form is connected to the first app route in application.py, "app.route("/")". Once the user fills out the form
on the webpage and clicks submit, when app.route("/") is reached via POST, the user's inputs are stored in global variables using
"sessions['variable name']". Once user input is stored, application.py calls on functions in helpers.py.

The three functions in helpers.py are lookup functions (similar to CS50 Finance). Three lookup functions exist, lookupis (lookup income statement),
lookupbs (lookup balance sheet), and lookupcf(lookup cash flow). Each lookup function calls on a url that provides a CSV for each financial statement.
Helpers.py skips the first two lines of the CSV (they are irrelevant) and then proceeds to iterate line by line through the CSV and stores the contents
into a list of python dictionaries. Each dictionary within the list is a unique line of the financial statement, with the contents of the dictionary being
each line item (aka revenue for year1, revenue for year2 etc...). Each dictionary (incomestatementdict, balancesheetdict, cashflowdict) is displayed
on inputs.html using Jinja to iterate through the dictionary and print the line item.

If the data displayed in the financial statements seem accurate, the user can click submit at the bottom of the page to input those values into the
DCF Model. The second part of application.py is on app.route("/inputs"). The loops at the beginning are simply iterating through the lines to find the
index of specific line items. For example, the first loop iterates through incomestatementdict to find the index for EBITDA, then stores that index
to be used later in the code (to find EBIT).

Once all of the relevant data has been retrieved from the financial statements, using xlsxwriter python library, I build the DCF model with the inputs.
Once the code has finished running, you can view the DCF model.

The final page, result.html, was meant to display the stock price determined by the DCF model, however I have not yet found a way to read the value of
the cell in Excel versus the formula. At the moment, it displays the formula.

I learned that the DCF Model is quite sensitive, and it is difficult to apply a "one-size fits all" template to valuing companies. However my project
serves as a base for me to continue building accuracy on the DCF model, by adding more inputs and creating higher complexity DCFs. In future, I plan
to continue reviewing the model, trying to make it as accurate as possible.
