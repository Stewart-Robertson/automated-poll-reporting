# Multi-region-Polling-Cover-Sheet

## Background: The 2024 US Election

In the run-up to the 2024 US election, Redfield & Wilton Strategies conducted weekly polling in numerous Swing States. [The Telegraph](https://www.telegraph.co.uk) (a leading news outlet in the UK) required the results from this research to be sent as quickly as possible to allow them to capitalise on the coverage it would generate, especially as changing events make old results obsolete quickly.

So, I created [this script](https://github.com/Stewart-Robertson/Multi-region-Polling-Cover-Sheet/blob/main/Multi-region%20cover%20sheet.py) to quickly process an arbitrary number of polls (the Swing States polled changed), and present the results elegantly in Word document reports.

**Some features of the tabulated results:**

* Answers were bolded according to logic
* For questions of a certain format, net result rows were added to the bottom of the tables (see bottom image)
* For questions were a net result row was added, a further row was added with comparison to the previous result for that question in that State.

## The Data Source

Several (up to 10) Excel results files at a time were to be processed in the following format:

<img width="1482" alt="Screenshot 2025-06-02 at 12 44 11" src="https://github.com/user-attachments/assets/797d2ede-8481-47a5-8af5-6d764278b5d2" />


This data was normalised into columnar format using an excellent python script written by a former colleague, and then these tables were used in the script linked above.

## The Result

The answers from all sheets and all files were combined into a report with tables in the following format:

![image](https://github.com/user-attachments/assets/07622e70-3522-4f0a-8c1f-f447625d0959)

Note the addition of the "Net Result" and "Change from previous poll" rows at the bottom of the table, and the bolding applied to the answers.

N.B.: All the code written in the "normalise.py" file was not written by me.
