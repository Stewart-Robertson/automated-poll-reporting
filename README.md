# Multi-region-Polling-Cover-Sheet
Python script to generate a clean word document with tables of results for multiple polls, including formatting logic.

Created to service the requirement to create neatly presented results from US Swing State polling and avoid countless hours of manual work.

Example of a single table created from the code, where 10 polls were simultaneously processed (note the bolding applied programmatically):

![image](https://github.com/user-attachments/assets/ca35133f-847c-4c5f-b733-8f0e244730b8)

The formatting logic was extended to tables for questions where "Neither" is present in one of the answer codes. Users are prompted with a list of answers, e.g. "['Strongly agree', 'Agree', 'Neither agree nor disagree', 'Disagree', 'Strongly disagree', "Don't know"]" and asked if the results are to be combined. If yes, the answers are formatted as such:

![image](https://github.com/user-attachments/assets/07622e70-3522-4f0a-8c1f-f447625d0959)

Note the addition of the "Net Result" and "Change from previous poll" rows at the bottom of the table.

N.B.: All the code written in the "normalise.py" file was not written by me.
