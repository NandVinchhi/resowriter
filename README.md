# MUN Resolution Writer
This is an open-source tool to make drafting resolutions in MUN committees far easier. It is a python-based desktop application made with [PyQt5](https://pypi.org/project/PyQt5/) and [python-docx](https://python-docx.readthedocs.io/en/latest/) library. It takes in raw text data and generates a .docx file which is fully formatted and ready to submit. I made this tool because I have personal experience of my entire resolution being scrapped due to formatting errors XD. You can check out the main.py file for the source code, or simply run the standalone .exe file to use this app. 
__Here is the google drive link to just the .exe file:__
https://drive.google.com/drive/folders/1rp1FyKAJyAZCf6q5l6CnbLLZleMt8oae?usp=sharing

### How to use (IMPORTANT)

* First enter in the title of you resolution, your committee and the topic ... this should be pretty straight-forward.

![](pics/p1.png)

* Enter the sponsor and signatory nations, separated by a comma and a space. Do not add a comma or a semi-colon at the end of the line. __Do not separate into multiple lines.__

![](pics/p2.png)

* Enter the preambulatory clauses __One per line__ and separated by a comma. There needs to be atleast 1 preambulatory clause and each clause needs to contain atleast 2 words.

![](pics/p4.png)

* Operative clauses:
  * Enter the operative clauses __One per line__ and separated by a semicolon, or comma, depending on the situation. There needs to be atleast one operative clause and each operative clause needs to have atleast 2 words. 
  * If the clause is a sub-clause to a previous clause, add in a single asterisk (*), followed by a space, before it.
  * If it is a sub-sub-clause to a previous sub-clause, add in two asterisks (\**), followed by a space, before it.

![](pics/p6.png)

* Select the required formatting for the keywords for both preambs and operatives.

![](pics/p3.png)

### Output of the above example:
![](pics/output.png)

### Contributing guidelines
Any PRs are always welcome! The immediate things to add would be improving the error checking (right now it just checks for the fullstops and other simple things), as well as adding some indentation on the UI side for clauses and sub-clauses.
