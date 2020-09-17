---
title:  "Dissecting PDF documents"
date:   2019-10-10 21:45:46 +0100
categories: python
excerpt: "Very useful PDF dissection Python code"
header:
  overlay_image: /assets/images/john-finn--RU1NEhEUfs-unsplash.jpg
  caption: <span>Photo by <a href="https://unsplash.com/@john_finn?utm_source=unsplash&amp;utm_medium=referral&amp;utm_content=creditCopyText">John Finn</a> on <a href="https://unsplash.com/s/photos/ireland?utm_source=unsplash&amp;utm_medium=referral&amp;utm_content=creditCopyText">Unsplash</a></span>
---
If youre looking to extract stuff from PDFs without having to do the same PITA manual cutting and pasting repeatedly, week in, week out, then this is the right blog post for you to read.

In this example, I have created a pretty simple piece of Python code to pull apart a multi-page PDF file that came from Group, and I use Python to write out the tables to files which I then use to feed into a consolidation spreadsheet.

```python
# MIP PDF Reader & Parser Python Code
filename_in = 'C:\\Temp\\MIP EMEA Ireland Aug-20.pdf'
directory_name = 'C:\\Temp\\'
from tabula import wrapper
tables = wrapper.read_pdf(filename_in,multiple_tables=True,guess = False,pages='7')
i=1
for table in tables:
    table.to_excel(directory_name+'Services-MIP-'+str(i)+'.xlsx',index=False)
    i=i+1
tables = wrapper.read_pdf(filename_in,multiple_tables=True,guess = True,pages='3')
i=1
for table in tables:
    table.to_excel(directory_name+'Overall-MIP-'+str(i)+'.xlsx',index=False)
    i=i+1
```
Simple but effective - saves me loadsa time every month!