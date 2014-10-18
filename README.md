# office-tools

`officetools` is a Python package that provides command-line tools for
handling Microsoft Office files (with file extension `.docx`). In particular,
it provides a script to extract the comments from a Microsoft Word file and
write them to a CSV file.

In addition to the `officetools` package, this distribution installs
`docx.py`, a command-line Python script to extract the comments from a
Microsoft Word file. `docx.py` can be used as follows, to extract the comments
from a Microsoft Word document (`draft_report.docx`) and write them to a CSV
file (`comments.csv`):

```
python docx.py extract -o comments.csv draft_report.docx
```
