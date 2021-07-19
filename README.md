# GIFT Moodle to XLSX Kahoot converter
This script is useful in order to migrate Moodle GIFT questions to a [Kahoot Quiz](https://create.kahoot.it/)

## Installation
You need to install some packages using `pip` from pypi.org: `pip install pygiftparserrgmf` and `pip install xlsxwriter`

## Usage
```
$ python3 gift-to-xslx.py -f questions-file.txt
```
It will generate a new `questions-file.txt.xslx` that can be directly imported into a Kahoot Quiz.
Only "True/False" and "Multichoice" questions are compatible.
