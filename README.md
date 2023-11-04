
# Excel worksheet password remover

This project aims at creating a small Python program to remove eficently the password protection that can be applied on Excel worksheet. 
This tool is built around an small librairie that contains Excel document and worksheet classes.

The use Python makes this project cross-plateform, it can be use on Windows, MacOs and Linux.







## Installation

Install python 3.x : [Python installation page](https://www.python.org/downloads/windows/)

```bash
git clone git@github.com:valentinmlt/ExcelPasswordRemover.git

```
or 

Download and unzip git folder


## Features

- Remove password 
- Save for later the password protections
- Reapplied saved password 


```bash
python protect.py <file_name> <security_file>
python unprotect <file_name>
```

The python script for unprotecting worksheets will create a saving file that contains all security informations. 

This file will be used when using protect script don't throw it away. 

