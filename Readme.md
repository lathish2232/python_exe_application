# python_exe_application
Using python 3.9[tkinter] created Exe application
# This a Webscraping application used client Website URLS, I am Hiding all Urls becase it's not allowd to publish Client Website
website :- ![alt text](https://unificater.com/favicon.ico)  [unificater.com](https://unificater.com/#/)
contact us:-[Email:- contact@unificater.com](contact@unificater.com)


1. To convert python code[tkinter] to Exe i am using [cx-Freeze](https://pypi.org/project/cx-Freeze/) Module (Note:Tried with[pyinstaller](https://pypi.org/project/pyinstaller/) but not working as expected)
2. First install all pakages used in code including cx-freeze
3. for cx-Freeze we need to create setup.py file and sample setup.py file attached in repo
4. open Comand prompt and navigate to your file located dir then run bellow commands 
```CMD
python setup.py build
```
then it will create build folder inside that lib and exe application we be there 
Note:- copy all dependency files including images,logos and pate inside build folder Exe file location then only it will work 
