import re

txt =input("enter mail id: ")
x = re.search("@gmail.com", txt)
if x is None:
    print(txt, "is not  valid")
else:
    print(txt,"is  valid")


num =input("enter number: ")
y = re.search("[a-zA-Z]", num)
if y is None:
    print(num, "is valid")
else:
    print(num,"is not valid")
