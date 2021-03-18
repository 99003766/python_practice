import re

txt = "gjsimhareddy.39@gmail.com"
x = re.search("@gmail.com", txt)
print(x, "is valid")
num = "9010365337"

y = re.search("[a-zA-Z]", num)

print(y, "not valid")
