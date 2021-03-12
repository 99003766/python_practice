nums = [1, 2, 3, 4, 5, 6, 7, 8]
evens = list(filter(lambda n:n%2 == 0, nums))
nums = [1, 2, 3, 4, 5, 6, 7, 8]
double = list(map(lambda n:n*2, evens))
print(double)