from datetime import date

dates1 = '3/12/2020'
dates2 = '4/12/2020'

f_date = date(2020, 3, 11)
l_date = date(2020, 4, 1)
delta = l_date - f_date
print(delta.days)