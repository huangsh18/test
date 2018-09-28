list = [5,1,7,3,9,6,2,8,4,10]
list_new = sorted(list)
print(list_new)
lens = len(list)
print(lens)
if lens%2 == 0:
    print((list_new[int(lens/2)]+list_new[int(lens/2-1)])/2)

else:
    print(list_new[int(lens/2)])
