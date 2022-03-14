n = 1
for i in range(6):
    for j in range(i):
        print(n,end=' ')
        n = n+1
    for j in range(i+1):
        print('#',' ',end='')
    print()

    