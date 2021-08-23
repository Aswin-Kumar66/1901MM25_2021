def is_digit(lis) :
    list1 = []
    for i in lis :
        if(type(i) != int) :
            list1.append(i)
    return list1

    

def get_memory_score(alist):
    lis = []    
    score  = 0
    for i in range(len(alist)) :
        if(alist[i] in lis) :
            score = score + 1
        elif((alist[i] not in lis) & (len(lis)<5) ) :
            lis.append(alist[i])
        else :
            for j in range(len(lis) - 1) :
                lis[j] = lis[j+1]
            lis[4] = alist[i]
    print('Score :',score)


input_nums = [3, 4, 1, 6, 3, 3, 9, 0, 0, 0]
if(len(is_digit(input_nums)) == 0) :
    get_memory_score(input_nums)  
else :
    print("'Please enter a valid input.Invalid inputs detected' :",is_digit(input_nums))

