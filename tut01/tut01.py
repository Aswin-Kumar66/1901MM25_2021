def meraki_helper(n) :

        lis = []
        while(n != 0) :
             no = n%10
             lis.append(no)
             n = int(n/10)
        for i in range(len(lis) - 1) :
         if abs((lis[i] - lis[i+1])) != 1 :
             return False
        return True

    

input = [12,14,56,78,98,54,678,134,789,0,7,5,123,45,76345,987654321]
n = len(input)
count = 0
for i in range(n) :
     if(meraki_helper(input[i])) :
       print('Yes -',input[i], 'is a meraki number')
       count = count + 1
     else :
       print('No  -',input[i], 'is  not a meraki number')

print('the input list contains',count,'meraki and' ,n-count,'non meraki numbers.')

