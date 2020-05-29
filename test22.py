def tt(mylist,**kwargs):
    for k,v in kwargs.items():
        mylist[k]=v
        print(mylist)
if __name__ == '__main__':
    tt([0,1,2,3],{1:4})