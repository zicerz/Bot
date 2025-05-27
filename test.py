def test():
    try :
        a = 5.0 / 1.0
        print('输出：我是try')
        return 0
    except :
        print('输出：我是except')
        return 1
    else :
        print('输出：我是else')
        return 2
    finally :
        print('输出：finally')
        return 3
print('test: ',test())