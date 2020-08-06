print('Following choices are available: \n1 - function one, \n2 - function two , \n3 - function three, \nall - run all functions')

def rep1():
     print('hello')

def rep2():
    print('hello 2')

def wrapper():
    print('This is the wrapper func')

def all():
	print('Executing all funcs')
	rep1()
	rep2()
	wrapper()
	

dispatcher = {
    '1': rep1, '2': rep2, '3': wrapper, 'all': all
}

action = input('Option: - ')

dispatcher[action]()

