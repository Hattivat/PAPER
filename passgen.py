import random

letters = 'abcdefghjkmnqprstuvwx'
passoptions = letters + letters.upper() + '123456789' * 2

def pass_gen(size=8):
    print(''.join(random.choice(passoptions) for char in range(size)))

if __name__ == "__main__":
    pass_gen()