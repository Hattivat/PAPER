import random
from os import urandom

acceptable_chars = 'abcdefghjkmnqprstuvwx23456789ABCDEFGHJKMNQPRSTUVWX23456789'


def pass_gen(size=12):
    password = ''
    for character in range(size):
        randomness = urandom(1)
        password.join(acceptable_chars[randomness % 58])
    print(password)


def old_pass_gen(size=8):
    print(''.join(random.choice(acceptable_chars) for char in range(size)))

if __name__ == "__main__":
    pass_gen()
