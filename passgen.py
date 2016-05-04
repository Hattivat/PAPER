import random
from os import urandom

acceptable_chars = 'abcdefghjkmnqprstuvwx23456789ABCDEFGHJKMNQPRSTUVWX23456789'


def pass_gen(size=12):
    password = ''
    for i in range(size):
        randomness = int.from_bytes(urandom(1), byteorder='little')
        character = acceptable_chars[randomness % 58]
        password += character
    print(password)


def old_pass_gen(size=8):
    print(''.join(random.choice(acceptable_chars) for char in range(size)))

if __name__ == "__main__":
    pass_gen()
