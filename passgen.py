import random
from os import urandom

# Omitting potentially confusing characters, such as 0, O, o, i, l, I, 1
acceptable_chars = 'abcdefghjkmnqprstuvwx23456789ABCDEFGHJKMNQPRSTUVWX23456789'


def pass_gen(size=12):
    """ Generates a reasonably secure password using the OS's built-in RNG,
    converting the byte outputted by it into an integer and using that to pick
    one of the 58 characters in the 'acceptable_chars' string.
    """
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
