def decrypte(encrypted_str, encryption_key="EncryptionKey"):
    encrypted_numbers = encrypted_str.split(",")

    decrypted_str = ""
    t = 0

    for num in encrypted_numbers:
        u = ord(encryption_key[t % len(encryption_key)]) % 96
        r = int(num) - u

        if r < 32:
            r += 96

        decrypted_str += chr(r)
        t += 1

    return decrypted_str


encrypted_i = '78,124,105,33,89,113,40,117,126,124,121,104,33'

decrypted_part_2 = decrypte(encrypted_i)
print(f"Decrypted Email: {decrypted_part_2}")
