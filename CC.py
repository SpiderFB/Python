import hashlib
# print(hashlib.sha256("Hello".encode()).hexdigest())
NONCE_LIMIT = 10000000000
zeros = 5
def mine(blok_number, transactions, previous_hash):
    for nonce in range(NONCE_LIMIT):
        base_text = str(block_number) + transactions + previous_hash + str(nonce)
        hash_try = hashlib.sha256(base_text.encode()).hexdigest()
        if hash_try.startswith('0' * zeros):
            print(f"Found hash with Nonce: {nonce}")
            return hash_try
    return -1
block_number = 24
transactions = "83262391hfs91"
previous_hash = "92983h12u420hghn"

mine(block_number, transactions, previous_hash)

# combined_text =  str(block_number) + transactions + previous_hash + str(107617)
# print(hashlib.sha256(combined_text.encode()).hexdigest())