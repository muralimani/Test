def is_prime(n):
    if n < 2:
        return False
    for i in range(2, int(n ** 0.5) + 1):
        if n % i == 0:
            return False
    return True

def largest_prime_below(limit):
    for num in range(limit - 1, 1, -1):
        if is_prime(num):
            return num
    return None

if __name__ == "__main__":
    print(largest_prime_below(100))