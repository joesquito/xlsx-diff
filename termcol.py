def print_color_chart():
    for i in range(8):
        for j in range(32):
            code = i * 32 + j
            print(f'\033[38;5;{code}m {code:3}\033[0m', end=' ')
        print()

    print("\nGrayscale:")
    for i in range(232, 256):
        print(f'\033[38;5;{i}m {i:3}\033[0m', end=' ')
    print('\033[0m')

print_color_chart()
