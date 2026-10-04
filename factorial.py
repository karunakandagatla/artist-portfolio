"""Read a number and print its factorial."""

number = int(input("Enter a number: "))

if number < 0:
    print("Factorial is not defined for negative numbers.")
else:
    factorial = 1
    for value in range(1, number + 1):
        factorial = factorial * value
    print(f"Factorial of {number} is {factorial}")
