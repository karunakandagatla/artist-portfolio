"""Read two numbers and print the bigger one."""

first = float(input("Enter the first number: "))
second = float(input("Enter the second number: "))

if first > second:
    biggest = first
else:
    biggest = second

print("Original numbers:", first, "and", second)
print("Biggest number:", biggest)
