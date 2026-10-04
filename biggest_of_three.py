"""Read three numbers and print the biggest one."""

first = float(input("Enter the first number: "))
second = float(input("Enter the second number: "))
third = float(input("Enter the third number: "))

if first >= second and first >= third:
    biggest = first
elif second >= first and second >= third:
    biggest = second
else:
    biggest = third

print(f"Original numbers: {first}, {second}, and {third}")
print(f"Biggest number: {biggest}")
