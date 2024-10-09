try:
    print("B")
    import main.py
except Exception as error:
    print(error)
    input()
else:
    print("A")