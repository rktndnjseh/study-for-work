class Person:
    def __init__(self, name):
        self.name = name

    def say_hello(self):
        print(f"Hi, I'm {self.name}!")

p = Person("Alice")
p.say_hello()
