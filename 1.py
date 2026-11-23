class Vehicle:
    def __init__(self, name, brand):
        self.name = name
        self.brand = brand

    def move(self):
        print("move on")


class Car(Vehicle):

    def move(self):
        print("Drive!")

class Plane(Vehicle):

    def move(self):
        print("Fly!")

class Boat(Vehicle):

    def move(self):
        print("Sail!")


car_1 = Car("Porsche", "911")
plane_1 = Plane("Boeing", "747")
boat_1 = Boat("boat", "110")

for x in (car_1, plane_1, boat_1):
    x.move()