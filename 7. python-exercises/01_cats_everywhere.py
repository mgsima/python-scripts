#Given the below class:
class Cat:
    species = 'mammal'
    def __init__(self, name, age):
        self.name = name
        self.age = age


# 1 Instantiate the Cat object with 3 cats
cat1 = Cat("Arya", 3)
cat2 = Cat("Sansa", 2)
cat3 = Cat("Salem", 400)


# 2 Create a function that finds the oldest cat
def oldest_cat(*cats):
    oldest = cats[0]
    for i in cats:
        if i.age > oldest.age:
            oldest = i
    return oldest

# 3 Print out: "The oldest cat is x years old.". x will be the oldest cat age by using the function in #2

print (f"The oldest cat is {oldest_cat(cat1, cat2, cat3).age} years old.")