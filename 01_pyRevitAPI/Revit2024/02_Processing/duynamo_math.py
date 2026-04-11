#region capitalizeWords
def capitalizeWords(input_string):
    words = input_string.split()
    capitalizeWords = [word.capitalize() for word in words]
    return ' '.join(capitalizeWords)
#endregion
#region generate random list
import random
def genRandomList(length):
    return [random.randint(1, 10) for _ in range(length)]
#endregion
#region shuffle
def shuffleList(inList):
    random.shuffle(inList)
    return inList
#endregion