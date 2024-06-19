def roundToNearestFive(number):
    remainder = number % 5
    if remainder == 0:
        return number
    return number + (5 - remainder)

def processList(number):
	roundedNumber = []
	if number % 5 != 0:
		roundedNumber.append(roundToNearestFive(number))
	else:
		roundedNumber.append(number)
	return roundedNumber