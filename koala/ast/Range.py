# from numpy import array, ndarray, asarray, multiply, divide

def check_array(array, index):
        try:
            return array[index]
        except:
            raise ValueError('Could\'t access indexes')

class Range(list):

    # def __new__(cls, input):
    #     return asarray(input).view(cls) # http://stackoverflow.com/questions/5149269/subclassing-numpy-ndarray-problem

    def add(self, other, index):
        if type(other) == Range:
            return check_array(self, index) + check_array(other, index)
        else:
            return check_array(self, index) + other

    def substract(self, other, index):
        if type(other) == Range:
            return check_array(self, index) - check_array(other, index)
        else:
            return check_array(self, index) - other

    def multiply(self, other, index):
        if type(other) == Range:
            return check_array(self, index) * check_array(other, index)
        else:
            return check_array(self, index) * other

    def divide(self, other, index):
        if type(other) == Range:
            return check_array(self, index) / check_array(other, index)
        else:
            return check_array(self, index) / other

    # not sure if this is needed:
    
    # def OR(self, other, index):
    #     if type(other) == Range:
    #         return check_array(self, index) or check_array(other, index)
    #     else:
    #         return check_array(self, index) or other

    # def AND(self, other, index):
    #     if type(other) == Range:
    #         return check_array(self, index) and check_array(other, index)
    #     else:
    #         return check_array(self, index) and other

    def is_equal(self, other, index):
        if type(other) == Range:
            return check_array(self, index) == check_array(other, index)
        else:
            return check_array(self, index) == other

    def is_not_equal(self, other, index):
        if type(other) == Range:
            return check_array(self, index) != check_array(other, index)
        else:
            return check_array(self, index) != other

    def is_strictly_superior(self, other, index):
        if type(other) == Range:
            return check_array(self, index) > check_array(other, index)
        else:
            return check_array(self, index) > other

    def is_strictly_inferior(self, other, index):
        if type(other) == Range:
            return check_array(self, index) < check_array(other, index)
        else:
            return check_array(self, index) < other

    def is_superior_or_equal(self, other, index):
        if type(other) == Range:
            return check_array(self, index) >= check_array(other, index)
        else:
            return check_array(self, index) >= other

    def is_inferior_or_equal(self, other, index):
        if type(other) == Range:
            return check_array(self, index) <= check_array(other, index)
        else:
            return check_array(self, index) <= other
    
    # operators[':'] = Operator(':',8,'left')
    # operators[''] = Operator(' ',8,'left')
    # operators[','] = Operator(',',8,'left')
    # operators['u-'] = Operator('u-',7,'left') #unary negation
    # operators['^'] = Operator('^',5,'left')
