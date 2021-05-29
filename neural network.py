from numpy import exp, array, random, dot
import xlrd
import numpy as np
file = '大学排名.xlsx'

wb = xlrd.open_workbook(filename=file)
randomsheet = wb.sheet_by_index(0)

class NeuralNetwork():
    def __init__(self):

        self.synaptic_weights =  random.random((3, 1))
        # [[0.2], [0.5], [0.3]]
        # [[0.3], [0.4], [0.3]]

    def __sigmoid(self, x):
        return 1 / (1 + exp(-x))

    def __sigmoid_derivative(self, x):
        return x * (1 - x)


    def train(self, training_set_inputs, training_set_outputs, number_of_training_iterations):
        training_set_outputs = self.__sigmoid(training_set_outputs)
        for iteration in range(number_of_training_iterations):
            output = self.think(training_set_inputs)
            error = training_set_outputs - output
            adjustment = dot(training_set_inputs.T, error * self.__sigmoid_derivative(output))

            self.synaptic_weights += adjustment

    # The neural network thinks.
    def think(self, inputs):
        return self.__sigmoid(dot(inputs, self.synaptic_weights))




if __name__ == "__main__":

    neural_network = NeuralNetwork()

    print("Random starting synaptic weights: ")
    print(neural_network.synaptic_weights)

    training_set_inputs = array([[1,1,1]])
    for i in range(2,1719):
        a = array([[int(randomsheet.cell_value(i,1)),int(randomsheet.cell_value(i,2)),int(randomsheet.cell_value(i,3))]])
        training_set_inputs = np.r_[training_set_inputs,a]

    training_set_outputs = array([[1]])
    for i in range(2,1719):
        b = array([[int(randomsheet.cell_value(i,4))]])
        training_set_outputs = np.r_[training_set_outputs,b]    

    neural_network.train(training_set_inputs, training_set_outputs, 10000)

    print ("New synaptic weights after training: ")
    print( neural_network.synaptic_weights)
    
    # dot(inputs, self.synaptic_weights)

    # Test the neural network with a new situation.
    # print ("Considering new situation  -> ?: ")
    # print (dot([309,281,281],neural_network.synaptic_weights))

    # for i in range(2,1719):
    #     print (dot([int(randomsheet.cell_value(i,1)), int(randomsheet.cell_value(i,2)), int(randomsheet.cell_value(i,3))], neural_network.synaptic_weights))

