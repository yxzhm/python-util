

def logistic(x,t):
    r=0.2155
    max_x=443.99
    return np.array(r*x*(1-(x/max_x))+0*t)

time = np.arr