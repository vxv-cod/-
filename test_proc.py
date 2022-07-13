import multiprocessing
from multiprocessing import Queue 
import time

potokAll = 20
countelem = 100
elempot = [0]
for i in range(0, potokAll):
    xxx = round(countelem/potokAll)
    elempot.append(xxx)
    countelem = countelem - xxx
    potokAll = potokAll - 1
elempot = [sum(elempot[0:i+1]) for i in range(0, len(elempot))]
# print('elempot = ', elempot)

lock = multiprocessing.Lock()
queue = Queue() 
def potok_X(elem_1, elem_2):
    # lock = multiprocessing.Lock()
    # queue = Queue()  
    for i in range(elem_1, elem_2):
        time.sleep(0.5)
        # lock.acquire()
        queue.put(i)  
        print('hello world', i)
        # lock.release()
        queue.get()

if __name__ == '__main__':
    for i in range(1, len(elempot)):
        thr = multiprocessing.Process(target=potok_X, args=(elempot[i-1], elempot[i], ), name=f"thr-{i}")
        thr.start()