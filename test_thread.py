import threading
import time


class MyThread(threading.Thread):
    def __init__(self, thread_id, name, counter):
        threading.Thread.__init__(self)
        self.thread_id = thread_id
        self.name = name
        self.counter = counter

    def run(self):
        print "running " + self.name + '\n'
        # thread_lock.acquire()
        execute(self.name, self.counter, self.counter)
        # thread_lock.release()


def execute(thread_name, delay, counter):
    print thread_name+"----begin"
    while counter:
        time.sleep(delay)
        print "%s: %s" % (thread_name, time.ctime(time.time()))
        counter -= 1
    print thread_name + "---end"

print "START"
threads = []
thread_lock = threading.Lock()
thread1 = MyThread(3, "T-1", 3)
thread2 = MyThread(2, "T-2", 2)

thread1.start()
thread2.start()

threads.append(thread1)
threads.append(thread2)

for t in threads:
    t.join()


print "OVER"