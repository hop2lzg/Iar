import threading


def execute_timer(is_true):
    if is_true:
        print "Hi,I'm timer."
    else:
        print "Hi,You're timer."

print "START"
timer = threading.Timer(0, execute_timer, [False])
timer.start()
print "OVER"
