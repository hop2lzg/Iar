import os

path = "zg" + '\\' + "day"
if not os.path.exists(path):
    os.makedirs(path)


# with open(path + '\\' + "test.csv", 'wb') as f:
#     print "start %s " % f.closed
#     print "start write %s " % f.closed
#     f.write("test")
#     print "end write %s " % f.closed

# print "end %s " % f.closed

try:
    f = open(path + 's\\' + "test.csv", 'wb')
    # print "start %s " % f.closed
    #
    # print "start write %s " % f.closed
    f.write("test")
    # print "end write %s " % f.closed
finally:
    if f:
        f.close()
    # print "start close %s " % f.closed
    #
    # print "end close %s " % f.closed

# print "end %s " % f.closed