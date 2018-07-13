import os
import datetime

# def read_file(path, file_name):
#     file_full_path = path + '/' + file_name
#     if not os.path.isfile(file_full_path):
#         return None
#
#     with open(file_full_path, 'r') as f:
#         return f.read()
#
# s = read_file("file", "ped.txt").upper()
# print type(s)

today = datetime.datetime.now().strftime('%d%b%y').upper()
print today