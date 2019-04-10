# def outer(func):
#     def inner():
#         print("before")
#         func()
#         print("after")
#     return inner
#
# @outer
# def f1():
#
#
# f1()

def outer():
    def inner():
        print("before")
        print("F1")
        print("after")
    return inner
