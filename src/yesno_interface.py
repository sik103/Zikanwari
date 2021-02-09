def yesno(msg, y_n0=None):
    ans = None
    while ans is None:
        if y_n0 is True:
            ans = True
            y_n = "[y]"
        elif y_n0 is False:
            ans = False
            y_n = "[n]"
        elif y_n0 is None:
            y_n = ""
        else:
            raise TypeError

        ans0 = input(msg + "(y/n)" + y_n + ":")
        if ans0 == "y":
            ans = True
        elif ans0 == "n":
            ans = False
        elif ans0 == "":
            pass
        else:
            ans = None
    return ans
