def whichOperation(myCell, eachOperation, myValue, tipo):
    # print("asdasd333333")
    if (tipo == "common"):
        # print("common")
        if (eachOperation.find("==") > -1):
            if (str(myCell) == str(myValue)):
                return True
            else:
                return False
        else:
            return eval(str(myCell) + eachOperation + str(myValue))

        # if (eachOperation.find("==") > -1):
        #     if (str(myCell) == str(myValue)):
        #         return True
        #     else:
        #         return False
        # elif (eachOperation.find(">=") > -1):
        #     if (myCell >= myValue):
        #         return True
        #     else:
        #         return False
        # elif (eachOperation.find("<=") > -1):
        #     if (myCell <= myValue):
        #         return True
        #     else:
        #         return False
        # elif(eachOperation.find("!=") > -1):
        #     if (myCell != myValue):
        #         return True
        #     else:
        #         return False
        # elif(eachOperation.find(">") > -1):
        #     if (myCell > myValue):
        #         return True
        #     else:
        #         return False
        # elif(eachOperation.find("<") > -1):
        #     if (myCell < myValue):
        #         return True
        #     else:
        #         return False
    elif (tipo == "re"):
        # print("Estamos en re")
        cant = 0
        # print(eachOperation[1:])
        if(eachOperation.startswith("*") and eachOperation.endswith("*")):
            if (cant != 1):
                cant = 1
                # print("Here we are")
                isIt = str(eachOperation)[1:-1] in str(myCell)
                # print(isIt)
                return isIt
        elif (eachOperation.startswith("*") and str(myCell).startswith(str(eachOperation[1:]))):
            if (cant != 1):
                cant = 1
                # print("start with")
                return True
        elif (eachOperation.endswith("*") and str(myCell).endswith(str(eachOperation[:-1]))):
            if (cant != 1):
                # print(str(myValue)[:-1])
                cant = 1
                # print("ends with")
                return True