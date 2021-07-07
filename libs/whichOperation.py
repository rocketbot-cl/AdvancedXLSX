def whichOperation(myCell, eachOperation, myValue, tipo):
    if (tipo == "common"):
        if (eachOperation.find("==") > -1):
            if (str(myCell) == str(myValue)):
                return True
            else:
                return False
        else:
            return eval(str(myCell) + eachOperation + str(myValue))

    elif (tipo == "re"):
        cant = 0
        if(eachOperation.startswith("*") and eachOperation.endswith("*")):
            if (cant != 1):
                cant = 1
                isIt = str(eachOperation)[1:-1] in str(myCell)
                return isIt
        elif (eachOperation.startswith("*") and str(myCell).startswith(str(eachOperation[1:]))):
            if (cant != 1):
                cant = 1
                return True
        elif (eachOperation.endswith("*") and str(myCell).endswith(str(eachOperation[:-1]))):
            if (cant != 1):
                cant = 1
                return True