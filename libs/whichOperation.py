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
        if(eachOperation.startswith("*") and eachOperation.endswith("*")):
            if (str(eachOperation)[1:-1] in str(myCell)):
                return True
            return False
        elif (eachOperation.startswith("*") and str(myCell).startswith(str(eachOperation[1:]))):
            return True
        elif (eachOperation.endswith("*") and str(myCell).endswith(str(eachOperation[:-1]))):
            return True
        else:
            return False