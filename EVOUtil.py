from PlayActions import send_keys as send

TAS = {
        "INA" : ("\\\\ERP\\dbamfg\\T7INA.RWN", 10),
        
        "DCD" : ("\\\\ERP\\dbamfg\\T7DCD.RWN", 3),
        "WOLE" : ("\\\\ERP\\dbamfg\\T7WOLE.RWN", 3),
        }

def openTASProgram(program):
    send(["focus EVO ~ ERP", 1, "alt m z u a", 1, TAS[program][0], "enter", TAS[program][1]])