import winsound

def Alarm(status):
    if status=='completed':
        M = (800,900,1000,900,800)
    elif status=='Error':
        M = (1400,400,1400,400)
        winsound.MessageBeep()
    elif status=='Beep':
        M = (2000,2000)
    elif status=='Treble':
        M = (1800,1000)
    elif status=='Start':
        M = (1100,1000,900,1000,1100)
    for k in M:
        winsound.Beep(k, 100)