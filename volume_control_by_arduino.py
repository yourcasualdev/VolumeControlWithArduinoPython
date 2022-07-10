import serial
import math
import numpy as np

from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


def main():
    # make sure the 'COM#' is set according the Windows Device Manager
    ser = serial.Serial('COM7', 9800, timeout=1)

    prevValue = 0

    while True:
        try:
            # get the current device
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(
                IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
            volume = cast(interface, POINTER(IAudioEndpointVolume))

            volRange = volume.GetVolumeRange()
            minVol = volRange[0]
            maxVol = volRange[1]

            # get the data from the serial port
            data = ser.readline()

            # convert the data to an integer
            if(len(data) == 6):
                # b'4.09\r\n' take 4.09 and convert it to int
                volumeData = float(data[0:4].decode('utf-8'))

                # rebase the volume to 0-100
                volumeData = float(volumeData * 20)

                # if change is bigger than 1, change the volume
                if(math.fabs(volumeData - prevValue) > 1):

                    set_vol = np.interp(volumeData, [0, 100], [minVol, maxVol])
                    prevValue = volumeData
                    if(volumeData > 95):
                        volume.SetMasterVolumeLevel(maxVol, None)
                    else:
                        volume.SetMasterVolumeLevel(set_vol, None)
                    print("Volume: ", volumeData)
        except:
            print("Error")
            pass


if(__name__ == "__main__"):
    main()
