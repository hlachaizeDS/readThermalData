from PIL import Image
import numpy


def readTiff(tifPath):
    im = Image.open(tifPath)
    imarray=numpy.array(im)
    return imarray



if __name__ == "__main__":

    tifPath="C:\\Users\\HL\\Desktop\\Prototype\\Thermal_Camera\\201909\\20190903_200629_ExtraWashes\\1\\20190903_200631_C1_BeforeAnything.tiff"
    imarray=readTiff(tifPath)
    print(imarray[50,159])
    print(imarray.shape)
