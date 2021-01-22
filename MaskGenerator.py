from lib_viewer import read_excel, distance_mm
import cv2
import numpy as np
import os

# tuto https://gist.githubusercontent.com/abhirooptalasila/8e46033b1190279040060588f44d8dc1/raw/0e39ebf52857468d223298f0600a380e75a34f00/maskGen.py

excel_file = 'C:/Users/xle/Bureau/XL/Privé/Thèse/LungAI/cfg/COVID.xls' #'C:/Users/xle/Downloads/EXL2.xlsx'
input_directory = 'C:/Users/xle/Bureau/mask/images/'
output_directory = 'C:/Users/xle/Bureau/mask/masks/'
img_nbr = 1  # 945

# lecture du fichier Excel
cursor = read_excel(excel_file)
i = 0
while i < img_nbr:
    i = i + 1
    png_file = cursor[i][19]    # nom du fichier pour lequel il faut générer le masque
    print(png_file)

    # recupération des coordonnées de la première anomalie
    x1 = int(float(cursor[i][21]))
    y1 = int(float(cursor[i][22]))
    x2 = int(float(cursor[i][23]))
    y2 = int(float(cursor[i][24]))

    x0 = x1
    y0 = x2
    dx = abs(x1-x2)    # vérifier si ce n'est pas l'inverse
    dy = abs(y1-y2)

    #os.chdir(input_directory)
    # generation du masque    mask_creation(x1,y1,x2,y2)
    src = cv2.imread(input_directory + png_file)
    mask = np.zeros_like(src)
    cv2.rectangle(mask,(y1,y2 ),(x1,x2),(255,255,255),thickness=2)    # vérifier si il ne faut pas inverser dx et dy
    cv2.imwrite(output_directory + png_file, mask)











