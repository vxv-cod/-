from PIL import Image
path = r"C:\vxvproj\tnnc-StaticProcess\StatProc\excel.png"
filename = path
img = Image.open(filename)
img.save('logo.ico')

'''При желании вы можете указать нужные размеры значков:'''
icon_sizes = [(16,16), (32, 32), (48, 48), (64,64)]
# img.save('\icon\logo.ico', sizes=icon_sizes)
img.save(r"C:\vxvproj\tnnc-StaticProcess\StatProc\logo.ico", sizes=icon_sizes)
# ====================

# import imageio

# img = imageio.imread('123.png')
# imageio.imwrite('logo1.ico', img)