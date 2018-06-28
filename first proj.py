# from pillow import image
# note that you need an image saved into the same area
import time
from tkinter import *
from tkinter import _setit
import PIL
from PIL import Image
from PIL import ImageTk

screen = Tk()
screen.title("One Point Invoice")
screen.geometry("1500x900")
screen.iconbitmap('onepoint_logo.ico')
screen.configure(background='white')

# invoice title
invoicelbl = Label(screen, text="INVOICE", font='Helvetica 18 bold')
invoicelbl.config(font=("INVOICE", 14))
invoicelbl.grid(row=1, column=8)

# company adress label and widget
compad = Entry(screen, relief=SOLID)
compad = Text(screen, height=6, width=30)
compad.insert(END, "One Point Consulting Ltd\nAlpha House, Unit 14\n100 Villers Road\nLondon\nNW2 5PJ")
compad.grid(row=4, column=3)

# project code label and widget
projcodelbl = Label(screen, text="Project Code", bg="lightskyblue")
projectcode = Entry(screen, relief=SOLID)
projcodelbl.grid(row=6, column=8)
projectcode.grid(row=7, column=8)

# W.O.NUMBER
wonumlbl = Label(screen, text="W.O.NUMBER", bg="lightskyblue")
wonum = Entry(screen, relief=SOLID)
wonumlbl.grid(row=6, column=6)
wonum.grid(row=7, column=6)

# Terms label and widget
termslbl = Label(screen, text="Terms", bg="lightskyblue")
terms = Entry(screen, relief=SOLID)
terms.insert(END, "30 days")
termslbl.grid(row=6, column=7)
terms.grid(row=7, column=7)

# tax date label and widget
tdlbl = Label(screen, text="Tax Date")
td = Entry(screen, relief=SOLID)
td.insert(END, time.strftime("%d/%m/%Y"))
td.config(state='readonly')
tdlbl.grid(row=2, column=7)
td.grid(row=3, column=7)

# VAT reg widget
vatno = 866120039
vatreglbl = Label(screen, text="VAT Registration no.")
vatreg = Entry(screen, relief=SOLID)
vatreg.insert(END, vatno)
vatreg.config(state='readonly')
vatreglbl.grid(row=2, column=6)
vatreg.grid(row=3, column=6)

# Bank details part 1
bdlbl = Label(screen, text="Bank details")
bd = Entry(screen, relief=SOLID)
bd = Text(screen, height=7, width=50)
bd.insert(END,
          "Sterling Account\n\nBank Adress: HSBC, 1 The Town,EN2 6LD,UK.\n\nAccount: One Point Consulting Limited\n\nSort Code: 40-18-35.   Account Number:61237039")
bd.config(state='disabled')
bdlbl.grid(row=12, column=3)
bd.grid(row=13, column=3)

# bank details part 2
banlbl = Label(screen, text="BAN:")
ban = Entry(screen, relief=SOLID)
ban.insert(END, "GB74HBUK40183561237039")
banlbl.grid(row=14, column=2)
ban.grid(row=14, column=3)

# bank details part 3
swiftcodlbl = Label(screen, text="swiftcode:")
swiftcod = Entry(screen, relief=SOLID)
swiftcod.insert(END, "HBUKGB4B")
swiftcodlbl.grid(row=15, column=2)
swiftcod.grid(row=15, column=3)


# Product table
class Product:
    def __init__(self, Name, Price, Vat, Desc, Qty):
        self.NAME = Name
        self.PRICE = Price
        self.VAT = Vat
        self.DESC = Desc
        self.QTY = Qty


CAR = Product("Car", 3000, 1, "A LARGE METALLIC OBJECT THAT LIKE DRIVES AND STUFF", 1)

value = BooleanVar()
VAT = 0
NAME = Entry(screen)
QTY = Entry(screen)
PRICE = Entry(screen)
NAME.insert(0,"[Type Product here]")
QTY.insert(0,0)
PRICE.insert(0,0)
VATbutton = Checkbutton(screen, text="VAT?", variable=value)
VATAMOUNT = VAT * (int(PRICE.get()) * 0.2)
VATAMOUNTLBL = Label(screen,text = VATAMOUNT)
VATPERCENTlablel = Label(screen, text=str(VAT * 20) + "%")
VATlablel = Label(screen, text="£" + str(round(int(VATAMOUNT),2)))
finalpricelabel = Label(screen, text="£" + str(round(((int(PRICE.get()) * int(QTY.get())) + int(VATAMOUNT)), 2)))


def updateTable():
    VAT = value.get()
    VATAMOUNT = VAT * ((int(PRICE.get()) * int(QTY.get())) * 0.2)
    VATAMOUNTLBL['text'] = VATAMOUNT
    VATPERCENTlablel['text'] = str(VAT * 20) + "%"
    VATlablel['text'] = "£" + str(round(int(VATAMOUNT), 2))
    finalpricelabel['text'] ="£" + str(round(((int(PRICE.get()) * int(QTY.get())) + int(VATAMOUNT)), 2))


Confirmbutton = Button(screen,text = "Calculate",command = updateTable)
Confirmbutton.grid(row = 12, column = 7)


NAME.grid(row=9, column=5)
QTY.grid(row=9, column=4)
VATPERCENTlablel.grid(row=9, column=8)
VATlablel.grid(row=10, column=7)
PRICE.grid(row=9, column=6)
finalpricelabel.grid(row=11, column=7)
VATbutton.grid(row=10, column=4)


# edit tops:

QTYtop = Label(screen, text="QTY")
producttop = Label(screen, text="DESCRIPTION")
Vattop = Label(screen, text="VAT %")
Pricetop = Label(screen, text="PRICE")
vattottop = Label(screen, text="VAT TOTAL:")
tottop = Label(screen, text="TOTAL:")

QTYtop.grid(row=8, column=4)
producttop.grid(row=8, column=5)
Vattop.grid(row=8, column=8)
Pricetop.grid(row=8, column=7)
tottop.grid(row=11, column=6)
vattottop.grid(row=10, column=6)

# invoice number maker
file = open("file.txt", "r")
a = file.read()
file.close()
if (a == ""):
    a = 1
else:
    b = int(a)
    b += 1
    a = b
file = open("file.txt", "w")
file.write(str(a))
file.close()

Clientlist = [["HTL consulting legislation", "84 Willow Road\nWA3 UY2"],
              ["Warner bros inc.", "32 Rainbow Road\nRD4 U11"]]


# These are all functions for how the popup boxes work. Leave this pretty much as is.
def Namenewclient():
    global Namebox
    Namebox = Createnewclient("Enter client name:", 1)
    Namebox.button = Button(Namebox.popupbox, text="Continue", command=AddressNewclient)
    Namebox.button.pack()


def AddressNewclient():
    # Clientlist.append(Namebox.clientent.get())

    global Name
    Name = str(Namebox.clientent.get('1.0', '1.23'))
    Namebox.popupbox.destroy()
    global Addressbox
    Addressbox = Createnewclient("Enter client address", 2)
    Addressbox.button = Button(Addressbox.popupbox, text="Continue", command=Addressregister)
    Addressbox.button.pack()


def Addressregister():
    # Clientlist.append(Addressbox.clientent.get())
    global Address
    Address = str(Addressbox.clientent.get('1.0', '3.0'))
    AddnewClient()


def AddnewClient():
    Clientlist.append([Name, Address])
    client1d.append(Name)
    Addressbox.popupbox.destroy()

    # ClientMenu = OptionMenu(screen, dropdown, *client1d, "Add new Client+", command=dropdowncheck)

    for i in range(0, len(Clientlist)):
        client1d.append(Clientlist[i][0])
    ClientMenu.children["menu"].delete(0, len(Clientlist))

    for i in range(0, len(Clientlist)):
        nm = Clientlist[i][0]
        client1d.append(nm)
        ClientMenu.children["menu"].add_command(label=nm, command=_setit(dropdown, nm, optioncheck))

    ClientMenu.children["menu"].add_command(label="Add new Client+",
                                            command=_setit(dropdown, "Add new Client+", dropdowncheck))


def Createnewclient(txt, lines):
    class popup:
        popupbox = Toplevel()
        popupbox.title(txt)
        label = Label(popupbox, text=txt)
        clientent = Entry(popupbox)
        clientent = Text(popupbox, height=lines, width=30)
        label.pack(side="top", fill="x", pady=10)
        clientent.pack()
        popupbox.geometry("200x200")

    return (popup)


# function which gets address from the list depending on the name

def getName():
    for j in range(0, len(Clientlist)):
        if Clientlist[j][0] == dropdown.get():
            global displayName
            invoiceno.delete(0, END)
            sentence = str(Clientlist[j][0])[:9]
            invoiceno.insert(0, sentence.replace(" ", "") + "-" + str(a))



#           NameLbl['text'] = Clientlist[j][0]

def getAddress():
    for j in range(0, len(Clientlist)):
        if Clientlist[j][0] == dropdown.get():
            global displayAdress
            AddressLabel['text'] = str(Clientlist[j][1])


def optioncheck(self):
    if "Add new Client+" == dropdown.get():
        Namenewclient()
    else:
        getAddress()
        getName()


def dropdowncheck(self):
    if "Add new Client+" == dropdown.get():
        Namenewclient()


# Createnewclient("Text")

client1d = []
for i in range(0, len(Clientlist)):
    client1d.append(Clientlist[i][0])

dropdown = StringVar(screen)
displayAdress = "[Select Client]"
displayName = "[Select Client]"
dropdown.set("[Select Client]")
ClientMenu = OptionMenu(screen, dropdown, *client1d, "Add new Client+", command=optioncheck)
# Position of the dropdown
ClientMenu.grid(row=400, column=400)

AddressLabel = Label(screen, text=displayAdress)
# position where the Address is displayed
AddressLabel.grid(row=5, column=3)

# Invoice number widgets
invoicenolbl = Label(screen, text="INVOICE NO.")
invoiceno = Entry(screen, relief=SOLID)
invoiceno.insert(END, displayName[:9] + "-" + str(a))
invoicenolbl.grid(row=2, column=8)
invoiceno.grid(row=3, column=8)

#image for one point logo

imagefile="onepoint_logo.ico"
image = ImageTk.PhotoImage(Image.open(imagefile))

label = Label(image=image)
label.image = image
label.grid(row=1,column=1)

'''#image for one point logo#

imagefile="image.jpg"
image=  ImageTk.PhotoImage(Image.open(imagefile))
image.grid(row=1,colmn=1)

'''

screen.mainloop()