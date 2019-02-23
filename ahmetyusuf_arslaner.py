from Tkinter import *
from xlrd import *
from recommendations import *
import ttk
import tkFont
import os
import tkFileDialog
import anydbm
import pickle

cwd = os.getcwd()


class Cafedata:  # This class gets the cafe names from the excel file using xlrd and tkFileDialog module
    def __init__(self):  # It adds cafes to list and shows them on the combobox.
        self.cafelist = []  # Used try-except in open_excel function to address the IOError in case user does not
        # choose a database file.

    def open_excel(self):
        try:
            filename = tkFileDialog.askopenfilename(initialdir=cwd, title="Select file",
                                                    filetypes=(("excel files", "*.xlsx"), ("all files", "*.*")))
            book1 = open_workbook(filename)
            first_sheet = book1.sheet_by_index(0)
            cells = first_sheet.col_slice(colx=0, start_rowx=1, end_rowx=38)
            self.cafelist = []
            for cell in cells:
                self.cafelist.append(cell.value)
            Gui(root).cafecombobox['values'] = self.cafelist
        except IOError:
            print "ERROR!\nYou did not choose a file or you chose an invalid file"


class Ratingsdb:  # This class gets the ratings from the .db file and adds them to a dictionary using anydbm.
    def __init__(self):  # Used try-except in open_db function to address the IOError in case user does not \
        self.databasedict = {}  # choose a database file.

    def open_db(self):
        try:
            filename1 = tkFileDialog.askopenfilename(initialdir=cwd, title="Select file",
                                                     filetypes=(("database files", "*.db"), ("all files", "*.*")))
            ratingsdb = anydbm.open(filename1, 'r')
            for key, value in ratingsdb.iteritems():
                self.databasedict[key] = pickle.loads(value)
        except IOError:
            print "ERROR!\nYou did not choose a file or you chose an invalid file!"


class Gui(Frame):  # GUI Class that contains rest of the code
    def __init__(self, root):
        Frame.__init__(self, root)
        self.root = root
        self.initUI()

    def initUI(self):  # initUI
        self.dbclass = Ratingsdb()   # To access Ratingsdb easier
        self.frame1 = Frame(bg="lightgreen", borderwidth=2, relief=SOLID)  # First frame holds the title
        self.frame1.grid(row=0, column=0, sticky=EW, columnspan=5)

        self.frame2 = Frame(borderwidth=2, relief=SOLID)  # This frame holds the upload buttons
        self.frame2.grid(row=1, column=0, sticky=EW, columnspan=5)

        self.frame3 = Frame(borderwidth=2, relief=SOLID)  # This frame holds rating and recommend buttons in the left
        self.frame3.grid(row=2, column=0, sticky=W)

        self.frame4 = Frame(bg="lightgreen", borderwidth=2, relief=SOLID)  # This frame holds the left part of the
        self.frame4.grid(row=2, column=0, sticky=W, padx=70)  # changing UI when buttons are pressed

        self.frame5 = Frame(borderwidth=2, relief=SOLID)  # This frame holds the right part of the changing UI when
        self.frame5.grid(row=2, column=0, sticky=W, padx=430)  # buttons are pressed

        self.titlelabel = Label(self.frame1, text="CAFE RECOMMENDER", fg="red", bg="lightgreen", font=("Helvetica", 25),
                                width=45, height=2)
        self.titlelabel.grid(row=0, column=0, sticky=EW)

        self.uploadlabel = Label(self.frame2, bg="lightgreen", width=160, height=4)  # used some empty labels to make
        self.uploadlabel.grid(row=0, column=0, sticky=EW, columnspan=3)  # the GUI look better

        self.uploadcafebutton = Button(self.frame2, text="Upload Cafe Data", bg="red", fg="white", height=3, width=20,
                                       command=Cafedata().open_excel)  # Button that makes user to choose an excel file
        self.uploadcafebutton.grid(row=0, column=0)

        self.uploadratingsbutton = Button(self.frame2, text="Upload Ratings", bg="red", fg="white", height=3, width=20,
                                          command=self.dbclass.open_db)  # Button that makes user to choose a .db file
        self.uploadratingsbutton.grid(row=0, column=1)

        self.verticallabel = Label(self.frame3, bg="lightgreen", width=10, height=30)
        self.verticallabel.grid(row=0, column=0, rowspan=4)

        self.var1 = StringVar()
        self.var1 = "pink"

        self.countervar = IntVar()  # variable that will be used for keeping track of the GUI requests
        self.countervar = 0

        self.ratingbutton = Button(self.frame3, text="RATING", wraplength=1, bg=self.var1, fg="white", justify=CENTER,
                                   width=6, height=10, state=ACTIVE, command=self.ratingcommand, activebackground="pink"
                                   , relief=SUNKEN)  # Button that shows us the rating part of the GUI, comes as active
        self.ratingbutton.grid(row=0, column=0, pady=5, padx=10)
        self.ratingbutton.bind("<Enter>", self.on_enter)  # used events for the bonus part of the project
        self.ratingbutton.bind("<Leave>", self.on_leave)

        self.var2 = StringVar()     # variable that keeps track of the rating and recommend buttons
        self.var2 = "red"

        self.recommendbutton = Button(self.frame3, text="RECOMMEND", wraplength=1, bg=self.var2, fg="white",
                                      justify=CENTER, width=6, height=10, command=self.recommendcommand,
                                      activebackground="pink")  # this button shows the recommendation part of the GUI
        self.recommendbutton.grid(row=1, column=0, pady=50, padx=10)
        self.recommendbutton.bind("<Enter>", self.on_enter1)  # used events for the bonus part of the project
        self.recommendbutton.bind("<Leave>", self.on_leave1)

        self.verticallabel2 = Label(self.frame4, bg="lightgreen", width=52, height=30)
        self.verticallabel2.grid(row=0, column=1, rowspan=8, sticky=EW)

        self.combovar = StringVar()  # variable for the combobox that keeps track of what is selected

        self.cafecombobox = ttk.Combobox(self.frame4, textvariable=self.combovar, width=50)  # combobox widget
        self.cafecombobox['values'] = ''
        self.cafecombobox.grid(column=1, row=1, padx=20)

        self.choosecafelabel = Label(self.frame4, text="Choose Cafe", font=("Helvetica", 20), bg="lightgreen")
        self.choosecafelabel.grid(column=1, row=0, sticky=N, pady=20)

        self.chooseratinglabel = Label(self.frame4, text="Choose Rating", font=("Helvetica", 20), bg="lightgreen")
        self.chooseratinglabel.grid(column=1, row=2, sticky=S)

        self.ratingslider = Scale(self.frame4, from_=1, to=10, orient=HORIZONTAL, length=320, bg="lightgreen",
                                  troughcolor="white", highlightthickness=0, activebackground="lightgreen")
        self.ratingslider.grid(column=1, row=3)

        self.addbutton = Button(self.frame4, text="ADD", bg="red", fg="white", justify=CENTER, width=15, height=2,
                                command=self.addbuttonfunc)
        self.addbutton.grid(column=1, row=4)

        self.verticallabel3 = Label(self.frame5, bg="lightgreen", width=75, height=30)
        self.verticallabel3.grid(column=0, row=0, rowspan=8)

        self.treedict = {}  # Defined a dictionary to store all the items that are added to it
        self.tree = ttk.Treeview(self.frame5, height=17)  # tree widget
        self.tree["columns"] = ("cafe", "ratings")
        self.tree.column("cafe", width=200)
        self.tree.column("ratings", width=100)
        self.tree.heading("cafe", text="Cafe")
        self.tree.heading("ratings", text="Rating")
        self.tree["show"] = "headings"
        self.tree.grid(column=0, row=0, sticky=NW, pady=10, rowspan=8, padx=75)

        self.removebutton = Button(self.frame5, bg="red", fg="white", height=3, width=10, text="REMOVE",
                                   command=self.removebuttonfunc)
        self.removebutton.grid(column=0, row=3, sticky=E, padx=65)

        self.settingslabel = Label(self.frame4, bg="lightgreen", fg="black", text="Settings", font=("Helvetica", 20))
        underlinedfont = tkFont.Font(self.settingslabel, self.settingslabel.cget("font"))
        underlinedfont.configure(underline=True)  # using tkFont was te only way I could make Settings underlined
        self.settingslabel.configure(font=underlinedfont)

        self.numberofreclabel = Label(self.frame4, bg="lightgreen", fg="black", text="Number of Recommendation",
                                      font=("Helvetica", 12))

        self.numentry = Entry(self.frame4, width=10)

        self.simlabel = Label(self.frame4, bg="lightgreen", fg="black", text="Similarity Metrics",
                              font=("Helvetica", 12))

        self.euclideanvar = IntVar()
        self.euclideancheck = Checkbutton(self.frame4, text="Euclidian", variable=self.euclideanvar, bg="lightgreen",
                                          activebackground="lightgreen", onvalue=1, offvalue=0,
                                          command=self.euclideancheckf)

        self.pearsonvar = IntVar()
        self.pearsoncheck = Checkbutton(self.frame4, text="Pearson", variable=self.pearsonvar, bg="lightgreen",
                                        activebackground="lightgreen", onvalue=1, offvalue=0,
                                        command=self.pearsoncheckf)

        self.jaccardvar = IntVar()
        self.jaccardcheck = Checkbutton(self.frame4, text="Jaccard", variable=self.jaccardvar, bg="lightgreen",
                                        activebackground="lightgreen", onvalue=1, offvalue=0,
                                        command=self.jaccardcheckf)

        self.recbutstate = StringVar()
        self.recsimuserbutton = Button(self.frame4, text="Recommend" + "\n" + "Similar User", bg="red", fg="white",
                                       width=11, height=2, command=self.recsimuserfunc)

        self.reccafebutton = Button(self.frame4, text="Recommend" + "\n" + "Cafe", bg="red", fg="white", width=11,
                                    height=2, command=self.reccafefunc)

        self.simuserlabel = Label(self.frame5, text="Similar User", bg="lightgreen", fg="black", font=("Helvetica", 12))

        self.simusertree = ttk.Treeview(self.frame5, height=6)  # shows similar users
        self.simusertree["columns"] = ("user", "similarity")
        self.simusertree.column("user", width=200)
        self.simusertree.column("similarity", width=100)
        self.simusertree.heading("user", text="User")
        self.simusertree.heading("similarity", text="Similarity")
        self.simusertree["show"] = "headings"

        self.getusersratbutton = Button(self.frame5, text="Get User's Rating", width=15, command=self.getusersratfunc)

        self.userratinglabel = Label(self.frame5, text="Select user to see his/her rating", bg="lightgreen", fg="black")

        self.userrattree = ttk.Treeview(self.frame5, height=6)  # shows selected users ratings
        self.userrattree["columns"] = ("cafe", "ratings")
        self.userrattree.column("cafe", width=200)
        self.userrattree.column("ratings", width=100)
        self.userrattree.heading("cafe", text="Cafe")
        self.userrattree.heading("ratings", text="Rating")
        self.userrattree["show"] = "headings"

        self.similarcaflabel = Label(self.frame5, bg="lightgreen", fg="black", text="Similar Cafes",
                                     font=("Helvetica", 12))

        self.simcafetree = ttk.Treeview(self.frame5, height=16)  # shows similar cafes
        self.simcafetree["columns"] = ("cafe", "similarity")
        self.simcafetree.column("cafe", width=200)
        self.simcafetree.column("similarity", width=100)
        self.simcafetree.heading("cafe", text="Cafe")
        self.simcafetree.heading("similarity", text="Similarity")
        self.simcafetree["show"] = "headings"

    def ratingcommand(self):  # command of the rating button
        self.var1 = "pink"
        self.var2 = "red"
        self.ratingbutton.config(state=ACTIVE, relief=SUNKEN, bg=self.var1)
        self.recommendbutton.config(bg="red", state=NORMAL, relief=RAISED)

        self.cafecombobox.grid(column=1, row=1, padx=20)
        self.choosecafelabel.grid(column=1, row=0, sticky=N, pady=20)
        self.chooseratinglabel.grid(column=1, row=2, sticky=S)
        self.ratingslider.grid(column=1, row=3)
        self.addbutton.grid(column=1, row=4)
        self.tree.grid(column=0, row=0, sticky=N, pady=10, rowspan=8)
        self.removebutton.grid(column=1, row=0, sticky=W)

        self.settingslabel.grid_forget()
        self.numberofreclabel.grid_forget()
        self.numentry.grid_forget()
        self.simlabel.grid_forget()
        self.euclideancheck.grid_forget()
        self.pearsoncheck.grid_forget()
        self.jaccardcheck.grid_forget()
        self.recsimuserbutton.grid_forget()
        self.reccafebutton.grid_forget()
        self.simuserlabel.grid_forget()
        self.simusertree.grid_forget()
        self.getusersratbutton.grid_forget()
        self.userratinglabel.grid_forget()
        self.userrattree.grid_forget()
        self.similarcaflabel.grid_forget()
        self.simcafetree.grid_forget()

    def recommendcommand(self):  # command of the recommendation button
        self.var1 = "red"
        self.var2 = "pink"
        self.recommendbutton.config(state=ACTIVE, relief=SUNKEN, bg=self.var2)
        self.ratingbutton.config(bg="red", state=NORMAL, relief=RAISED)

        self.cafecombobox.grid_forget()
        self.choosecafelabel.grid_forget()
        self.chooseratinglabel.grid_forget()
        self.ratingslider.grid_forget()
        self.addbutton.grid_forget()
        self.tree.grid_forget()
        self.removebutton.grid_forget()

        self.settingslabel.grid(column=0, row=0, padx=80)
        self.numberofreclabel.grid(column=0, row=1, padx=20)
        self.numentry.grid(column=0, row=2, padx=80)
        self.simlabel.grid(column=0, row=3, padx=50)
        self.euclideancheck.grid(column=0, row=4, sticky=N)
        self.pearsoncheck.grid(column=0, row=4)
        self.jaccardcheck.grid(column=0, row=4, sticky=S)
        self.recsimuserbutton.grid(column=1, row=2, sticky=W)
        self.reccafebutton.grid(column=1, row=3, sticky=W)

        if self.countervar == 1:  # to prevent the Rating page resetting
            if self.recbutstate == "user":
                self.recsimuserinterface()
            elif self.recbutstate == "cafe":
                self.recsimcafeinterface()
            else:
                pass
        else:
            pass

        self.countervar = 1

    def on_enter(self, event):  # These four functions from this one makes the rating and recommend buttons turn orange
        self.ratingbutton.config(bg="orange")  # when user hovers the mouse over them

    def on_enter1(self, event):
        self.recommendbutton.config(bg="orange")

    def on_leave(self, event):
        if self.var1 == "red":
            self.ratingbutton.config(bg="red")
        else:
            self.ratingbutton.config(bg="pink")

    def on_leave1(self, event):
        if self.var2 == "red":
            self.recommendbutton.config(bg="red")
        else:
            self.recommendbutton.config(bg="pink")

    def addbuttonfunc(self):    # command of th eadd button
        if self.combovar.get() != '':
            if self.combovar.get() in self.treedict:
                print "Item is already added"
            else:
                self.treedict[self.combovar.get()] = self.ratingslider.get()
                self.tree.insert("", END, values=(self.combovar.get(), self.ratingslider.get()))
                print self.treedict
        else:
            print "ERROR!\nYou did not choose a cafe"

    def removebuttonfunc(self):  # command of the remove button
        try:
            selecteditem = self.tree.focus()
            del self.treedict[self.tree.item(selecteditem)['values'][0]]
            self.tree.delete(selecteditem)
        except IndexError:
            print "ERROR! \nYou did not add anything to the list or you did not choose anything from the list"

    def pearsoncheckf(self):
        self.euclideancheck.deselect()
        self.jaccardcheck.deselect()

    def euclideancheckf(self):
        self.pearsoncheck.deselect()
        self.jaccardcheck.deselect()

    def jaccardcheckf(self):
        self.pearsoncheck.deselect()
        self.euclideancheck.deselect()

    def recsimuserinterface(self):
        self.simuserlabel.grid(column=0, row=0, sticky=N, pady=10, padx=120)
        self.simusertree.grid(column=0, row=0, padx=100, sticky=N, pady=35)
        self.getusersratbutton.grid(column=0, row=0, sticky=N, padx=120, pady=190)
        self.userratinglabel.grid(column=0, row=0, padx=130, sticky=S, pady=170)
        self.userrattree.grid(column=0, row=0, padx=100, sticky=S, pady=22)

        self.similarcaflabel.grid_forget()
        self.simcafetree.grid_forget()

    def recsimcafeinterface(self):
        self.similarcaflabel.grid(column=0, row=0, sticky=N, pady=10, padx=120)
        self.simcafetree.grid(column=0, row=0, sticky=N, pady=40, rowspan=8)

        self.simuserlabel.grid_forget()
        self.simusertree.grid_forget()
        self.getusersratbutton.grid_forget()
        self.userratinglabel.grid_forget()
        self.userrattree.grid_forget()

    def recsimuserfunc(self):  # function that recommends similar user
        self.recbutstate = "user"
        self.simusertree.delete(*self.simusertree.get_children())
        self.userrattree.delete(*self.userrattree.get_children())
        self.userratinglabel.config(text="Select user to see his/her rating")
        if self.dbclass.databasedict == {}:
            print "ERROR!\nPlease upload a valid database file!"
        elif self.treedict == {}:
            print "ERROR!\nPlease rate cafes in order to get recommendations!"
        else:
            self.dbclass.databasedict['user'] = self.treedict
            try:
                self.entryval = int(self.numentry.get())
                if self.euclideanvar.get() == 1:
                    self.topdicteuclidean = topMatches(self.dbclass.databasedict, 'user', self.entryval,
                                                       similarity=sim_distance)
                    for value in self.topdicteuclidean:
                        self.simusertree.insert("", END, values=(value[1], value[0]))
                    self.recsimuserinterface()
                elif self.pearsonvar.get() == 1:
                    self.topdictpearson = topMatches(self.dbclass.databasedict, 'user', self.entryval,
                                                     similarity=sim_pearson)
                    for value1 in self.topdictpearson:
                        self.simusertree.insert("", END, values=(value1[1], value1[0]))
                    self.recsimuserinterface()
                elif self.jaccardvar.get() == 1:
                    self.topdictjaccard = topMatches(self.dbclass.databasedict, 'user', self.entryval,
                                                     similarity=sim_jaccard)
                    for value2 in self.topdictjaccard:
                        self.simusertree.insert("", END, values=(value2[1], value2[0]))
                    self.recsimuserinterface()
                else:
                    print "ERROR!\nYou did not choose any type of similarity metric!"
            except ValueError:
                print "ERROR!\nPlease enter a valid number in the box!"

    def reccafefunc(self):  # function that recommends cafes
        self.recbutstate = "cafe"
        self.simcafetree.delete(*self.simcafetree.get_children())
        if self.dbclass.databasedict == {}:
            print "ERROR!\nPlease upload a valid database file!"
        elif self.treedict == {}:
            print "ERROR!\nPlease rate cafes in order to get recommendations!"
        else:
            self.dbclass.databasedict['user'] = self.treedict
            try:
                self.entryval = int(self.numentry.get())
                itemsim = calculateSimilarItems(self.dbclass.databasedict, n=1)
                masterlist = getRecommendedItems(self.dbclass.databasedict, itemsim, 'user')
                print
                for value in masterlist:
                    self.simcafetree.insert("", END, values=(value[1], value[0]))
                self.recsimcafeinterface()
            except ValueError:
                print "ERROR!\nPlease enter a valid number in the box!"

    def getusersratfunc(self):
        self.userrattree.delete(*self.userrattree.get_children())
        try:
            selecteduser = self.simusertree.focus()
            getuserratsdict = self.dbclass.databasedict[self.simusertree.item(selecteduser)['values'][0]]
            self.userratinglabel.config(text=self.simusertree.item(selecteduser)['values'][0] + "'s Ratings")
            for key, value in getuserratsdict.iteritems():
                self.userrattree.insert("", END, values=(key, value))
        except IndexError:
            print "ERROR!\nYou did not choose a user from the list!"


if __name__ == '__main__':
    root = Tk()
    root.geometry('900x550+300+60')
    root.title("Cafe Recommender 1.0")
    app = Gui(root)
    root.update()
    root.mainloop()
