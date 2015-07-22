from Tkinter import *
import tkFileDialog
import tkMessageBox
import pdfquery
import re, os
from win32com.client.dynamic import Dispatch
import shutil

class App(Frame):
    def  __init__(self, master):
        Frame.__init__(self, master)
        self.grid()
        self.create_widgets()
        
        msg = Text(master)
        
        class infoMessage():
            def write(self, s):
                msg.insert(END, s)
                msg.yview_pickplace("end")
                msg.update_idletasks()
        
        Label(master, text="").grid(row=4)
        msg.config(bg="azure")
        msg.place(x=5, y=170, height=155, width=530)
        
        sys.stdout = infoMessage()
    
    def create_widgets(self):
        
        about = Label(master, text="\n\
                This app will merge the .pdf files with Table of Contents from the selected directory.\n\n")
        
        
        w = Label(master, text="          Folder Path:")
        var = StringVar()
        e = Entry(master, textvariable=var, width=52)
        b = Button(master, text="Browse", command=lambda:var.set(tkFileDialog.askdirectory()))
        
        about.grid(row=0, column=1, columnspan=5)
        w.grid(row=1, column=1)
        e.grid(row=1, column=2, columnspan=3)
        b.grid(row=1, column=5)
        
        Label(master, text="").grid(row=2)
        
        def parsePDFs():
            path = var.get()
            if path.find("htdocs") == -1:
                print ("error")
                tkMessageBox.showwarning("Wrong Path", "The path provided is not good. Please try again.")
                return
            else:
                if tkMessageBox.askyesno("Working Folder", "The selected folder for processing files is: \n\n" + path + "\n\nIs this correct?"):
                    print ("Parsing files in the working folder... Please wait until you see the message TASK COMPLETED.\n")
                    print (path)



                    file_list = []
                    for dirpath, dirnames, files in os.walk(path):
                        for f in files:
                            if f.lower().endswith("0001.pdf") and (f.lower().find('_fwd_') == -1) and (f.lower().find('_ek_') == -1):
                                fullpath = os.path.join(dirpath, f)
                                file_list.append(fullpath)
                                
                    for file in file_list:
                        pdf = pdfquery.PDFQuery(file)
                        pdf.load(0)
                        
                        label = pdf.pq('LTTextLineHorizontal:contains("CONTENTS")')
                        
                        left_corner = float(label.attr('x0'))
                        bottom_corner = float(label.attr('y0'))
                        
                        deep = 10
                        firstLineContents = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % (left_corner - 210, bottom_corner - deep, left_corner + 100, bottom_corner - 37)).text()
                        
                        _digits = re.compile('\d')
                        def contains_digits(d):
                            return bool(_digits.search(d))
                    
                        while contains_digits(firstLineContents[-2:]) == False:
                            deep = deep + 10
                            firstLineContents = pdf.pq('LTTextLineHorizontal:in_bbox("%s, %s, %s, %s")' % (left_corner - 210, bottom_corner - deep, left_corner + 100, bottom_corner - 37)).text()
                            print ("First Line is: %s" % firstLineContents)
                    
                        myValue = int(''.join(x for x in firstLineContents[-2:] if x.isdigit()))
                        print ("First TOC Number: %s." % myValue)




                        if myValue > 2:
                            print ('\nMerging file:')
                            print (str(os.path.basename(file)) + ' ==> starts from page ' + str(myValue))
                            
                            for i in range(1, myValue):
                                index = str(i)
                                # converts file path to string and deletes the last 5 characters from string
                                if i < 10:
                                    partfile = ''.join(str(file)[:-5])
                                else:
                                    partfile = ''.join(str(file)[:-6])
                                f = ''.join(partfile + '%s.pdf' % index)
                                print (f)
                                src = globals()['src' + str(i)] = os.path.abspath(f)
                                globals()['avdoc' + str(i)] = Dispatch("AcroExch.AVDoc")  # create an object for the document seen in the user interface
                                app = globals()['app' + str(i)] = Dispatch("Acroexch.App")  # instantiante Acrobat Application
                                app.Hide()  # Hide Acrobat application
                                aform = globals()['aform' + str(i)] = Dispatch("AFormAut.App")  # instantiate the main exposed object
                                
                                globals()['avdoc' + str(i)].Open(src, src)  # This is the view of the .pdf object in a window. There is one AVDoc object per displayed document
                                pddoc = globals()['pddoc' + str(i)] = globals()['avdoc' + str(i)].GetPDDoc()  # get PDDoc, the underlying representation of the document
                                globals()['N' + str(i)] = pddoc.GetNumPages()  # gets the number of the pages in the file
                            
                            i = myValue - 1
                            while (i > 0):
                                pddoc.InsertPages(N2 - 1, globals()['pddoc' + str(i)], 0, N1, 0)
                                i -= 1
                            
                            pddoc.DeletePages(0, 0)
                            
                            part_file = ''.join(str(os.path.basename(f))[:-6])
                                
                            temp = 'C:/temporary/'  # set a temporary folder to work with files
                            destination = "C:/dest/"  # set the destination directory
                            
                            if not os.path.isfile(temp) and not os.path.isdir(temp):
                                os.mkdir(temp)
                            
                            dest = str(os.path.dirname(f))
                            for i in range(1, myValue):
                                if i < 10:
                                    newfile = ''.join(part_file + '0%s.pdf' % i)
                                else:
                                    newfile = ''.join(part_file + '%s.pdf' % i)
                                
                                print ("New file created: %s" % newfile)
                                print ("Setting open to page.")
                                
                                
                                jscript = 'this.addScript("init", "this.pageNum = %s;");' % (i - 1)  # this is the javascript code, assigned to a variable
                                aform.Fields.ExecuteThisJavaScript(jscript)  # execute the jscript code on the .pdf
                                
                                pddoc.Save(1, os.path.abspath(temp + newfile))  # saves the document
                            
                            for i in range(1, myValue):
                                globals()['avdoc' + str(i)].Close(-1)  # close avdoc without saving
                            
                            # if the temp folder doesn't exist, create it
                            if not os.path.exists(temp):
                                os.mkdir(temp)
                                
                            # if the destination folder doesn't exist, create it
                            if not os.path.exists(destination):
                                os.mkdir(destination)
                                
                            for file in os.listdir(temp):
                                if file.endswith(".pdf"):
                                    src_file = os.path.join(temp, file)
                                    # declares destination file (path, file) | destination is for test. for normal app change it to dest
                                    dst_file = os.path.join(dest, file)
                                    
                                    shutil.move(src_file, dst_file)
                            
                            
                            
                            
                             
                        else:
                            print ('\nIgnoring file:')
                            print (str(os.path.basename(file)) + ' ==> starts from page ' + str(myValue))
                    
                    print ('\nTASK COMPLETED')
                    tkMessageBox.showwarning("TASK COMPLETED", "All the .pdf files containing the table of contents were merged. \nYou can close the application.")
              
        
        start = Button(master, text="         Start         ", command=parsePDFs)
        start.grid(row=3, column=3, columnspan=1)

master = Tk()
master.title("TOC Merger")
master.geometry("540x330")
master.resizable(width=FALSE, height=FALSE)
    
app = App(master)
master.mainloop()
