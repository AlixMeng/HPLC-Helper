__author__ = 'Stephan'
import os, sys, xlsxwriter
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog, messagebox
from datetime import datetime
from collections import OrderedDict


rootdir = "."

#This section builds the display window
class Window(Frame):

    def __init__(self, parent):
        Frame.__init__(self, parent)
        self.simple=1
        self.parent = parent

        self.initUI()

        self.FolderLabel = Label(self, text="Folder:")
        self.FolderLabel.place(x=15, y=10)

        self.FolderDir = Label(self, text="", wraplength=215)
        self.FolderDir.place(x=55, y=10)

        SimpleRadio = Radiobutton (self, text="Extract Only", variable=self.simple, value=1, command=lambda: self.selon())
        SimpleRadio.place(x=75, y=70)

        SortedRadio = Radiobutton (self, text="Extract and Sort", variable=self.simple, value=0, command=lambda: self.seloff())
        SortedRadio.place(x=75, y=90)

        quitButton = Button(self, text="Quit", width=12,  command=self.quit)
        quitButton.place(x=5, y=120)

        GoButton = Button(self, text="Extract Data", width=12, command=self.go)
        GoButton.place(x=185, y=120)

        FolderButton = Button(self, text="Select Folder", width=12, command=lambda: self.setrootdir())
        FolderButton.place(x=95, y=120)

    def selon(self):
        self.simple=1

    def seloff(self):
        self.simple=0

    def setrootdir(self):
        global rootdir
        rootdir = filedialog.askdirectory()
        self.FolderDir['text'] = rootdir

    def go(self):
        if rootdir == ".":
            messagebox.showinfo(title="Select Directory", message='Please select a directory' + rootdir)
            return
        print(self.simple)
        if self.simple == 1:
            simplego(rootdir)
        elif self.simple == 0:
            go(rootdir)

    def initUI(self):

        self.parent.title("FlaskMan")
        self.style = Style()
        self.style.theme_use("default")

        self.pack(fill=BOTH, expand=1)

class DataPoint:
    def __init__(self):
        self.values = []

    def addname(self, id, name, path):

        self.id = id
        self.name = name
        self.path = path

        try: self.experiment = name.split('-')[1]
        except: self.experiment = "Exp"

        try: self.flask = int(name.split('-')[2])
        except ValueError: self.flask = name.split('-')[2]
        except IndexError: self.flask = str(id) + "-Media"

        try: self.sample = int(name.split('-')[3])
        except: self.sample = 0

    def addmethod(self, method):
        self.method = method

    def addvalue(self, value):
        self.values.append(value)

class Method:
    def __init__(self, name):
        self.name = name
        self.compounds = []

    def addcompound(self, compound):
        self.compounds.append(compound)

def simplesort(datapoints):
    experiments = {}

    for datafile in datapoints:
        if datapoints[datafile].method not in experiments:
            experiments[datapoints[datafile].method] = []
        experiments[datapoints[datafile].method].append(datapoints[datafile])

    for method in experiments:
        experiments[method].sort(key=lambda x: x.name)

    return experiments

def datasort(datapoints):

    expholder = {}
    methodholder = {}
    experiments = {}

    #sorts all files by experiment
    for datafile in datapoints:

        if datapoints[datafile].experiment not in expholder:
            expholder[datapoints[datafile].experiment] = []
        expholder[datapoints[datafile].experiment].append(datapoints[datafile])

    #within each experiment, sorts by method
    for experiment in expholder:
        for item in expholder[experiment]:

            if item.experiment not in methodholder:
                methodholder[item.experiment] = {}
            if item.method not in methodholder[experiment]:
                methodholder[experiment][item.method] = []
            methodholder[experiment][item.method].append(item)

    #within each experiment's method, sorts by flask
    for experiment in methodholder:
        for method in methodholder[experiment]:
            for item in methodholder[experiment][method]:

                if item.experiment not in experiments:
                    experiments[item.experiment] = {}
                if item.method not in experiments[experiment]:
                    experiments[experiment][item.method] = {}

                if any(x in str(item.flask).lower() for x in ["media", "yp"]):
                    if messagebox.askyesno("Media?", "Is " + item.name + " a media sample?"):
                        if "media" not in experiments[experiment][method]:
                            experiments[experiment][method]["media"] = []
                        experiments[experiment][method]["media"].append(item)
                        experiments[experiment][method]["media"].sort(key=lambda x: x.sample)
                        continue

                if item.flask not in experiments[experiment][method]:
                    experiments[experiment][method][item.flask] = []
                experiments[experiment][method][item.flask].append(item)
                experiments[experiment][method][item.flask].sort(key=lambda x: x.sample)

    return experiments

def go(rootdir):

    datapoints, methods = extract(rootdir)
    experiments = OrderedDict(datasort(datapoints))
    export(experiments, methods)
    messagebox.showinfo(title="Finished", message='Data extracted to:' + rootdir)
    rootdir = '.'
    del datapoints, methods, experiments

def simplego(rootdir):

    datapoints, methods = extract(rootdir)
    experiments = simplesort(datapoints)
    simpleexport(experiments, methods)
    messagebox.showinfo(title="Finished", message='Data extracted to:' + rootdir)
    rootdir = '.'
    del datapoints, methods, experiments

def simpleexport(experiments, methods):

    workbook = xlsxwriter.Workbook(os.path.join(rootdir, 'Data ' + datetime.now().strftime('%m-%d-%Y  %Ih%Mm%Ss') + ' .xlsx'))
    bold = workbook.add_format({'bold': True})

    for method in experiments:
        worksheetname = method[0:31]
        worksheet = workbook.add_worksheet(worksheetname)
        worksheet.set_column("A:A", 15)
        worksheet.set_column(("B:" + str(chr(len(methods[method].compounds)+65))), 12)
        worksheet.write(0,0, "Method:", bold)
        worksheet.write(0,1, method)

        worksheet.write(2,0,"Sample Name", bold)
        compoundcount=1

        for compound in methods[method].compounds:
            worksheet.write(2, compoundcount, compound, bold)
            compoundcount+=1

        linecount=3
        for datapoint in experiments[method]:

            worksheet.write(linecount, 0, datapoint.name)
            valuecount = 0

            for value in datapoint.values:
                valuecount+=1
                worksheet.write(linecount, valuecount, value)
            linecount+=1

    workbook.close()

def export(experiments, methods):

    for experiment in experiments:

        workbook = xlsxwriter.Workbook(os.path.join(rootdir, experiment + ' Data ' + datetime.now().strftime('%m-%d-%Y %Ih%Mm%Ss') + '.xlsx'))
        bold = workbook.add_format({'bold': True})

        for method in experiments[experiment]:
            singlemedia = 0
            try:
                if len(experiments[experiment][method]["media"]) == 1:
                    singlemedia = 1
            except:
                pass

            worksheetname = method[0:31]
            worksheet = workbook.add_worksheet(worksheetname)

            worksheet.set_column(("A:" + str(chr(len(methods[method].compounds)+65))), 12)
            worksheet.write(0, 0, "Experiment:", bold)
            worksheet.write(0, 1, str(experiment))
            worksheet.write(1, 0, "Method:", bold)
            worksheet.write(1, 1, str(method))

            flasks = experiments[experiment][method].keys()

            linecount = 2

            for item in flasks:
                linecount+=2
                ccount = 1
                samplecount=0

                worksheet.write(linecount, 0,"Flask: " + str(item), bold)
                for compound in methods[method].compounds:
                    ccount +=1
                    worksheet.write(linecount, ccount, str(compound), bold)

                worksheet.write(linecount, 1, "Sample", bold)

                if singlemedia == 1 and item != "media":
                    linecount+=1
                    worksheet.write(linecount, 0, experiments[experiment][method]["media"][0].name, bold)
                    worksheet.write(linecount, 1, "Media", bold)
                    mcount = 2
                    for mediavalue in experiments[experiment][method]["media"][0].values:
                        worksheet.write_number(linecount, mcount, mediavalue)
                        mcount+=1

                for sample in experiments[experiment][method][item]:
                    linecount +=1
                    samplecount+=1
                    worksheet.write(linecount, 0, str(sample.name), bold)
                    worksheet.write(linecount, 1, str(sample.sample), bold)
                    xcount = 1

                    for value in sample.values:
                        xcount+=1
                        worksheet.write_number(linecount, xcount,  float(value))


        workbook.close()


def extract(rootdir):

    datapoints = dict()
    methods = dict()


    filecounter = 0

    for root, subFolders, files in os.walk(rootdir):
        if 'Report.TXT' in files:
            filecounter += 1
            datafile = DataPoint()
            a = 0
            switch = 0
            methodswitch = 0
            methodloop = 0

            with open(os.path.join(root, 'Report.TXT'), 'r', encoding='utf-16') as vreport:
                report = vreport.read().splitlines()
                for line in report:
                    if str(line.split(' ', 1)[0]) == "Totals":
                        break

                    if methodloop == 1:
                        methodloop = 0
                        if "Last changed" not in line:
                            method = str(os.path.split(method + line.strip())[1]).split(' ')[0]
                        else:
                            method = str(os.path.split(method)[1]).split(' ')[0]

                        datafile.addmethod(method)

                        if method not in methods:
                            newmethod = Method(method)
                            methodswitch = 1

                    if line.split(' ', 1)[0] == "Sample":
                        samplename = str(line[13:43]).strip()
                        datafile.addname(filecounter, samplename, root)

                    if line.split(' ', 1)[0] in ["Method", "Analysis"]:
                        if "Method Info" not in line:
                            method = line[17:300].strip()
                            methodloop = 1

                    if "[ng/ul]" in line:
                        readloc = int(line.index('[ng/ul]') - 1)
                        switch = 1

                    if "[mg/L]" in line:
                        readloc = int(line.index('[mg/L]') - 1)
                        switch = 1

                    if "[g/l]" in line:
                        readloc = int(line.index('[g/l]') - 1)
                        switch = 1

                    if switch == 1: a+=1
                    if a >= 3:
                        if methodswitch == 1:
                            newmethod.addcompound(str(line[(readloc+10):(readloc+40)]).strip())

                        value = str(line[readloc:(readloc+10)]).strip()
                        if value == "-":
                            value = 0.0
                        else:
                            value = float(value)
                        datafile.addvalue(value)

            datapoints[samplename + "-" + str(filecounter)] = datafile
            if methodswitch == 1: methods[method] = newmethod

    return (datapoints, methods)

##Window definitions (specific size, not resizable, main loop)

def main():

    root = Tk()
    root.resizable(0,0)
    root.geometry("275x150+300+300")

    app = Window(root)
    root.mainloop()

if __name__ == '__main__':
    main()
