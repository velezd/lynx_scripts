import tkinter
import traceback
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from trap_activity import TrapActivity

class GUI():
    def __init__(self):
        """ Init GUI """
        root = tkinter.Tk()
        root.title('Trap activity')

        frm = ttk.Frame(root, padding=3)
        frm.grid(sticky=(N, S, E, W))
        ttk.Label(frm, text="Files").grid(column=0, row=0, sticky=(W))
        self.tk_filelist = ttk.Treeview(frm, show=['tree'])
        self.tk_filelist.grid(column=0, row=1, columnspan=2, sticky=(N, S, E, W))
        ttk.Button(frm, text='Add', command=self.add_files).grid(column=0, row=2, sticky=(E, W))
        ttk.Button(frm, text='Del', command=self.del_file).grid(column=1, row=2, sticky=(E, W))
        ttk.Button(frm, text="Process", command=self.process).grid(column=0, row=3, columnspan=2, sticky=(E, W))
        
        for child in frm.winfo_children(): 
            child.grid_configure(padx=3, pady=3)
        
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        frm.columnconfigure(0, weight=1)
        frm.columnconfigure(1, weight=1)
        frm.rowconfigure(1, weight=1)
        root.mainloop()
        
    def add_files(self):
        """ Add files to filelist """
        for filename in filedialog.askopenfilenames(filetypes=[('xlsx', '*.xlsx')]):
            self.tk_filelist.insert('','end', text=filename)
        
    def del_file(self):
        """ Delete selected files from filelist """
        for item in self.tk_filelist.selection():
            self.tk_filelist.delete(item)

    def process(self):
        """ Process file in filelist """
        filelist = [ self.tk_filelist.item(item, option='text') for item in self.tk_filelist.get_children() ]

        if len(filelist) == 0:
            tkinter.messagebox.showwarning(title='Warning', message='You have to select at least one file.')
            return

        with filedialog.asksaveasfile('w', title='Where to save result?', defaultextension='.csv', filetypes=[('csv', '*.csv')]) as result_file:
            try:
                ta = TrapActivity()
                
                for filename in filelist:
                    ta.load(filename)
                    
                ta.process_data()
                
                result_file.write(ta.format_data())
            except Exception as e:
                tkinter.messagebox.showerror(title='Python exception', message=traceback.format_exc())
            else:
                tkinter.messagebox.showinfo(title='Done', message=f'Saved in {result_file.name}')


if __name__ == '__main__':
    GUI()