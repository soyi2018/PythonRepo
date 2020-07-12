
# coding: utf-8

# In[ ]:


import tkinter as tk 
from tkinter import messagebox
from tkinter import filedialog
from tkinter import ttk 
from tkinter import font
import os
import shutil
from datetime import datetime
import pandas as pd
import glob
import random
import xlwings as xw
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import seaborn
import PIL
from PIL import Image, ImageTk

class PM_Data():
    def __init__(self, name, date1, date2, rtype, otype, thresh1, plist):
        self.name = name
        self.date1 = date1
        self.date2 = date2
        self.rtype = rtype
        self.otype = otype
        self.thresh1 = thresh1
        self.plist = plist
    
    def check_vals(self):
        try:
            date_obj1 = datetime.strptime(self.date1, '%Y-%m-%d')
            date_obj2 = datetime.strptime(self.date2, '%Y-%m-%d')
            float(self.thresh1)
            #float(thresh2)
            #float(divider)
        except:
            win = tk.Tk()
            win.withdraw()
            messagebox.showerror('Error', 'Please check your input values!')
            win.destroy()
            return False
        else:
            return True
                   
    def get_pm_data(self):
        n = 0
        pm_dir = r'C:\Users\syi\Desktop\PM Data Files'
        cols = [0,3,6,10,11,14,15,53,76]
        date_obj1 = datetime.strptime(self.date1, '%Y-%m-%d').date()
        date_obj2 = datetime.strptime(self.date2, '%Y-%m-%d').date()

        for fname in glob.glob(pm_dir+r'\*.xls*'):
            pm_data = pd.read_excel(fname, sheet_name=0, header=2, usecols=cols, names=['Status','ProjID','Partner','OrderDate','Deadline','LT','Purpose','DeliveryDate','Payment'])
            pm_data.dropna(how='any',subset=['Status','ProjID','Partner','OrderDate','Deadline','LT','Purpose','DeliveryDate'],inplace=True)
            pm_data = pm_data[pm_data['Status'].str.contains('finished',case=False)]
            pm_data.fillna(value={'Payment':0},inplace=True)
            
            pm_data = pm_data[(pm_data['OrderDate'] >= date_obj1) & (pm_data['OrderDate'] <= date_obj2)]
            if len(self.plist) == 1:
                pm_data = pm_data[pm_data['Partner'] == self.name]
                if len(pm_data['Partner']) < float(self.thresh1):
                    return None
            else:
                pm_data['OrderNum'] = pm_data.groupby('Partner')['Partner'].transform('count')
                pm_data = pm_data[pm_data['OrderNum'] >= float(self.thresh1)]
                pm_data.drop(['OrderNum'], axis=1, inplace=True)
                
                           
            pm_data['Delivery_Days'] = pm_data['DeliveryDate']-pm_data['OrderDate']
            pm_data['Delayed_Days'] = pm_data['DeliveryDate']-pm_data['Deadline']
            pm_data['Delivery_Weeks'] = pm_data['Delivery_Days'].astype('timedelta64[W]')
            pm_data['Delayed_Weeks'] = pm_data['Delayed_Days'].astype('timedelta64[W]')
            pm_data['Delivery_Days'] = pm_data['Delivery_Days'].astype('timedelta64[D]')
            pm_data['Delayed_Days'] = pm_data['Delayed_Days'].astype('timedelta64[D]')

            pm_data.loc[pm_data['Delayed_Days'] < 4, 'Delayed_Weeks'] = 0
            pm_data.loc[(pm_data['Delayed_Days'] >= 4) & (pm_data['Delayed_Days'] < 11), 'Delayed_Weeks'] = 1
            pm_data.loc[(pm_data['Delayed_Days'] >= 11) & (pm_data['Delayed_Days'] < 18), 'Delayed_Weeks'] = 2
            pm_data.loc[(pm_data['Delayed_Days'] >= 18) & (pm_data['Delayed_Days'] < 25), 'Delayed_Weeks'] = 3
            pm_data.loc[(pm_data['Delayed_Days'] >= 25) & (pm_data['Delayed_Days'] < 32), 'Delayed_Weeks'] = 4
            pm_data.loc[(pm_data['Delayed_Days'] >= 32) & (pm_data['Delayed_Days'] < 39), 'Delayed_Weeks'] = 5
            pm_data.loc[(pm_data['Delayed_Days'] >= 39) & (pm_data['Delayed_Days'] < 46), 'Delayed_Weeks'] = 6
            pm_data.loc[(pm_data['Delayed_Days'] >= 46) & (pm_data['Delayed_Days'] < 53), 'Delayed_Weeks'] = 7
            pm_data.loc[pm_data['Delayed_Days'] >= 53, 'Delayed_Weeks'] = 8
    
            pm_data.loc[pm_data['Delivery_Days'] < 4, 'Delivery_Weeks'] = 0
            pm_data.loc[(pm_data['Delivery_Days'] >= 4) & (pm_data['Delivery_Days'] < 11), 'Delivery_Weeks'] = 1
            pm_data.loc[(pm_data['Delivery_Days'] >= 11) & (pm_data['Delivery_Days'] < 18), 'Delivery_Weeks'] = 2
            pm_data.loc[(pm_data['Delivery_Days'] >= 18) & (pm_data['Delivery_Days'] < 25), 'Delivery_Weeks'] = 3
            pm_data.loc[(pm_data['Delivery_Days'] >= 25) & (pm_data['Delivery_Days'] < 32), 'Delivery_Weeks'] = 4
            pm_data.loc[(pm_data['Delivery_Days'] >= 32) & (pm_data['Delivery_Days'] < 39), 'Delivery_Weeks'] = 5
            pm_data.loc[(pm_data['Delivery_Days'] >= 39) & (pm_data['Delivery_Days'] < 46), 'Delivery_Weeks'] = 6
            pm_data.loc[(pm_data['Delivery_Days'] >= 46) & (pm_data['Delivery_Days'] < 53), 'Delivery_Weeks'] = 7
            pm_data.loc[pm_data['Delivery_Days'] >= 53, 'Delivery_Weeks'] = 8
            
            if n == 0:
                pm_data1 = pm_data
            else:
                pm_data1 = pm_data1.append(pm_data,ignore_index=True)
                
            n = n + 1
        pm_data1 = pm_data1.drop_duplicates(subset=['ProjID'], keep='last')
        if self.otype == 'Stock Orders':
            pm_data1 = pm_data1[pm_data1['Purpose'] == 'Stock']
        elif self.otype == 'Back Orders':
            pm_data1 = pm_data1[pm_data1['Purpose'] == 'BO']
        elif self.otype == 'Custom Synthesis':
            pm_data1 = pm_data1[(pm_data1['Purpose'] == 'BO') & (pm_data1['LT'] > 2)]
            pm_data1 = pm_data1[~((pm_data1['Purpose'] == 'BO') & (pm_data1['LT'] <= 3) & (pm_data1['Partner'] == 'FDC'))]
            
        self.pm_data = pm_data1    
        return pm_data1
    
    def get_plot(self, partner):
        #date_obj1 = datetime.strptime(self.date1, '%Y-%m-%d').date()
        #date_obj2 = datetime.strptime(self.date2, '%Y-%m-%d').date()
        #date_str1 = datetime.strftime(date_obj1, '%m/%d/%Y')
        #date_str2 = datetime.strftime(date_obj2, '%m/%d/%Y')
        #if len(self.plist) ==1:
            pm_data = self.pm_data[self.pm_data.Partner == partner]
            if self.rtype == 'Delay-Rate Report':
                ncount = len(pm_data)
                plt.figure(figsize=(6,4))
                ax = seaborn.countplot(x='Delayed_Weeks', data=pm_data)
                plt.title('%s Delay-Rate Plot of %s\n Total %d Orders' %(partner,self.otype,ncount))
                plt.xlabel('Number of Delayed Weeks')
                # Make twin axis
                ax2 = ax.twinx()
                # Switch so count axis is on right, frequency on left
                ax2.yaxis.tick_left()
                ax.yaxis.tick_right()
                # Also switch the labels over
                ax.yaxis.set_label_position('right')
                ax2.yaxis.set_label_position('left')
                ax2.set_ylabel('Percentage [%]')

                for p in ax.patches:
                    x=p.get_bbox().get_points()[:,0]
                    y=p.get_bbox().get_points()[1,1]
                    ax.annotate('{:.1f}%'.format(100.*y/ncount), (x.mean(), y), ha='center', va='bottom') # set the alignment of the text

                # Use a LinearLocator to ensure the correct number of ticks
                ax.yaxis.set_major_locator(ticker.LinearLocator(11))
                # Fix the frequency range to 0-100
                ax2.set_ylim(0,100)
                ax.set_ylim(0,ncount)
                # And use a MultipleLocator to ensure a tick spacing of 10
                ax2.yaxis.set_major_locator(ticker.MultipleLocator(10))

                # Need to turn the grid on ax2 off, otherwise the gridlines end up on top of the bars
                #ax2.grid(None)

                plt.savefig('./tempt/{}_delay_rate_plot.png'.format(partner))
                
            if self.rtype == 'Delivery-Rate Report':
                ncount = len(pm_data)
                plt.figure(figsize=(6,4))
                ax = seaborn.countplot(x='Delivery_Weeks', data=pm_data)
                plt.title('%s Delivery-Rate Plot of %s\n Total %d Orders' %(partner,self.otype,ncount))
                plt.xlabel('Number of Delivery Weeks')
                # Make twin axis
                ax2 = ax.twinx()
                # Switch so count axis is on right, frequency on left
                ax2.yaxis.tick_left()
                ax.yaxis.tick_right()
                # Also switch the labels over
                ax.yaxis.set_label_position('right')
                ax2.yaxis.set_label_position('left')
                ax2.set_ylabel('Percentage [%]')

                for p in ax.patches:
                    x=p.get_bbox().get_points()[:,0]
                    y=p.get_bbox().get_points()[1,1]
                    ax.annotate('{:.1f}%'.format(100.*y/ncount), (x.mean(), y), ha='center', va='bottom') # set the alignment of the text

                # Use a LinearLocator to ensure the correct number of ticks
                ax.yaxis.set_major_locator(ticker.LinearLocator(11))
                # Fix the frequency range to 0-100
                ax2.set_ylim(0,100)
                ax.set_ylim(0,ncount)
                # And use a MultipleLocator to ensure a tick spacing of 10
                ax2.yaxis.set_major_locator(ticker.MultipleLocator(10))

                # Need to turn the grid on ax2 off, otherwise the gridlines end up on top of the bars
                #ax2.grid(None)

                plt.savefig('./tempt/{}_delivery_rate_plot.png'.format(partner))
                
    def get_all_plots(self):
        if len(self.plist) > 1:
            plt.figure(figsize=(16,10))
            axx=seaborn.countplot('Partner', data=self.pm_data, order=self.pm_data['Partner'].value_counts().index)
            plt.title('Total # of Orders from Different Partners', fontsize = 20)
            plt.xlabel('Partner Name', fontsize = 15)
            plt.ylabel('Number of Orders', fontsize = 15)
            for p in axx.patches:
                x=p.get_bbox().get_points()[:,0]
                y=p.get_bbox().get_points()[1,1]
                axx.annotate('{:.0f}'.format(y), (x.mean(), y), ha='center', va='bottom') # set the alignment of the text
            
            plt.savefig('./tempt/Partner_Orders_{}.png'.format(self.otype))
            
            if self.rtype == 'Delay-Rate Report':
                plt.figure(figsize=(16,10))
                plt.title('Average Delay-Rate of Different Partners', fontsize=20)
                seaborn.barplot(x='Partner',y='Delayed_Weeks',data=self.pm_data,capsize=0.1)
                plt.xlabel('Partner Name',fontsize=15)
                plt.ylabel('Average # of Delayed Weeks',fontsize=15)
               
                plt.savefig('./tempt/Average_Delay_Plot_{}.png'.format(self.otype))
                
            if self.rtype == 'Delivery-Rate Report':
                plt.figure(figsize=(16,10))
                plt.title('Average Delivery-Rate of Different Partners', fontsize=20)
                seaborn.barplot(x='Partner',y='Delivery_Weeks', data=self.pm_data, capsize=0.1)
                plt.xlabel('Partner Name',fontsize=15)
                plt.ylabel('Average # of Delivery Weeks', fontsize=15)
            
                plt.savefig('./tempt/Average_Delivery_Plot_{}.png'.format(self.otype))
            
        for partner in self.pm_data['Partner'].unique().tolist():
            self.get_plot(partner)
            
    def get_dir(self):
        dirname = filedialog.askdirectory(title='Please select a directory:', initialdir=os.getcwd())
        self.dirname = dirname
        return dirname
                
    
    def save_to_excel(self):
        app = xw.App(visible=False)
        wb = app.books.add()
        wb.sheets[0].range('A1').options(pd.DataFrame, index=False).value = self.pm_data
        wb.save('pm_data_test_file.xlsx')
        wb.close()
        app.quit()
        
    def clean_tempt(self):
        path = './tempt'
        if not os.path.exists(path):
            os.mkdir(path)
        for fname in os.listdir(path):
            fname = os.path.join(path, fname)
            try:
                if os.path.isfile(fname):
                    os.unlink(fname)
                    #os.remove(fname)
                #elif os.path.isdir(fname):
                    #shutil.rmtree(fname)
            except Exception as e:
                print(e)
                
            

class MyCatalog():

    def __init__(self):
        self.path = r'C:\Users\syi\Desktop\Backup Files\AST-CAT_All\Catalog Management\Update Files\Catalog Release Sources.xlsm'
        
    def get_wkcatalog(self):
        try:
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(self.path)
            wkcatalog = wb.macro('CatalogExcel')
            wkcatalog()
            wb.save()
        except:
            print('Error: Please check the data file!')
            win = tk.Tk()
            win.withdraw()
            messagebox.showerror('Error', 'Please check your data file!')
            win.destroy()
            #win.mainloop()
        else:
            print('The catalog file has been generated successfully!')
        finally:
            wb.close()
            app.quit()
            
    def get_emolecules(self):
        try:
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(self.path)
            wkcatalog = wb.macro('Catalog_eMolecules')
            wkcatalog()
            wb.save()
        except:
            print('Error: Please check the data file!')
        else:
            print('The catalog file has been generated successfully!')
        finally:
            wb.close()
            app.quit()
            
    def get_namiki(self):
        try:
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(self.path)
            wkcatalog = wb.macro('NamikiStock')
            wkcatalog()
            wb.save()
        except:
            print('Error: Please check the data file!')
        else:
            print('The catalog file has been generated successfully!')
        finally:
            wb.close()
            app.quit()
            
    def get_sciquest(self):
        try:
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(self.path)
            mthcatalog = wb.macro('SciQuest_File')
            mthcatalog()
            wb.save()
        except:
            print('Error: Please check the data file!')
        else:
            print('The catalog file has been generated successfully!')
        finally:
            wb.close()
            app.quit()
            
    def get_fisher(self):
        try:
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(self.path)
            mthcatalog = wb.macro('FisherQuater')
            mthcatalog()
            wb.save()
        except:
            print('Error: Please check the data file!')
        else:
            print('The catalog file has been generated successfully!')
        finally:
            wb.close()
            app.quit()
            
    def get_labnetwork(self):
        try:
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(self.path)
            mthcatalog = wb.macro('LabNetwork_File')
            mthcatalog()
            wb.save()
        except:
            print('Error: Please check the data file!')
        else:
            print('The catalog file has been generated successfully!')
        finally:
            wb.close()
            app.quit()
            
    def get_labnetwork_sdf(self):
        try:
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(self.path)
            mthcatalog = wb.macro('LabNetwork_SDF')
            mthcatalog()
            wb.save()
        except:
            print('Error: Please check the data file!')
        else:
            print('The catalog file has been generated successfully!')
        finally:
            wb.close()
            app.quit()
            
    def get_pfizer_sdf(self):
        try:
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(self.path)
            mthcatalog = wb.macro('Pfizer_FlatPrice')
            mthcatalog()
            wb.save()
        except:
            print('Error: Please check the data file!')
        else:
            print('The catalog file has been generated successfully!')
        finally:
            wb.close()
            app.quit()
            
    def get_ariba(self):
        try:
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(self.path)
            qutcatalog = wb.macro('Ariba_File')
            qutcatalog()
            wb.save()
        except:
            print('Error: Please check the data file!')
        else:
            print('The catalog file has been generated successfully!')
        finally:
            wb.close()
            app.quit()
            
    def get_neta(self):
        try:
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(self.path)
            qutcatalog = wb.macro('GSK_Neta')
            qutcatalog()
            wb.save()
        except:
            print('Error: Please check the data file!')
        else:
            print('The catalog file has been generated successfully!')
        finally:
            wb.close()
            app.quit()
            
    def get_acd(self):
        try:
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(self.path)
            qutcatalog = wb.macro('ACD_File')
            qutcatalog()
            wb.save()
        except:
            print('Error: Please check the data file!')
        else:
            print('The catalog file has been generated successfully!')
        finally:
            wb.close()
            app.quit()
            
    def get_vwr_eu(self):
        try:
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(self.path)
            qutcatalog = wb.macro('VWR_Europe')
            qutcatalog()
            wb.save()
        except:
            print('Error: Please check the data file!')
        else:
            print('The catalog file has been generated successfully!')
        finally:
            wb.close()
            app.quit()
            
    def get_vwr_us(self):
        try:
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(self.path)
            qutcatalog = wb.macro('VWRWebItems_New')
            qutcatalog()
            wb.save()
        except:
            print('Error: Please check the data file!')
        else:
            print('The catalog file has been generated successfully!')
        finally:
            wb.close()
            app.quit()
            
    def get_namiki_bulk(self):
        try:
            app = xw.App(visible=False, add_book=False)
            wb = app.books.open(self.path)
            qutcatalog = wb.macro('BulkCatalog')
            qutcatalog()
            wb.save()
        except:
            print('Error: Please check the data file!')
        else:
            print('The catalog file has been generated successfully!')
        finally:
            wb.close()
            app.quit()

class ShowPlot(tk.Toplevel):
    def __init__(self, rtype, otype, name):
        #Frame.__init__(self, master)
        super().__init__()
        #board = tk.Frame(self)
        #self.board = board
        self.rtype = rtype
        self.otype = otype
        self.name = name
        self.title('Display Plot')
        self.iconbitmap('.\\AT logo.ico')
        self.columnconfigure(0,weight=1)
        self.rowconfigure(0,weight=1)
        if self.name == 'All':
            if self.rtype == 'Delay-Rate Report':
                self.original = Image.open('./tempt/Average_Delay_Plot_{}.png'.format(self.otype))
            if self.rtype == 'Delivery-Rate Report':
                self.original = Image.open('./tempt/Average_Delivery_Plot_{}.png'.format(otype))
        else:
            if self.rtype == 'Delay-Rate Report':
                self.original = Image.open('./tempt/{}_delay_rate_plot.png'.format(self.name))
            if self.rtype == 'Delivery-Rate Report':
                self.original = Image.open('./tempt/{}_delivery_rate_plot.png'.format(self.name))
        self.image = ImageTk.PhotoImage(self.original)
        self.display = tk.Canvas(self, bd=0, highlightthickness=0)
        self.display.create_image(0, 0, image=self.image, anchor=tk.NW, tags="IMG")
        self.display.grid(row=0, sticky=tk.W+tk.E+tk.N+tk.S)
        #board.pack(fill=tk.BOTH, expand=1)
        tk.Button(self, text='Save',
              fg='red', command=self.get_dir, width=10, 
              height=1, font=('Helvetica','10','bold')).grid(row=1, column=0)
        #tk.Button(self, text='Cancel',
        #      fg='red', command=self.get_dir, width=10, 
        #      height=1, font=('Helvetica','9','bold')).grid(row=1, column=1)
        self.bind("<Configure>", self.resize)

    def resize(self, event):
        size = (event.width, event.height)
        resized = self.original.resize(size,Image.ANTIALIAS)
        self.image = ImageTk.PhotoImage(resized)
        self.display.delete("IMG")
        self.display.create_image(0, 0, image=self.image, anchor=tk.NW, tags="IMG")
        
    def save_to_dir(self):
        path = './tempt'
        if self.name == 'All':
            for fname in os.listdir(path):
                fname = os.path.join(path, fname)
                shutil.copy(fname, self.dirname)
        else:
            if self.rtype == 'Delay-Rate Report':
                fname = '{}_delay_rate_plot.png'.format(self.name)
            if self.rtype == 'Delivery-Rate Report':
                fname = '{}_delivery_rate_plot.png'.format(self.name)
            fname = os.path.join(path, fname)
            shutil.copy(fname, self.dirname)
        
    def get_dir(self):
        dirname = filedialog.askdirectory(title='Please select a directory:', initialdir=os.getcwd())
        if dirname is not None and dirname != '':
            self.dirname = dirname
            self.save_to_dir()
            return dirname
        else:
            return None
            
class MyDialog1(tk.Toplevel):
    def __init__(self):
        super().__init__()
        self.title('Enter Parameters:')
        self.setup_GUI()

    def setup_GUI(self):
        n = 0
        pm_dir = r'C:\Users\syi\Desktop\PM Data Files'
        cols = [6,10]
        #Combine all files to get the list of partner names
        #for fname in glob.glob(pm_dir+r'\*.xls*'):
            #pm_data = pd.read_excel(fname, sheet_name=0, header=2, usecols=cols, names=['Partner','OrderDate'])
            #pm_data.dropna(how='any', inplace=True)
            #if n == 0:
                #pm_data1 = pm_data
            #else:
                #pm_data1 = pm_data1.append(pm_data, ignore_index=True)
            #n = n + 1
        for fname in glob.glob(pm_dir+r'\*.xls*'):
            if fname == 'C:\\Users\\syi\\Desktop\\PM Data Files\\Projects Closed Since 2010.xlsx':
                n = n - 1
            else:
                pm_data = pd.read_excel(fname, sheet_name=0, header=2, usecols=cols, names=['Partner','OrderDate'])
                pm_data.dropna(how='any', inplace=True)
                if n == 0:
                    pm_data1 = pm_data
                else:
                    pm_data1 = pm_data1.append(pm_data, ignore_index=True)
                n = n + 1
                
        pm_data1['Partner'] = list(map(lambda x: x.upper().strip(), pm_data1['Partner']))
        plist0 = pm_data1['Partner'].unique().tolist()
        #plist0 = sorted(plist0, key=str.lower)
        plist0 = sorted(plist0)
        self.plist1 = plist0
        #date1 = min(pm_data1['OrderDate']).date()
        date1 = '2009-08-10'
        date2 = max(pm_data1['OrderDate']).date()
        #Convert date to the format 'mm/dd/yyyy'
        #date1 = datetime.strftime(date1, '%m/%d/%Y')
        
        self.geometry('300x350+500+200')
        self.iconbitmap('.\\AT logo.ico')
        self.resizable(0,0)
        row1 = tk.Frame(self)
        row1.pack(fill='x')
        row2 = tk.Frame(self)
        row2.pack(fill='x')
        tk.Label(row1, text='Select a supplier name').pack(pady=5)
        self.name = tk.StringVar()
        plist = ttk.Combobox(row2, textvariable=self.name, width=19, state='readonly')
        #plist['values'] = ('All', 'ABCHEM', 'PHARMABLOCKS', 'LABNETWORK', 'INFOARK', 'SUNWAY', 'ANGENE')
        plist['values'] = ['All'] + plist0
        plist.current(0)
        plist.pack()

        row3 = tk.Frame(self)
        row3.pack(fill='x')
        row4 = tk.Frame(self)
        row4.pack(fill='x')
        tk.Label(row3, text='Start Date (yyyy-mm-dd)').pack(pady=5)
        self.date1 = tk.StringVar()
        self.date1.set(date1)
        tk.Entry(row4, textvariable=self.date1, width=22).pack()

        row5 = tk.Frame(self)
        row5.pack(fill='x')
        row6 = tk.Frame(self)
        row6.pack(fill='x')
        tk.Label(row5, text='End Date (yyyy-mm-dd)').pack(pady=5)
        self.date2 = tk.StringVar()
        self.date2.set(date2)
        tk.Entry(row6, textvariable=self.date2, width=22).pack()

        row7 = tk.Frame(self)
        row7.pack(fill='x')
        row8 = tk.Frame(self)
        row8.pack(fill='x')
        tk.Label(row7, text='Choose Report Type').pack(pady=5)
        spin = tk.Spinbox(row8, values=('Delay-Rate Report', 'Delivery-Rate Report'), width=22, bd=1)
        spin.config(state='readonly')
        self.rtype = spin
        spin.pack()
        
        row9 = tk.Frame(self)
        row9.pack(fill='x')
        row10 = tk.Frame(self)
        row10.pack(fill='x')
        tk.Label(row9, text='Choose Order Type').pack(pady=5)
        self.otype = tk.StringVar()
        olist = ttk.Combobox(row10, textvariable=self.otype, width=19, state='readonly')
        olist['values'] = ('All Orders', 'Stock Orders', 'Back Orders', 'Custom Synthesis')
        olist.current(0)
        olist.pack()
        
        row11 = tk.Frame(self)
        row11.pack(fill='x')
        row12 = tk.Frame(self)
        row12.pack(fill='x')
        tk.Label(row11, text='Thresh (Total # of Orders)').pack(pady=5)
        self.thresh1 = tk.StringVar(value=0)
        #self.thresh1.set(0)
        tk.Entry(row12, textvariable=self.thresh1, width=22).pack()

        row13 = tk.Frame(self)
        row13.pack(fill='x')
        tk.Button(row13, text='RUN', command=self.ok).pack(pady=5)
        #tk.Button(row13, text='Cancel', command=self.cancel).pack()

    def ok(self):
        self.inputval = [self.name.get(), self.date1.get(), self.date2.get(), self.rtype.get(), 
                         self.otype.get(), self.thresh1.get(), self.plist1]
        self.destroy()

    def cancel(self):
        self.inputval = None
        self.destroy()


class MyApp(tk.Tk):    
    def __init__(self):
        super().__init__()
        #self.pack()
        self.title('AstaTech Data Management v1.0')
        self.setup_GUI()

    def setup_GUI(self):
        self.geometry('450x450+30+30')
        self.iconbitmap('.\\AT logo.ico')
        
        #Set main menus
        menuBar = tk.Menu(self)

        fileMenu = tk.Menu(menuBar, tearoff=0)
        fileMenu.add_command(label='New...')
        fileMenu.add_command(label='Open...')
        fileMenu.add_command(label='Save')
        fileMenu.add_command(label='Save As...')
        fileMenu.add_command(label='Close')
        fileMenu.add_separator()
        fileMenu.add_command(label='Exit', command=self._quit)
        menuBar.add_cascade(label='File', menu=fileMenu)
        
        editMenu = tk.Menu(menuBar, tearoff=0)
        editMenu.add_command(label='Undo')
        editMenu.add_command(label='Cut')
        editMenu.add_command(label='Copy')
        editMenu.add_command(label='Paste')
        editMenu.add_command(label='Delete')
        editMenu.add_command(label='Select All')
        menuBar.add_cascade(label='Edit', menu=editMenu)
        
        helpMenu = tk.Menu(menuBar, tearoff=0)
        helpMenu.add_command(label='Help Index')
        helpMenu.add_command(label='About...')
        helpMenu.add_separator()
        helpMenu.add_command(label='Version')
        menuBar.add_cascade(label='Help', menu=helpMenu)
        
        self.config(menu=menuBar)
        
        
        #Set all Tabs
        tabNames = ['Data Reports', 'Data Analysis', 'Catalog Management', 'Others']
        self.tabnames = tabNames
        tabControl = ttk.Notebook(self)
        for tabnum in range(len(self.tabnames)):
            tabname = tabNames[tabnum]
            self.tabnames[tabnum] = ttk.Frame(tabControl)
            tabControl.add(self.tabnames[tabnum], text=tabname)
            tabControl.pack(expand=0, fill='both')
            
        #Tab1-Data Reports
        row11 = tk.Frame(self.tabnames[0])
        row11.pack(expand=1, fill='x')
        ct = [random.randrange(256) for x in range(3)]
        brightness = int(round(0.299*ct[0]+0.587*ct[1]+0.114*ct[2]))
        ct_hex = "%02x%02x%02x" % tuple(ct)
        bg_color = '#'+"".join(ct_hex)
        tk.Button(row11, text='Supplier Delay/Delivery Rate', bg=bg_color, 
              fg='white' if brightness<120 else 'black', command=self.pop_msg, width=25, 
              height=2, font=('Helvetica','10','bold')).grid(row=0,column=0,padx=5,pady=5)
        
        #ct = [random.randrange(256) for x in range(3)]
        #brightness = int(round(0.299*ct[0]+0.587*ct[1]+0.114*ct[2]))
        #ct_hex = "%02x%02x%02x" % tuple(ct)
        #bg_color = '#'+"".join(ct_hex)
        #tk.Button(row11, text='Save the file', bg=bg_color, 
         #     fg='white' if brightness<120 else 'black', command=self.select_dir, width=25, 
         #     height=2, font=('Helvetica','10','bold')).grid(row=0,column=1,padx=5,pady=5)
        
        #Tab3-Catalog Management
        catalog = MyCatalog()
        row31 = tk.Frame(self.tabnames[2])
        row31.pack(expand=1, fill='x')
        row32 = tk.Frame(self.tabnames[2])
        row32.pack()
        tk.Label(row31, text='Weekly Update Excel File', fg='red', font=('Helvetica','12','bold')).pack(pady=5)
        tk.Button(row32, text='General Customer',
              fg='red', command=catalog.get_wkcatalog, width=15, 
              height=1, font=('Helvetica','9','bold')).pack(padx=5, pady=10, side=tk.LEFT)
        tk.Button(row32, text='eMolecules File',
              fg='red', command=catalog.get_emolecules, width=15, 
              height=1, font=('Helvetica','9','bold')).pack(padx=5, pady=10, side=tk.LEFT)
        tk.Button(row32, text='Namiki File',
              fg='red', command=catalog.get_namiki, width=15, 
              height=1, font=('Helvetica','9','bold')).pack(padx=5, pady=10, side=tk.LEFT)
        row33 = tk.Frame(self.tabnames[2])
        row33.pack(expand=1, fill='x')
        row34 = tk.Frame(self.tabnames[2])
        row34.pack()
        row35 = tk.Frame(self.tabnames[2])
        row35.pack()
        tk.Label(row33, text='Monthly Update Excel File', fg='magenta', font=('Helvetica','12','bold')).pack()
        tk.Button(row34, text='SciQuest File',
              fg='magenta', command=catalog.get_sciquest, width=15, 
              height=1, font=('Helvetica','9','bold')).pack(padx=5, pady=10, side=tk.LEFT)
        tk.Button(row34, text='Fisher File',
              fg='magenta', command=catalog.get_fisher, width=15, 
              height=1, font=('Helvetica','9','bold')).pack(padx=5, pady=10, side=tk.LEFT)
        tk.Button(row35, text='LabNetwork File',
              fg='magenta', command=catalog.get_labnetwork, width=15, 
              height=1, font=('Helvetica','9','bold')).pack(padx=5, pady=10, side=tk.LEFT)
        tk.Button(row35, text='LabNetwork SDF',
              fg='magenta', command=catalog.get_labnetwork_sdf, width=15, 
              height=1, font=('Helvetica','9','bold')).pack(padx=5, pady=10, side=tk.LEFT)
        tk.Button(row35, text='Pfizer SDF',
              fg='magenta', command=catalog.get_pfizer_sdf, width=15, 
              height=1, font=('Helvetica','9','bold')).pack(padx=5, pady=10, side=tk.LEFT)
        row36 = tk.Frame(self.tabnames[2])
        row36.pack(expand=1, fill='x')
        row37 = tk.Frame(self.tabnames[2])
        row37.pack()
        row38 = tk.Frame(self.tabnames[2])
        row38.pack()
        tk.Label(row36, text='Quarterly Update Excel File', fg='purple', font=('Helvetica','12','bold')).pack()
        tk.Button(row37, text='Ariba File',
              fg='purple', command=catalog.get_ariba, width=15, 
              height=1, font=('Helvetica','9','bold')).pack(padx=5, pady=10, side=tk.LEFT)
        tk.Button(row37, text='Neta File',
              fg='purple', command=catalog.get_neta, width=15, 
              height=1, font=('Helvetica','9','bold')).pack(padx=5, pady=10, side=tk.LEFT)
        tk.Button(row37, text='Biovia-ACD File',
              fg='purple', command=catalog.get_acd, width=15, 
              height=1, font=('Helvetica','9','bold')).pack(padx=5, pady=10, side=tk.LEFT)
        tk.Button(row38, text='VWR-EU File',
              fg='purple', command=catalog.get_vwr_eu, width=15, 
              height=1, font=('Helvetica','9','bold')).pack(padx=5, pady=10, side=tk.LEFT)
        tk.Button(row38, text='VWR-US File',
              fg='purple', command=catalog.get_vwr_us, width=15, 
              height=1, font=('Helvetica','9','bold')).pack(padx=5, pady=10, side=tk.LEFT)
        tk.Button(row38, text='Namiki-Bulk File',
              fg='purple', command=catalog.get_namiki_bulk, width=15, 
              height=1, font=('Helvetica','9','bold')).pack(padx=5, pady=10, side=tk.LEFT)
        #ct = [random.randrange(256) for x in range(3)]
        #brightness = int(round(0.299*ct[0]+0.587*ct[1]+0.114*ct[2]))
        #ct_hex = "%02x%02x%02x" % tuple(ct)
        #bg_color = '#'+"".join(ct_hex)
        #tk.Button(row1, text='Supplier Delay-Rate Report', bg=bg_color, 
         #     fg='white' if brightness<120 else 'black', command=self.pop_msg, width=25, 
          #    height=2, font=('Helvetica','10','bold')).pack(pady=10)
    
    def do_events(self):
        pm_obj = PM_Data(self.name,self.date1,self.date2,self.rtype,self.otype,self.thresh1,self.plist1)
        pm_obj.clean_tempt()
        if pm_obj.check_vals():
            pm_data = pm_obj.get_pm_data()
            pm_obj.get_all_plots()
            
            plotmat = ShowPlot(self.rtype,self.otype,self.name)
                  

    def pop_msg(self):
        res = self.get_inputval1()
        if res == None:
            messagebox.showinfo('Warning', 'No value has been input!')
            return
        self.name = res[0]
        self.date1 = res[1]
        self.date2 = res[2]
        self.rtype = res[3]
        self.otype = res[4]
        self.thresh1 = res[5]
        self.plist1 = res[6]
        if self.name != 'All':
            self.plist1 = [self.name]
        #msg = 'Partner Name: ' + str(self.name) + ' StartDate: ' + str(self.date1) + ' EndDate: ' + str(self.date2) + ' Threshhold1: ' + str(self.thresh1) + ' Threshhold2: ' + str(self.thresh2) + ' Divider: ' + str(self.divider) + ' PartnerList: ' + str(self.plist1)
        #messagebox.showinfo('Important', msg)
        
        self.do_events()


    def get_inputval1(self):
        inputDialog1 = MyDialog1()
        self.wait_window(inputDialog1)   #important!
        return inputDialog1.inputval
    
    def _quit(self):
        self.quit()
        self.destroy()
        exit()
        


if __name__ == '__main__': 
    app = MyApp() 
    app.mainloop()

