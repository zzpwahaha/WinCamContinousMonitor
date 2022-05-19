import wx
import wx.lib.activex
import csv
import comtypes.client
import pandas as pd
import numpy as np
import h5py

import datetime
import schedule
import time

class EventSink(object):

    def __init__(self, frame):
        self.counter = 0
        self.frame = frame

    def DataReady(self):
        self.counter +=1
        self.frame.Title= "DataReady fired {0} times".format(self.counter)


class MyApp( wx.App ): 
    
    def OnClick(self,e):
        pathname = self.ti.Value
        # pathname = r"Z:\common\LabData\NewRb\Users\Zhenpu\cryoParallelismTest\doubleLayer_cycleDown1\"
        
        def job(counter):
            print("I'm working... at {} with job {}".format(datetime.datetime.now(),counter['count']))
            data = self.gd.ctrl.GetWinCamDataAsVariant()
            
            f = h5py.File(pathname + 'data{}'.format(counter['count'])+ '.hdf5', "w")
            pixelxsize = self.gd.ctrl.GetWinCamDPixelSize(0)
            pixelysize = self.gd.ctrl.GetWinCamDPixelSize(1)
            verticalPixel = self.gd.ctrl.GetVerticalPixels()
            horizontalPixel = self.gd.ctrl.GetHorizontalPixels()
            exposure = self.gd.ctrl.Exposure(0)
            fullresolution=self.gd.ctrl.CaptureIsFullResolution()
            # pixelsize = self.gd.ctrl.GetROI
            print pixelxsize, pixelysize, exposure, verticalPixel, horizontalPixel
            f.create_dataset("pixelSizeX", data = pixelxsize)
            f.create_dataset("pixelSizeY", data = pixelysize)
            f.create_dataset("exposureTime(ms)", data = exposure)
            f.create_dataset("verticalPixelNumber", data = verticalPixel)
            f.create_dataset("horizontalPixelNumber", data = horizontalPixel)            
            f.create_dataset("image", data = np.array(data).reshape(verticalPixel,horizontalPixel))
            f.create_dataset("time (s)", data = time.time() )
            f.create_dataset("time with format",data = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            f.create_dataset("isFullResolution",data = fullresolution)
            f.close()
            df = pd.DataFrame(np.array(data).reshape(verticalPixel,horizontalPixel))
            df.to_csv(pathname + 'data{}'.format(counter['count']) + '.csv',index = None,header = None)
            
            counter['count']+=1
            print('done')
        cnts = {'count': 0}
        schedule.every(2).minutes.do(job,cnts) 
        job(cnts)
        while True:
            schedule.run_pending()
            time.sleep(10)
            
    def __init__( self, redirect=False, filename=None ):
        wx.App.__init__( self, redirect, filename )
        self.frame = wx.Frame( parent=None, id=wx.ID_ANY,size=(760,500), title='Python Interface to DataRay')
        #Panel
        p = wx.Panel(self.frame, wx.ID_ANY)
        #Get Data
        self.gd = wx.lib.activex.ActiveXCtrl(p, 'DATARAYOCX.GetDataCtrl.1')
        #The methods of the object are available through the ctrl property of the item
        self.gd.ctrl.StartDriver()
        self.counter = 0
        sink = EventSink(self.frame)
        self.sink = comtypes.client.GetEvents(self.gd.ctrl, sink)
        #Button Panel
        bp = wx.Panel(parent=self.frame, id=wx.ID_ANY, size=(215, 250))
        b1 = wx.lib.activex.ActiveXCtrl(parent=bp,size=(200,50), pos=(7, 0),axID='DATARAYOCX.ButtonCtrl.1')
        b1.ctrl.ButtonID =297 #Id's for some ActiveX controls must be set
        b2 = wx.lib.activex.ActiveXCtrl(parent=bp,size=(100,25), pos=(5, 55),axID='DATARAYOCX.ButtonCtrl.1')
        b2.ctrl.ButtonID =171
        b3 = wx.lib.activex.ActiveXCtrl(parent=bp,size=(100,25), pos=(110,55),axID='DATARAYOCX.ButtonCtrl.1')
        b3.ctrl.ButtonID =172
        b4 = wx.lib.activex.ActiveXCtrl(parent=bp,size=(100,25), pos=(5, 85),axID='DATARAYOCX.ButtonCtrl.1')
        b4.ctrl.ButtonID =177
        b4 = wx.lib.activex.ActiveXCtrl(parent=bp,size=(100,25), pos=(110, 85),axID='DATARAYOCX.ButtonCtrl.1')
        b4.ctrl.ButtonID =179
        #Custom controls
        t = wx.StaticText(bp, label="File:", pos=(5, 115))
        self.ti = wx.TextCtrl(bp, value=r"C:\Users\zzp\Desktop\CoolDown\SeventhCooldown\data\\", pos=(30, 115), size=(170, -1))
        self.rb = wx.RadioBox(bp, label="Data:", pos=(5, 140), choices=["Profile", "WinCam"])
        self.cb = wx.ComboBox(bp, pos=(5,200), choices=[ "Profile_X", "Profile_Y", "Both"])
        self.cb.SetSelection(0)
        myb = wx.Button(bp, label="Write", pos=(5,225))
        myb.Bind(wx.EVT_BUTTON, self.OnClick)
        # t2 = wx.StaticText(bp, label="Save per # sec", pos=(55, 225))
        # self.ti2 = wx.TextCtrl(bp, value=r"600", pos=(75, 225), size=(30, -1))
        
        #Pictures
        pic = wx.lib.activex.ActiveXCtrl(parent=self.frame,size=(250,250),axID='DATARAYOCX.CCDimageCtrl.1')
        tpic = wx.lib.activex.ActiveXCtrl(parent=self.frame,size=(250,250), axID='DATARAYOCX.ThreeDviewCtrl.1')
        palette = wx.lib.activex.ActiveXCtrl(parent=self.frame,size=(10,250), axID='DATARAYOCX.PaletteBarCtrl.1')
        #Profiles
        self.px = wx.lib.activex.ActiveXCtrl(parent=self.frame,size=(300,200),axID='DATARAYOCX.ProfilesCtrl.1')
        self.px.ctrl.ProfileID=22
        self.py = wx.lib.activex.ActiveXCtrl(parent=self.frame,size=(300,200),axID='DATARAYOCX.ProfilesCtrl.1')
        self.py.ctrl.ProfileID = 23
        #Formatting 
        row1 = wx.BoxSizer(wx.HORIZONTAL)
        row1.Add(bp,proportion=1,flag=wx.RIGHT, border=10)
        row1.Add(pic,proportion=1)
        row1.Add(tpic, proportion=1,flag=wx.LEFT, border=10)
        row2 = wx.BoxSizer(wx.HORIZONTAL)
        row2.Add(self.px, 1, wx.RIGHT, 100)# Arguments: item, proportion, flags, border
        row2.Add(self.py,proportion=1)
        col1 = wx.BoxSizer(wx.VERTICAL)
        col1.Add(row1, proportion=1, flag=wx.BOTTOM, border=10)
        col1.Add(row2, proportion=1, flag=wx.ALIGN_CENTER_HORIZONTAL)
        self.frame.SetSizer(col1)
        self.frame.Show()

if __name__ == "__main__":
     app = MyApp()
     app.MainLoop()