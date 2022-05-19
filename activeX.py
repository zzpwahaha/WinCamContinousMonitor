import wx.lib.activex
class MyApp( wx.App ):
	def __init__( self, redirect=False, filename=None ):
		wx.App.__init__( self, redirect, filename )
		self.frame = wx.Frame( parent=None, id=wx.ID_ANY,size=(500,500), title='Python Interface to DataRay')
		#Panel
		p = wx.Panel(self.frame, wx.ID_ANY)
		#Get Data
		self.gd = wx.lib.activex.ActiveXCtrl(p, 'DATARAYOCX.GetDataCtrl.1')
		self.gd.ctrl.StartDriver()

	