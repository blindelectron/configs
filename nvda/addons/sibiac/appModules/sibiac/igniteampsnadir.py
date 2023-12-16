# Single Image Blob Interface Accessible Control
# 
# ignite-amps NadIR overlay
#
# Nikolai Bukatin, 2018
from logHandler import log

from . import *

class NadIRCabControl(Container):
	def __init__(self, name, sibi, browseXShift=0, laXShift=0, irXShift=0, raXShift=0):
		super(NadIRCabControl,self).__init__(name, sibi, ())
		
		self.add(Clickable("Browse...", sibi, Pt(124+browseXShift, 167), Pt(145+browseXShift, 175), ("silent_action", )) )
		self.add(Clickable("Previous IR", sibi, Pt(100+laXShift, 182), Pt(112+laXShift, 196), ("silent_action", )) )
		self.add(Clickable("", sibi, Pt(121+irXShift, 181), Pt(347+irXShift, 198), ("dynamic_text", "silent_action")))
		self.add(Clickable("Next IR", sibi, Pt(352+raXShift, 180), Pt(361+raXShift, 196), ("silent_action", )) )
	

class IgniteAmpsNadIR(SIBINVDA):

	sname = "Ignite-amps NadIR"
	displayText = ""

	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 890, 323)
		controls = Group("Page", self.sibi)
		self.sibi.add(controls)
		
		c1 = NadIRCabControl("Cab 1", self.sibi)
		controls.add(c1)
		
		c2 = NadIRCabControl("Cab 2", self.sibi, browseXShift=421, laXShift=424, irXShift=420, raXShift=421)
		controls.add(c2)
	

