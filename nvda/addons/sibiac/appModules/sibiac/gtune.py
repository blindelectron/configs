# Single Image Blob Interface Accessible Control
# 
# GTune overlays 
#
# AZ (www.azslow.com), 2018
from ctypes import *
from ctypes.wintypes import *
import time
from logHandler import log
import winUser

from . import *

class GTuneCurrent(Control):
	def __init__(self, sibi):
		super(GTuneCurrent,self).__init__("Tune", sibi, None)
		self.box1 = TextBox(sibi.hwnd, sibi.xScale(175), sibi.yScale(225), sibi.xScale(209), sibi.yScale(243))
		self.box2 = TextBox(sibi.hwnd, sibi.xScale(232), sibi.yScale(225), sibi.xScale(278), sibi.yScale(243))
	
	def getTextInfo(self):
		note = self.box1.getText()
		tune = self.box2.getText()
		if note == "":
			text = "No signal"
		else:
			note = note.replace("#", " sharp")
			text = "%s %s" % (note, tune)
		return ( self.type, self.name, text )

	def onEnter(self):
		self.sibi.speakAfter(self.reactionTime())
		return True

class GTune(SIBINVDA):
	sname = "GTune"
	displayText = ""
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 300, 277, ("block_arrows", "textout"))
		self.sibi.add( Clickable("Reference frequency", self.sibi, Pt(231, 37), Pt(283, 47), ("dynamic_text", "silent_action")) )
		self.sibi.add( GTuneCurrent(self.sibi) )
