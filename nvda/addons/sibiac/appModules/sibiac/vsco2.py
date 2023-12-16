# Single Image Blob Interface Accessible Control
# 
# VDCO2 overlay
#
# AZ (www.azslow.com), 2018 - 2020
from ctypes import *
from ctypes.wintypes import *
import time
from logHandler import log
import winUser

from . import *

class VSCO2_InstrumentList(object):
	""" That is popup JUCE window which support up/down arrows """
	def __init__(self):
		self.box = TextBox(0, 20, 0) # just for cache...
		
	def getHWnd(self):
		menu = winUser.user32.FindWindowA(0, b("menu"))
		if menu:
			return menu
		return None

	def getSelectionText(self):
		menu = self.getHWnd()
		if menu is None:
			return ""
		# find the height
		r = RECT()
		winUser.user32.GetClientRect(menu, byref(r))
		y0, y1 = FindVRange(menu, 3, 3, r.bottom - 4, 0xffffff, 0x7070cc)
		if y0 < 0:
			return "No selection"
		log.error("%d %d"%(y0,y1))
		self.box.hwnd = menu
		self.box.top = y0
		self.box.right = r.right - 1
		self.box.bottom = y1
		return self.box.getText()
		
class VSCO2_InstrumentBtn(Label):
	def __init__(self, name, sibi):
		super(VSCO2_InstrumentBtn,self).__init__(name, sibi, Pt(149, 118), Pt(430, 134), ("dynamic_text", "silent_action"))
		self.ilist = VSCO2_InstrumentList()
		self.load_pt = sibi.ptScale(Pt(288, 125))
		
	def __isLoading(self):
		i = FindNearestColor(self.sibi.hwnd, self.load_pt.x, self.load_pt.y, (0xd3d3d3, 0x000000))
		return i == 0
		
	def __inList(self):
		return self.ilist.getHWnd() is not None
		
	def getTextInfo(self):
		if self.__inList():
			return (self.type, "Instrument list", self.ilist.getSelectionText())
		if self.__isLoading():
			self.sibi.speakFocusAfter(2000)
			return ("", "", "Loading")
		return super(VSCO2_InstrumentBtn,self).getTextInfo()
		
	def onEnter(self):
		if self.__inList():
			self.sibi.speakFocusAfter(self.reactionTime())
			return False
		self.box.getBox().leftClick()
		self.sibi.speakFocusAfter(self.reactionTime())
		return True

	def onUp(self):
		self.sibi.speakAfter(self.reactionTime())
		return not self.__inList()
		
	def onDown(self):
		self.sibi.speakAfter(self.reactionTime())
		return not self.__inList()
		
		
class VSCO2(SIBINVDA):

	sname = "VSCO2"
	displayText = ""
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 800, 400)
		self.sibi.add( VSCO2_InstrumentBtn("Instrument", self.sibi) )
		self.sibi.add( Label("Articulation", self.sibi, Pt(147, 158), Pt(266, 176), ("dynamic_text", "speak_on_enter")) )
