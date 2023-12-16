# Single Image Blob Interface Accessible Control
# 
# Pure Synth Platinum overlays 
#
# AZ (www.azslow.com), 2018
from ctypes import *
from ctypes.wintypes import *
import time
from logHandler import log
import winUser

from . import *

class PSP_Selector(Control):
	def __init__(self, name, sibi, points, off_color, on_color, opt = None, getShift = None):
		super(PSP_Selector,self).__init__(name, sibi, opt)
		self.__getShift = getShift
		self.points = []
		for pt in points:
			self.points.append(XY(sibi.xScale(pt), sibi.yScale(pt)))
		self.off_color = Color2Tuple(off_color)
		self.on_color = Color2Tuple(on_color)
	
	def _getCurrent(self):
		dx, dy = self.getShift()
		for i, xy in enumerate(self.points):
			j = FindNearestColor(self.sibi.hwnd, xy.x + dx, xy.y + dy, self.off_color + self.on_color)
			if j >= len(self.off_color):
				return i
		return -1
	
	def _setCurrent(self, delta):
		i = self._getCurrent()
		if i < 0:
			return False
		i += delta
		if i < 0:
			i = len(self.points) - 1
		elif i >= len(self.points):
			i = 0
		dx, dy = self.getShift()
		box = Box(self.sibi.hwnd, self.points[i].x + dx, self.points[i].y + dy)
		box.leftClick()
		self.sibi.speakAfter(self.reactionTime())		
		return True
	
	def getShift(self):
		if self.__getShift is None:
			return (0, 0)
		return self.__getShift()
		
	def getTextInfo(self):
		return (self.type, self.name, "%d" % (self._getCurrent() + 1))
		
	def onUp(self):
		return self._setCurrent(1)
		
	def onDown(self):
		return self._setCurrent(-1)
		
class PSP(SIBINVDA):
	sname = "Pure Synth Platinum"
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 1001, 864, set())
		self.sibi.add( SpinLabel("Preset", self.sibi, Pt(226, 39), Pt(475, 60), Pt(156, 46), Pt(192, 47), ("click_on_enter","slow_reaction")) )
		
		#####
		dlg = Dialog("Preset selection", self.sibi, Pt(910, 102), 0xafafaf, 0x1b1b1b, None)
		self.sibi.addDialog(dlg)
		dlg.add( VList("Library", self.sibi, Pt(50, 198), Pt(331, 213), Pt(50, 222), Pt(50, 564),
								(0x393939, 0x7490a5), 0x567892, set(("click_to_focus",))) )
		dlg.add( VList("Bank", self.sibi, Pt(338, 198), Pt(616, 213), Pt(338, 222), Pt(338, 564),
								(0x3a3a3a, 0x7590a5), 0x557791, set(("click_to_focus",))) )
		dlg.add( VList("Preset", self.sibi, Pt(626, 198), Pt(907, 213), Pt(626, 222), Pt(626, 564),
								(0x3a3a3a, 0x7590a5), 0x557791, set(("click_to_focus",))) )
		dlg.add( CloseBtn("Close", self.sibi, Pt(909, 83), Pt(909, 83), None) )

		self.sibi.add( PSP_Selector("Current oscillator", self.sibi, 
		                            (Pt(356, 145), Pt(398, 145), Pt(435, 145), Pt(477, 145)),
									0x141414, 0x333333) )
		self.sibi.add( SpinLabel("Oscillator", self.sibi, Pt(76, 121), Pt(223, 139), Pt(270, 129), Pt(317, 130), ("click_on_enter","slow_reaction")) )
		
		#####
		dlg = Dialog("Oscillator selection", self.sibi, Pt(899, 201), 0xababab, 0x1b1b1b, None)
		self.sibi.addDialog(dlg)
		dlg.add( VList("Type", self.sibi, Pt(50,223), Pt(331,238), Pt(50,247), Pt(50, 639),
								(0x393939, 0x7490a5), 0x567892, set(("click_to_focus",))) )
		dlg.add( VList("Category", self.sibi, Pt(338,223), Pt(616,238), Pt(338,247), Pt(338, 639),
								(0x393939, 0x7490a5), 0x567892, set(("click_to_focus",))) )
		dlg.add( VList("Sound", self.sibi, Pt(626,223), Pt(907,238), Pt(626,247), Pt(626, 639),
								(0x393939, 0x7490a5), 0x567892, set(("click_to_focus",))) )
		dlg.add( CloseBtn("Close", self.sibi, Pt(911, 155), Pt(911, 155), None) )
