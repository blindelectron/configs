# Single Image Blob Interface Accessible Control
# 
# Zampler overlays 
#
# AZ (www.azslow.com), 2018
from ctypes import *
from ctypes.wintypes import *
import time
from logHandler import log
import winUser

from . import *

class Zampler_ModifierChooser(Control):
	def __init__(self, name, sibi, container, opt = None):
		super(Zampler_ModifierChooser,self).__init__(name, sibi, opt);
		self.container = container
		self.n = 1 # current modifier
		self.getShift = lambda : self._getShift()

	def _getShift(self):
		return (315, 241 + (self.n - 1)*14)
		
	def getTextInfo(self):
		text = []
		for ctrl in self.container.ctrl:
			if ctrl == self:
				text.append( "Mod %d" % self.n )
			else:
				type, name, txt = ctrl.getTextInfo()
				if txt != "Empty":
					text.append( txt )
		return (self.type, self.name, ", ".join(text))

	def reactionTime(self):
		return 10
		
	def onUp(self):
		self.n -= 1
		if self.n < 1:
			self.n = 12
		self.sibi.speakAfter(self.reactionTime())
		return True
	
	def onLeft(self):
		return self.onUp()
		
	def onDown(self):
		self.n += 1
		if self.n > 12:
			self.n = 1
		self.sibi.speakAfter(self.reactionTime())
		return True

	def onRight(self):
		return self.onDown()
		
class Zampler_LoadSFZ(Label):
	def __init__(self, name, sibi, left_top, right_bottom, click_pt):
		super(Zampler_LoadSFZ,self).__init__(name, sibi, left_top, right_bottom, ("dynamic_text",))
		self.click_box = Box(sibi.hwnd, sibi.xScale(click_pt), sibi.yScale(click_pt))
		
	def onEnter(self):
		self.click_box.getBox().leftClick()

class Zampler_WheelLabel(Label):

	def __init__(self, name, sibi, left_top, right_bottom, opt = None, getShift = None):
		super(Zampler_WheelLabel,self).__init__(name, sibi, left_top, right_bottom, ("dynamic_text",), getShift)
		
	def onUp(self):
		self.box.getBox().moveTo()
		MouseScroll(120)
		self.sibi.speakAfter(self.reactionTime())
		return True

	def onDown(self):
		self.box.getBox().moveTo()
		MouseScroll(-120)
		self.sibi.speakAfter(self.reactionTime())
		return True

class Zampler_FilterType(SpinLabel):

	def onEnter(self):
		return True
		
class Zampler(SIBINVDA):
	sname = "Zampler"
	displayText = "" # garbage here
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 840, 611, ("block_arrows",))
		tabs = FixedTabControl("Control page", self.sibi, None)
		main_page = FixedTab("Main", self.sibi, Pt(293, 204), Pt(368, 217), Pt(296, 210), 0x8faac5, 0x39444e, None)
		main_page.add( SpinLabel("Patch", self.sibi, Pt(294, 260), Pt(546, 279), Pt(400, 230), Pt(440, 230), ("click_on_enter","slow_reaction")) )
		main_page.add( Clickable("Load bank", self.sibi, Pt(313, 243), Pt(313, 243), ("silent_action,") ) )
		main_page.add( Clickable("Save bank", self.sibi, Pt(353, 243), Pt(353, 243), ("silent_action,") ) )
		main_page.add( Clickable("Load patch", self.sibi, Pt(400, 243), Pt(400, 243), ("silent_action,") ) )
		main_page.add( Clickable("Save patch", self.sibi, Pt(441, 243), Pt(441, 243), ("silent_action,") ) )
		main_page.add( Zampler_LoadSFZ("SFZ/REX", self.sibi, Pt(463, 224), Pt(547, 235), Pt(500, 245) ) )
		main_page.add( Zampler_WheelLabel("Polyphony", self.sibi, Pt(334, 367), Pt(361, 379)) )
		main_page.add( Zampler_WheelLabel("Bend up", self.sibi, Pt(339, 394), Pt(360, 405)) )
		main_page.add( Zampler_WheelLabel("Bend down", self.sibi, Pt(430, 394), Pt(460, 405)) )
		tabs.add( main_page )
		
		matrix_page = FixedTab("Mod matrix", self.sibi, Pt(379, 206), Pt(454, 216), Pt(380, 210), 0x8faac5, 0x39444e, None)
		chooser = Zampler_ModifierChooser("Choose", self.sibi, matrix_page, None)
		matrix_page.add( chooser )
		matrix_page.add( Clickable("Source", self.sibi, Pt(0, 0), Pt(61, 12), ("dynamic_text",), chooser.getShift) )
		matrix_page.add( Zampler_WheelLabel("Amount", self.sibi, Pt(63, 0), Pt(146, 12), ("dynamic_text",), chooser.getShift) )
		matrix_page.add( Clickable("Distination", self.sibi, Pt(148, 0), Pt(231, 12), ("dynamic_text",), chooser.getShift) )
		tabs.add( matrix_page )
		
		arpeggiator_page = FixedTab("Arpeggiator", self.sibi, Pt(465, 205), Pt(544, 217), Pt(466, 210), 0x8faac5, 0x39444e, None)
		arpeggiator_page.add( Clickable("Pattern", self.sibi, Pt(487, 227), Pt(574, 241), ("dynamic_text",)) )
		tabs.add( arpeggiator_page )
		
		self.sibi.add( tabs )
		
		self.sibi.add( SpinLabel("Patch", self.sibi, Pt(680, 96), Pt(753, 108), Pt(765, 97), Pt(765, 106), ("slow_reaction",)) )
