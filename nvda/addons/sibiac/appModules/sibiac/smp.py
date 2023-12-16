# Single Image Blob Interface Accessible Control
# 
# SynthMaster Player overlay
#
# AZ (www.azslow.com), 2018
from ctypes import *
from ctypes.wintypes import *
import time
from logHandler import log
import winUser

from . import *

class SMP_Preset(Clickable):
	def __init__(self, sibi):
		super(SMP_Preset,self).__init__("Preset", sibi, Pt(476, 18), Pt(634, 34), ("dynamic_text", "silent_action"))
		self.prev_pt = Box(sibi.hwnd, sibi.xScale(646), sibi.yScale(26))
		self.next_pt = Box(sibi.hwnd, sibi.xScale(662), sibi.yScale(26))

	def onLeft(self):
		self.prev_pt.getBox().leftClick()
		self.sibi.speakAfter(self.reactionTime())
		return True
		
	def onUp(self):
		return self.onLeft()

	def onRight(self):
		self.next_pt.getBox().leftClick()
		self.sibi.speakAfter(self.reactionTime())
		return True
		
	def onDown(self):
		return self.onRight()

class SMP_Settings(Label):
	def __init__(self, sibi):
		super(SMP_Settings,self).__init__("Settings", sibi, Pt(479, 38), Pt(619, 54), ("dynamic_text",))
		self.type = "Settings"
		self.prev_pt = Box(sibi.hwnd, sibi.xScale(646), sibi.yScale(45))
		self.next_pt = Box(sibi.hwnd, sibi.xScale(662), sibi.yScale(45))
		self.value = TextBox(sibi.hwnd, sibi.xScale(480), sibi.yScale(56), sibi.xScale(624), sibi.yScale(67))

	def getTextInfo(self):
		type, dummy, name = super(SMP_Settings,self).getTextInfo()
		text = self.value.getBox().getText()
		return (self.type, name, text)
		
	def onLeft(self):
		self.prev_pt.getBox().leftClick()
		self.sibi.speakInFocusAfter(self.reactionTime())
		return True
		
	def onUp(self):
		return self.onLeft()

	def onRight(self):
		self.next_pt.getBox().leftClick()
		self.sibi.speakInFocusAfter(self.reactionTime())
		return True
		
	def onDown(self):
		return self.onRight()

	def onEnter(self):
		self.value.getBox().leftClick()
		return True

class SMP_Knob(Label):
	def __init__(self, name, sibi, pt):
		super(SMP_Knob,self).__init__(name, sibi, Pt(60 + pt.x - 95, 470 + pt.y - 505), Pt(137 + pt.x - 95, 484 + pt.y - 505), ("dynamic_text",))
		sibi = self.sibi
		self.pt = Box(sibi.hwnd, sibi.xScale(pt), sibi.yScale(pt))

	def isFocusable(self):
		return False
		
	def getNameInGroup(self):
		type, dummy, name = super(SMP_Knob,self).getTextInfo()
		return ("%s, %s" % (self.name, name), "")
		
	def onUp(self):
		self.pt.getBox().moveTo()
		MouseScroll(120)
		#self.sibi.speakAfter(self.reactionTime())
		return True

	def onDown(self):
		self.pt.getBox().moveTo()
		MouseScroll(-120)
		#self.sibi.speakAfter(self.reactionTime())
		return True

class SMP_X(Label):
	def __init__(self, name, sibi, pt):
		super(SMP_X,self).__init__(name, sibi, Pt(400 + pt.x - 410, 598 + pt.y - 593), Pt(467 + pt.x - 410, 610 + pt.y - 593), ("dynamic_text",))
		sibi = self.sibi

	def isFocusable(self):
		return False
		
	def getNameInGroup(self):
		type, dummy, name = super(SMP_X,self).getTextInfo()
		return ("%s, %s, not implemented" % (self.name, name), "")

class SMP_Y(Label):
	def __init__(self, name, sibi, x):
		super(SMP_Y,self).__init__(name, sibi, Pt(x.box.left + 68, x.box.top), Pt(x.box.right + 68, x.box.bottom), ("dynamic_text",))
		sibi = self.sibi
		self.x = x

	def isFocusable(self):
		return False
		
	def getNameInGroup(self):
		type, dummy, name = super(SMP_Y,self).getTextInfo()
		return ("%s, %s, not implemented" % (self.name, name), "")
		
		
class SMP(SIBINVDA):

	sname = "SynthMaster Player"
	displayText = ""
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 960, 640)
		sibi = self.sibi
		topgroup = Group("Page", sibi)
		sibi.add( topgroup )
		
		presets_page = Container("Presets", sibi, None)
		topgroup.add( presets_page )
		presets_page.add( SMP_Preset(sibi) )
		presets_page.add( Clickable("Search", sibi, Pt(609, 138), Pt(680, 151), ("dynamic_text", "silent_action")) )
		presets_page.add( Label("Filters are not implemented", sibi, Pt(0, 0), Pt(0, 0), None) )
		
		controls_page = Container("Controls", sibi, None)
		topgroup.add( controls_page )
		control_group = Group("Control", sibi)
		controls_page.add( control_group )
		control_group.add( SMP_Knob("1", sibi, Pt(95, 505)) )
		control_group.add( SMP_Knob("2", sibi, Pt(178, 505)) )
		control_group.add( SMP_Knob("3", sibi, Pt(261, 505)) )
		control_group.add( SMP_Knob("4", sibi, Pt(344, 505)) )
		control_group.add( SMP_Knob("5", sibi, Pt(95, 571)) )
		control_group.add( SMP_Knob("6", sibi, Pt(178, 571)) )
		control_group.add( SMP_Knob("7", sibi, Pt(261, 571)) )
		control_group.add( SMP_Knob("8", sibi, Pt(344, 571)) )
		x = SMP_X("X1", sibi, Pt(410, 593))
		control_group.add( x )
		control_group.add( SMP_Y("Y1", sibi, x) )
		x = SMP_X("X2", sibi, Pt(576, 593))
		control_group.add( x )
		control_group.add( SMP_Y("Y2", sibi, x) )

		layers_page = Container("Layers", sibi, None)
		topgroup.add( layers_page )
		layers_page.add( Label("Not implemented yet", sibi, Pt(0, 0), Pt(0, 0), None) )

		fx_page = Container("Global FX", sibi, None)
		topgroup.add( fx_page )
		fx_page.add( Label("Not implemented yet", sibi, Pt(0, 0), Pt(0, 0), None) )

		settings_page = Container("Settings", sibi, None)
		topgroup.add( settings_page )
		settings_page.add( SMP_Settings(sibi) )
		settings_page.add( Label("Velocity curve is not implemented", sibi, Pt(0, 0), Pt(0, 0), None) )

		
