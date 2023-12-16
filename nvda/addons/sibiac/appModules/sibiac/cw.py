# Single Image Blob Interface Accessible Control
# 
# Small overlays for Cakewalk plug-ins
#
# AZ (www.azslow.com), 2018
from ctypes import *
from ctypes.wintypes import *
import time
from logHandler import log
import winUser

from . import *


class SIElectricPiano(SIBINVDA):

	sname = "\t"
	displayText = ""
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 827, 587)
		self.sibi.add( PushBtn("Program", self.sibi, Pt(53, 81), Pt(147, 97), 0x000000, 0x303031, ("dynamic_text","silent_action")) )

class SIDrumKit(SIBINVDA):

	sname = "\t"
	displayText = ""
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 848, 588)
		self.sibi.add( PushBtn("Program", self.sibi, Pt(50, 82), Pt(146, 98), 0x000000, 0x303031, ("dynamic_text","silent_action")) )

class SIBassGuitar(SIBINVDA):

	sname = "\t"
	displayText = ""
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 827, 587)
		self.sibi.add( PushBtn("Program", self.sibi, Pt(53, 81), Pt(147, 97), 0x000000, 0x303031, ("dynamic_text","silent_action")) )

class SIStringSection(SIBINVDA):

	sname = "\t"
	displayText = ""
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 850, 570)
		self.sibi.add( PushBtn("Program", self.sibi, Pt(52, 82), Pt(146, 98), 0x000000, 0x303031, ("dynamic_text","silent_action")) )

class DimensionPro(SIBINVDA):

	sname = "\t"
	displayText = ""
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 752, 503)
		self.sibi.add( PushBtn("Program", self.sibi, Pt(37, 11), Pt(378, 25), 0x000000, 0x170f2d, ("dynamic_text","silent_action")) )

		
class SDKitPiece(Label):
	"""
	LeftClick on enter,
	Right click on down
	"""
	# Control Interface

	def onEnter(self):
		self.box.leftClick()
		return True

	def onDown(self):
		self.box.rightClick()
		return True

class SDOutput(Label):
	"""
	Combo with accessible list on clickLabel, support "dynamic_text"  and "silent_action"
	"""
	def __init__(self, name, sibi, left_top, right_bottom, combo_point):
		super(SDOutput,self).__init__( name, sibi, left_top, right_bottom, ("dynamic_text",))
		self.combo_box = Box( sibi.hwnd, sibi.xScale(combo_point), sibi.yScale(combo_point), sibi.xScale(combo_point), sibi.yScale(combo_point) )

	def onEnter(self):
		self.combo_box.leftClick()
		return True
		
class SessionDrummer(SIBINVDA):

	sname = "Session drummer"
	displayText = ""
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 964, 570)
		
		tabs = FixedTabControl("", self.sibi, ("slow_reaction",));
		self.sibi.add( tabs );
		
		drumkit = FixedTab("Drumkit", self.sibi, Pt(738, 29), Pt(806, 42), Pt(743, 34), 0x5d646d, 0xf3f8fd, ())
		drumkit.add( PushBtn("Program", self.sibi, Pt(112, 468), Pt(422, 482), 0x000000, 0x170f2d, ("dynamic_text","silent_action")) )
		drumkit.add( SDKitPiece("Kick", self.sibi, Pt(450, 350), Pt(479, 378), set()) )
		drumkit.add( SDKitPiece("Snare", self.sibi, Pt(290, 313), Pt(305, 327), set()) )
		drumkit.add( SDKitPiece("Hi Hat", self.sibi, Pt(181, 264), Pt(196, 276), set()) )
		drumkit.add( SDKitPiece("Hi tom", self.sibi, Pt(374, 173), Pt(396, 191), set()) )
		drumkit.add( SDKitPiece("Mid tom", self.sibi, Pt(516, 171), Pt(548, 196), set()) )
		drumkit.add( SDKitPiece("Low tom", self.sibi, Pt(646, 332), Pt(670, 349), set()) )
		drumkit.add( SDKitPiece("Crash", self.sibi, Pt(248, 151), Pt(262, 167), set()) )
		drumkit.add( SDKitPiece("Ride", self.sibi, Pt(724, 231), Pt(751, 251), set()) )
		drumkit.add( SDKitPiece("FX 1", self.sibi, Pt(128, 339), Pt(136, 358), set()) )
		drumkit.add( SDKitPiece("FX 2", self.sibi, Pt(162, 338), Pt(170, 346), set()) )
		drumkit.add( SDKitPiece("FX 3", self.sibi, Pt(139, 381), Pt(151, 391), set()) )
		drumkit.add( SDKitPiece("FX 4", self.sibi, Pt(175, 371), Pt(184, 379), set()) )
		tabs.add( drumkit )

		mixer = FixedTab("Mixer", self.sibi, Pt(817, 29), Pt(884, 42), Pt(823, 34), 0x5d646d, 0xf3f8fd, ())
		mixer.add( SDOutput("Kick output", self.sibi, Pt(53, 530), Pt(77, 539), Pt(85, 534)) )
		mixer.add( SDOutput("Snare output", self.sibi, Pt(107, 530), Pt(147, 539), Pt(152, 534)) )
		mixer.add( SDOutput("Hi Hat output", self.sibi, Pt(176, 530), Pt(218, 539), Pt(223, 534)) )
		mixer.add( SDOutput("Hi tom output", self.sibi, Pt(248, 530), Pt(288, 539), Pt(295, 534)) )
		mixer.add( SDOutput("Mid tom output", self.sibi, Pt(319, 530), Pt(358, 539), Pt(363, 534)) )
		mixer.add( SDOutput("Low tom output", self.sibi, Pt(389, 530), Pt(429, 539), Pt(435, 534)) )
		mixer.add( SDOutput("Crash output", self.sibi, Pt(462, 530), Pt(498, 539), Pt(504, 534)) )
		mixer.add( SDOutput("Ride output", self.sibi, Pt(529, 530), Pt(568, 539), Pt(574, 534)) )
		mixer.add( SDOutput("FX 1 output", self.sibi, Pt(605, 530), Pt(638, 539), Pt(644, 534)) )
		mixer.add( SDOutput("FX 2 output", self.sibi, Pt(670, 530), Pt(708, 539), Pt(715, 534)) )
		mixer.add( SDOutput("FX 3 output", self.sibi, Pt(737, 530), Pt(778, 539), Pt(784, 534)) )
		mixer.add( SDOutput("FX 4 output", self.sibi, Pt(807, 530), Pt(848, 539), Pt(854, 534)) )
		tabs.add( mixer )


class SIProgramBrowser(object):

	name = "Program browser"
	
	def event_gainFocus(self):
		try:
			self.firstChild.firstChild.setFocus()
		except:
			pass
