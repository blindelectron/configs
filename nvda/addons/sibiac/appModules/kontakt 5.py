# Single Image Blob Interface Accessible Control
#
# SIBIAC based NVDA Application Module
# Kontakt 5 Standalone
#
# AZ (www.azslow.com), 2017 - 2020
from __future__ import absolute_import
from __future__ import division
from __future__ import print_function


import appModuleHandler
from NVDAObjects.window import Window
import speech
from logHandler import log
import time
import queueHandler
import core
import oleacc
import controlTypes

from .sibiac import *

class PopupMenuButton_Kontakt(PopupMenuButton):
	def onEscape(self):
		if self._isOpen():
			box = Box(self.box.hwnd, 150, self.box.top)
			box.leftClick()
			time.sleep(0.05) # give time to close
			self.sibi.speakFocusAfter(self.reactionTime())
			return True
		return False

class ModeButton_Kontakt(Label):
	def __init__(self, name, sibi, left_top, right_bottom, click_point):
		super(ModeButton_Kontakt,self).__init__(name, sibi, left_top, right_bottom, ("dynamic_text",))
		self.clickbox = Box( sibi.hwnd, sibi.xScale(click_point), sibi.yScale(click_point) )
		
	def onEnter(self):
		self.clickbox.leftClick()
		if not "silent_action" in self.opt:
			self.sibi.speakAfter(self.reactionTime())
		return True
	
	def isEnabled(self):
		""" Return true in "Edit" mode """
		(type, name, text) = self.getTextInfo()
		return text == "Edit"


class ScriptEditor_Kontakt(SwitchBtn):
	def isEnabled(self):
		return self.getStateText() == self.on_text

class Kontakt5(SIBINVDA):

	def _get_name(self):
		if self.role == controlTypes.ROLE_UNKNOWN:
			return "Kontakt 5"  # We know it is from NI, no reason to remind us every time
		return super(Kontakt5,self).name # For menus
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 664, 640)
		
		self.sibi.add( PopupMenuButton_Kontakt("Load & Save", self.sibi, Pt(327, 4), Pt(357, 33),
						Pt(339, 32), Pt(339, 305), 0xf8f8f8, 0x3c3c3c,
						Pt(339, 35), Pt(515, 52), Pt(339, 53),
						0xf8f8f8, 0xcfcfcf, None, ("skip_spacers","slow_open") ) )
		instedit = Container("Edit mode", self.sibi, ("switchable",))
		self.sibi.add( instedit ) 
		instedit.add( ModeButton_Kontakt("Mode", self.sibi, Pt(5, 42), Pt(48, 54), Pt(30, 88) ) )
		scriptedit = Container("Script editor", self.sibi, ("switchable",))
		instedit.add( scriptedit ) 
		scriptedit.add( ScriptEditor_Kontakt("Script editor", self.sibi, Pt(513, 152), Pt(621,162), None, 0xe6e5dc, 0xbfb9aa, () ) )
		presetmenu = PopupMenuButton_Kontakt("Preset", self.sibi, Pt(19, 212), Pt(62, 220),
						Pt(19, 221), Pt(19, 280), 0xf8f8f8, 0x878782,
						Pt(20, 224), Pt(112, 241), Pt(20, 242),
						0xf8f8f8, 0xcfcfcf, None, ("skip_spacers","slow_open") )
		scriptedit.add( presetmenu )
		presetmenu.defineSubmenu( Pt(135, 0), 1, 10, 1, 1, 18, 18)
		
		

class AppModule(appModuleHandler.AppModule):

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		# Guitar Rig create "logo" first, with the same class but without text
		if "NINormalWindow" in obj.windowClassName and obj.role == controlTypes.ROLE_PANE and obj.windowText !="":
			clsList.insert(0, Kontakt5)
		elif obj.role == controlTypes.ROLE_PANE:
			chooseKnownOverlay(obj, clsList)
			#log.error("overlay not found " + str(obj.devInfo))
