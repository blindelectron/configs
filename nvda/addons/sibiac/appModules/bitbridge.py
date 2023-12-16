# Single Image Blob Interface Accessible Control
#
# SIBIAC based NVDA Application Module
# Sonar X3
#
# AZ (www.azslow.com), 2017 - 2020
from __future__ import absolute_import
from __future__ import division
from __future__ import print_function

import appModuleHandler
from NVDAObjects.window import Window
from NVDAObjects.IAccessible import IAccessible
import speech
from logHandler import log
import eventHandler
import time
import queueHandler
import core
import oleacc
import controlTypes

from .sibiac import *

class SoundCenter(SIBINVDA):

	def _get_name(self):
		if not hasattr(self, "sibi"):
			return super(SoundCenter,self)._get_name()
		if self.sibi.hwnd == self.windowHandle:
			return "Cakewalk Sound Center"
		return "\t"
	
	def defineSIBI(self):
		if self.hasKnownSIBI():
			return
		#log.error("New Sound Center")
		#log.error(self.devInfo)
		
		self.sibi = SIBI(self.windowHandle, 867, 614)
		self.sibi.nvda_class = SoundCenter
		self.sibi.add( Label("Current preset", self.sibi,  Pt(655, 79), Pt(841, 94), ("dynamic_text",) ) )
		self.sibi.add( VList("Type", self.sibi, Pt(21,99), Pt(190,112), Pt(21,115), Pt(21,368),
								(0x313130, 0x181717), 0x5f5135, set(("click_on_enter",))) )
		self.sibi.add( VList("SubType", self.sibi, Pt(194,99), Pt(352,112), Pt(194,115), Pt(194,368),
								(0x313130, 0x181717), 0x5f5135, set(("click_to_enter",))) )
		self.sibi.add( VList("Program", self.sibi, Pt(368, 99), Pt(550,112), Pt(368,115), Pt(268,368),
								(0x313130, 0x181717), 0x5f5135, set(("click_to_focus", "pass_enter"))) )

	
class AppModule(appModuleHandler.AppModule):

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		"""
		Important that we get 2 (!) objects per window, one with ROLE_WINDOW and another with ROLE_XXX
		I am not sure what is better... but lets ignore ROLE_WINDOW
		"""
		global gUnknownNI
		if obj.role == controlTypes.ROLE_WINDOW:
			return
		if obj.windowClassName == "SHClass3" and obj.windowText == "Cakewalk Sound Center":
			clsList.insert(0, SoundCenter)
		elif chooseKnownOverlay(obj, clsList):
			pass
		else:
			pass
			#log.error(obj.devInfo)
			
	def isBadUIAWindow(self,hwnd):
		"""
		If accessibility is enabled in Sonar, it tries to speak too much.
		AZ Controller does that better, so I prefer to switch UIA off
		"""
		return True
