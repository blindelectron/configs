# Single Image Blob Interface Accessible Control
#
# SIBIAC based NVDA Application Module
# X-Touch Compact/Mini Editor
#
# AZ (www.azslow.com), 2017
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

from sibiac import SIBI, SIBINVDA, VList, Label, PushBtn, SwitchBtn, FixedTab, FixedTabControl, Combo, OptionTable, Pt, FIXED, MOVE, PROPORTIONAL, gSIBI


class ConfirmationDlg(Window):
	role = controlTypes.ROLE_DIALOG
	name = "Confirm preset loading"

class XTouchEditor(SIBINVDA):

	def defineSIBI(self):
		if self.hasKnownSIBI():
			return

		self.sibi = SIBI(self.windowHandle, 702, 678)
		self.sibi.add( Label("Info", self.sibi, Pt(228,515,PROPORTIONAL), Pt(521,534,PROPORTIONAL),
								set(("dynamic_text",))) )
		self.sibi.add( SwitchBtn("Standard mode", self.sibi, Pt(64,281,PROPORTIONAL), Pt(114,285,PROPORTIONAL),
								None, 0x323232, 0x5a4f39, set()) )
		self.sibi.add( SwitchBtn("Mackie mode", self.sibi, Pt(154,281,PROPORTIONAL), Pt(200,285,PROPORTIONAL),
								None, 0x323232, 0x5a4f39, set()) )
		self.sibi.add( PushBtn("Load", self.sibi, Pt(50,153,PROPORTIONAL), Pt(96,158,PROPORTIONAL),
								0x101010, 0x323232, set(("silent_action",))) )
		self.sibi.add( PushBtn("Upload A", self.sibi, Pt(412,153,PROPORTIONAL), Pt(458,158,PROPORTIONAL),
								0x101010, 0x323232, set()) )
		self.sibi.add( PushBtn("Upload B", self.sibi, Pt(486,153,PROPORTIONAL), Pt(539,158,PROPORTIONAL),
								0x101010, 0x323232, set()) )


class AppModule(appModuleHandler.AppModule):

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		"""
		Important that we get 2 (!) objects per window, one with ROLE_WINDOW and another with ROLE_XXX
		I am not sure what is better... but lets ignore ROLE_WINDOW
		"""
		if obj.role == controlTypes.ROLE_WINDOW:
			return
		if obj.windowText == " X-TOUCH Editor":
			clsList.insert(0, XTouchEditor)
		elif obj.windowText == "        LOAD preset?":
			clsList.insert(0, ConfirmationDlg)
		#else:
		#	log.error("'" + obj.windowText + "'")
