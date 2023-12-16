# Single Image Blob Interface Accessible Control
# 
# Small overlays for Native Instruments synthes
#
# AZ (www.azslow.com), 2018
from ctypes import *
from ctypes.wintypes import *
import time
from logHandler import log
import winUser

from . import *

class GuitarRig5(SIBINVDA):

	sname = "Guitar Rig dialog" # to avoid "None", plug-in window already has the text

	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 1018, 810, ("block_arrows"))
		self.sibi.add( Label("Current preset", self.sibi, Pt(416, 57), Pt(565, 79), ("dynamic_text",) ) )
		self.sibi.add( VList("Tag", self.sibi, Pt(19,122), Pt(132,134), Pt(19,139), Pt(19,406,MOVE,2),
							0x242424, 0xfaaf2f, set(("click_on_enter",))) )
		self.sibi.add( VList("Subtag", self.sibi, Pt(138,122), Pt(251,134), Pt(138,139), Pt(138,406,MOVE,2),
							0x242424, 0xfaaf2f, set(("click_on_enter",))) )
		self.sibi.add( VList("Group", self.sibi, Pt(257,122), Pt(370,134), Pt(257,139), Pt(257,406,MOVE,2),
							0x242424, 0xfaaf2f, set(("click_on_enter",))) )
		self.sibi.add( VList("Preset", self.sibi, Pt(47,481,MOVE,2), Pt(193,497,MOVE,2), Pt(47,501,MOVE,2), Pt(47,765,MOVE),
							0x242424, 0xfaaf2f, set(("dblclick_on_enter", "page_keys", "homeend_keys"))) )


class Absynth5(SIBINVDA):

	sname = "Absynth 5"
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 764, 687, ("block_arrows",))
		self.sibi.add( Label("Current preset", self.sibi,  Pt(175, 59), Pt(378, 71), ("dynamic_text",) ) )
		self.sibi.add( VList("Bank", self.sibi, Pt(8,142), Pt(122,153), Pt(8,157), Pt(8,438,MOVE),
								0x0d5a5c, 0x0f999f, set(("click_on_enter","slow_reaction"))) )
		self.sibi.add( VList("Type", self.sibi, Pt(161,142), Pt(250,153), Pt(161,157), Pt(161,438,MOVE),
								0x0d5a5c, 0x0f999f, set(("click_on_enter","slow_reaction"))) )
		self.sibi.add( VList("SubType", self.sibi, Pt(284,142), Pt(366,153), Pt(284,157), Pt(284,438,MOVE),
								0x0d5a5c, 0x0f999f, set(("click_on_enter","slow_reaction","multi_one"))) )
		self.sibi.add( VList("Mode", self.sibi, Pt(406,142), Pt(502,153), Pt(406,157), Pt(406,438,MOVE),
								0x0d5a5c, 0x0f999f, set(("click_on_enter","slow_reaction","multi_one"))) )
		self.sibi.add( VList("Preset", self.sibi, Pt(536,143), Pt(719,155), Pt(536,161), Pt(536,434,MOVE),
								(0x093f42,0x2eafb1), (0x137072,0x115052), set(("dblclick_on_enter", "page_keys", "homeend_keys"))) )
