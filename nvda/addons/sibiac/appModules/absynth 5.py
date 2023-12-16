# Single Image Blob Interface Accessible Control
#
# SIBIAC based NVDA Application Module
# Absynth 5 Standalone
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

class Absynth5AudioMidi(SIBINVDA):

	def _get_name(self):
		if self.role == controlTypes.ROLE_UNKNOWN:
			return "\t" # I DO NOT WANT to speak anything there, but how???
		return super(Absynth5AudioMidi,self).name

	def defineSIBI(self):
		"""
		There is a panel inside dialog, since I could not get throw NVDA logic, I had to defined this overlay class for both...
		"""
		if self.hasKnownSIBI():
			return
		
		self.sibi = SIBI(self.windowHandle, 500, 342)
		self.sibi.nvda_class = Absynth5AudioMidi
		tc = FixedTabControl("Settings", self.sibi, ());
		self.sibi.add( tc );
		audio = FixedTab("Audio", self.sibi, Pt(12, 5), Pt( 88, 27), Pt(53, 30), 0xdcdcdc, 0xffffff, ())
		audio.add(Combo("Driver,", self.sibi, Pt(213, 61), Pt(458, 75), Pt(211, 83), Pt(0,0), 0xe6e6e6, (0x000000, 0xdedede, 0xf2f2f2), Pt(213,83), Pt(458,100), Pt(213, 101), 0xe6e6e6, 0xffffff, ()))
		audio.add(Combo("Device,", self.sibi, Pt(213, 108), Pt(458, 122), Pt(211, 130), Pt(0,0), 0xe6e6e6, (0x000000, 0xdedede, 0xf2f2f2), Pt(213,130), Pt(458,147), Pt(213, 148), 0xe6e6e6, 0xffffff, ()))
		audio.add(Label("Status", self.sibi, Pt(214,156), Pt(471,172), ("dynamic_text",)))
		audio.add(Combo("Sample Rate", self.sibi, Pt(213, 202), Pt(458, 216), Pt(211, 224), Pt(0,0), 0xe6e6e6, (0x000000, 0xdedede, 0xf2f2f2), Pt(213,224), Pt(458,241), Pt(213, 242), 0xe6e6e6, 0xffffff, ()))
		audio.add(Label("Latency", self.sibi, Pt(98,244), Pt(266,264), ("dynamic_text",)))
		audio.add(Label("Overal latency", self.sibi, Pt(419,277), Pt(475,291), ("dynamic_text",)))
		tc.add(audio)
		
		routing = FixedTab("Routing", self.sibi, Pt(115, 5), Pt( 188, 27), Pt(150, 30), 0xdcdcdc, 0xffffff, ())
		rtc = FixedTabControl("", self.sibi, ());
		routing.add( rtc )
		inputs = FixedTab("Inputs", self.sibi, Pt(28, 61), Pt(83, 74), Pt(25, 68), 0xe5e9eb, 0x98b7c7, ())
		inputs.add(OptionTable("Inputs", self.sibi, Pt(39,91), Pt(200,108), Pt(39,110), Pt(39,285),
		           Combo("", self.sibi, Pt(241,91), Pt(465,108), Pt(225,110), Pt(0,0), 0xe6e6e6 , 0xffffff, Pt(227,112), Pt(453,129), Pt(227,130), 0xe6e6e6, 0xffffff, ()),
				   ()))
		rtc.add(inputs)
		outputs = FixedTab("Outputs", self.sibi, Pt(89, 61), Pt(144, 74), Pt(88, 68), 0xe5e9eb, 0x98b7c7, ())
		outputs.add(OptionTable("Outputs", self.sibi, Pt(39,91), Pt(200,108), Pt(39,110), Pt(39,285),
		           Combo("", self.sibi, Pt(241,91), Pt(465,108), Pt(225,110), Pt(0,0), 0xe6e6e6 , 0xffffff, Pt(227,112), Pt(453,129), Pt(227,130), 0xe6e6e6, 0xffffff, ()),
				   ()))
		rtc.add(outputs)
		tc.add(routing)
		
		midi = FixedTab("MIDI", self.sibi, Pt(210, 5), Pt( 290, 27), Pt(250, 30), 0xdcdcdc, 0xffffff, ())
		mtc = FixedTabControl("", self.sibi, ());
		midi.add( mtc )
		inputs = FixedTab("Inputs", self.sibi, Pt(28, 61), Pt(83, 74), Pt(25, 68), 0xe5e9eb, 0x98b7c7, ())
		inputs.add(OptionTable("Inputs", self.sibi, Pt(28,110), Pt(382,127), Pt(28,129), Pt(28,285),
		           Combo("", self.sibi, Pt(390,110), Pt(470,127), Pt(402,128), Pt(0,0), 0xe6e6e6 , 0xffffff, Pt(404,131), Pt(495,148), Pt(404,149), 0xe6e6e6, 0xffffff, ()),
				   ()))
		mtc.add(inputs)
		outputs = FixedTab("Outputs", self.sibi, Pt(89, 61), Pt(144, 74), Pt(88, 68), 0xe5e9eb, 0x98b7c7, ())
		outputs.add(OptionTable("Outputs", self.sibi, Pt(28,110), Pt(382,127), Pt(28,129), Pt(28,285),
		           Combo("", self.sibi, Pt(390,110), Pt(470,127), Pt(402,128), Pt(0,0), 0xe6e6e6 , 0xffffff, Pt(404,131), Pt(495,148), Pt(404,149), 0xe6e6e6, 0xffffff, ()),
				   ()))
		mtc.add(outputs)
		tc.add(midi)

		self.sibi.add( PushBtn("Cancel", self.sibi, Pt(334, 314, MOVE), Pt(399, 331, MOVE), None, None, ("silent_action",) ) )
		self.sibi.add( PushBtn("Ok", self.sibi, Pt(421, 314, MOVE), Pt(484, 332, MOVE), None, None, ("silent_action",) ) )


class Absynth5DemoDialog(SIBINVDA):

	def _get_name(self):
		if self.role == controlTypes.ROLE_UNKNOWN:
			return "\t" # I DO NOT WANT to speak anything there, but how???
		return super(Absynth5DemoDialog,self).name

	def defineSIBI(self):
		if self.hasKnownSIBI():
			return
		self.sibi = SIBI(self.windowHandle, 520, 350)
		self.sibi.nvda_class = Absynth5DemoDialog
		self.sibi.add( PushBtn("Run demo", self.sibi, Pt(35, 210), Pt(150, 225), None, None, ("silent_action",) ) )
		self.sibi.add( PushBtn("Buy", self.sibi, Pt(195, 210), Pt(322, 225), None, None, () ) )
		self.sibi.add( PushBtn("Activate", self.sibi, Pt(365, 210), Pt(482, 226), None, None, () ) )

class Absynth5(SIBINVDA):

	def _get_name(self):
		if self.role == controlTypes.ROLE_UNKNOWN:
			return "Absynth 5"  # We know it is from NI, no reason to remind us every time
		return super(Absynth5,self).name # For menus
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 764, 687)
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
								(0x093f42,0x2eafb1), (0x137072,0x115052), set(("click_to_focus", "pass_enter"))) )


class AppModule(appModuleHandler.AppModule):

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		# Guitar Rig create "logo" first, with the same class but without text
		if "NINormalWindow" in obj.windowClassName and obj.role == controlTypes.ROLE_PANE and obj.windowText !="":
			clsList.insert(0, Absynth5)
		elif "Absynth 5" in obj.windowText and obj.role == controlTypes.ROLE_DIALOG:
			clsList.insert(0, Absynth5DemoDialog)
		elif obj.windowText == "Audio and MIDI Settings" and obj.role == controlTypes.ROLE_DIALOG:
			clsList.insert(0, Absynth5AudioMidi)
		elif obj.role == controlTypes.ROLE_PANE:
			chooseKnownOverlay(obj, clsList)
			#log.error("overlay not found " + str(obj.devInfo))
