# Single Image Blob Interface Accessible Control
#
# SIBIAC based NVDA Application Module
# REAPER 64bit bridge
#
# AZ (www.azslow.com), 2018 - 2020
from __future__ import absolute_import
from __future__ import division
from __future__ import print_function


import appModuleHandler
from NVDAObjects.window import Window
from logHandler import log
import winUser
import core
import api

from .sibiac import MoveFocusTo

from .sibiac.sforzando import Sforzando
from .sibiac.vsco2 import VSCO2
from .sibiac.smp import SMP
from .sibiac.xln import AD2, AK
from .sibiac.ni import GuitarRig5, Absynth5
from .sibiac.cw import SIElectricPiano, SIDrumKit, SIBassGuitar, SIStringSection, DimensionPro, SessionDrummer, SIProgramBrowser
from .sibiac.gtr import GTRStomp2, GTRStomp4, GTRStomp6, GTRRack
from .sibiac.gtune import GTune
from .sibiac.zampler import Zampler
from .sibiac.psp import PSP
from .sibiac.ezmix import EZMix
from .sibiac.igniteampsnadir import IgniteAmpsNadIR
from .sibiac.ezdrummer import EZDrummer
from .sibiac.stlthb    import SHB, SHB_Dialog

class ReaBridge(Window):
	def _afterFocus(self):
		"""
		Move focus to reasonable child
		"""
		focus = api.getFocusObject()
		if self == focus:
			try:
				child = self.getChild(0)
				MoveFocusTo(child.windowHandle)
			except:
				pass

	def event_gainFocus(self):
		""" Transfer focus to child """
		super(ReaBridge,self).event_gainFocus()
		# give some time, then check (beter way?)
		core.callLater(0.5, self._afterFocus)


gCWPlugs = {
  "SI-Electric": SIElectricPiano, "SI-Drum": SIDrumKit, "VSTi: SI-Bass": SIBassGuitar, "SI-String": SIStringSection, "Dimension Pro": DimensionPro, "SessionDrummer": SessionDrummer 
}

gWavesPlugs = { 
	"GTR Stomp 2": GTRStomp2, "GTR Stomp 4": GTRStomp4, "GTR Stomp 6": GTRStomp6, "GTR Tool Rack": GTRRack 
}

gJUCEPlugs = { "Addictive Drums 2": AD2, "Addictive Keys": AK, "VSCO2": VSCO2, "Pure Synth Platinum": PSP, "VST3: STL Tonality - Ho": SHB }
gJUCEDialogs = {
	"Edit preset name": { "cls": SHB_Dialog, "size": (350, 199) },
	"Create preset": { "cls": SHB_Dialog, "size": (394, 149) },
	"Remove preset": { "cls": SHB_Dialog, "size": (394, 149) },
	"Error": { "cls": SHB_Dialog, "size": (463, 165) },
	"Edit bank name": { "cls": SHB_Dialog, "size": (350, 199) },
	"Create new Bank": { "cls": SHB_Dialog, "size": (394, 149) },
	"Remove bank": { "cls": SHB_Dialog, "size": (394, 149) },
}

gVSTGUIPlugs = { "VST: NadIR": IgniteAmpsNadIR }

gOtherPlugs1 = { "sforzando": Sforzando, "SynthMaster" : { "Player": SMP }, "Zampler": Zampler, "EZmix": EZMix, "EZdrummer": EZDrummer }


gKnownHandles = {}
def clearKnownCache(cache):
	for hwnd in list(cache):
		if not winUser.user32.IsWindow(hwnd):
			del cache[hwnd]

class AppModule(appModuleHandler.AppModule):

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		if obj.windowClassName.startswith("JUCE_"): # JUCE dialogs, the parent is not plug-in...
			if obj.windowText != "" and obj.windowText in gJUCEDialogs:
				gKnownHandles[obj.windowHandle] = gJUCEDialogs[obj.windowText]["cls"]
		if obj.role not in (3, 4, 5):
			return
		if obj.windowHandle not in gKnownHandles:
			gKnownHandles[obj.windowHandle] = False
			if obj.windowClassName == "REAPERb32host":
				gKnownHandles[obj.windowHandle] = ReaBridge
			elif obj.windowClassName == "#32770" and obj.windowText == "Program Browser":
				gKnownHandles[obj.windowHandle] = SIProgramBrowser
			else:
				try:
					prt = obj.parent
					while prt:
						if prt.windowClassName == "REAPERb32host":
							break
						prt = prt.parent
				except:
					prt = None
				if  prt:
					if "JUCE_" in obj.windowClassName:
						fxname = prt.windowText
						for key, cls in gJUCEPlugs.items():
							if fxname.startswith(key):
								gKnownHandles[obj.windowHandle] = cls
								break
					elif "NIVSTChildWindow" in obj.windowClassName:
						fxname = prt.windowText
						if "Guitar Rig" in fxname :
							gKnownHandles[obj.windowHandle] = GuitarRig5
						elif "Absynth 5" in fxname:
							gKnownHandles[obj.windowHandle] = Absynth5
					elif obj.windowText == "External VST Window":
						fxname = prt.windowText
						for key in gCWPlugs:
							if key in fxname:
								gKnownHandles[obj.windowHandle] = gCWPlugs[key]
								break
					elif obj.windowClassName.startswith("WaveShell"):
						fxname = prt.windowText
						for key in gWavesPlugs:
							if fxname.startswith(key):
								gKnownHandles[obj.windowHandle] = gWavesPlugs[key]
								break								
					elif obj.windowClassName.startswith("GWinClass_"):
						fxname = prt.windowText
						if fxname.startswith("GTune"):
							gKnownHandles[obj.windowHandle] = GTune
					elif obj.windowClassName.startswith("VSTGUI"):
						fxname = prt.getFXName()
						for key in gVSTGUIPlugs:
							if fxname.startswith(key):
								gKnownHandles[obj.windowHandle] = gVSTGUIPlugs[key]
								break
					elif obj.windowClassName.startswith("Plugin"):
						fxname = prt.windowText
						for key, cls in gOtherPlugs1.items():
							if fxname.startswith(key):
								if isinstance(cls, dict):
									found = False
									for key2, cls2 in cls.items():
										if fxname.find(key2) > 0:
											gKnownHandles[obj.windowHandle] = cls2
											found = True
											break
									if found:
										break
								else:
									gKnownHandles[obj.windowHandle] = cls
									break
						
			if gKnownHandles[obj.windowHandle]:
				clearKnownCache(gKnownHandles)
		if gKnownHandles[obj.windowHandle]:
			clsList.insert(0, gKnownHandles[obj.windowHandle])
