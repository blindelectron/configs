# Single Image Blob Interface Accessible Control
#
# SIBIAC based NVDA Application Module
# REAPER
#
# AZ (www.azslow.com), 2018-2020
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
import winUser
import api
import ctypes

from . import sibiac
from .sibiac import MoveFocusTo, SIBI, SIBINVDA, XY, Box, TextBox, TextOutBox, ScrollV, Color2Tuple, ColorMatcherObj, ColorMatcher, FindInXRight, FindInYDown, FindRow, FindVRange, FindHRange, chooseKnownOverlay, MouseSlowLeftClick, MouseScroll, FindNearestColor, Control, VList, YRange, YBar, Container, Label, ScrollLabel, PopupLabel, PushBtn, SwitchBtn, FixedTab, FixedTabControl, Combo, Dialog, OpenBtn, PopupMenuButton, Clickable, OptionTable, Pt, FIXED, MOVE, PROPORTIONAL, gSIBI
from .sibiac import SpinLabel

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

from .sibiac import focusNextTopWindowFor

from ctypes import *
import ctypes
from ctypes.wintypes import *
import winKernel
from locationHelper import RectLTRB

#Window messages
LVM_FIRST=0x1000
LVM_GETITEMW=LVM_FIRST+75
LVM_GETITEMSTATE=LVM_FIRST+44
LVM_GETFOCUSEDGROUP=LVM_FIRST+93
LVM_GETITEMCOUNT=LVM_FIRST+4
LVM_GETITEM=LVM_FIRST+75
LVN_GETDISPINFO=0xFFFFFF4F
LVM_GETITEMTEXTW=LVM_FIRST+115
LVM_GETHEADER=LVM_FIRST+31
LVM_GETCOLUMNORDERARRAY=LVM_FIRST+59
LVM_GETCOLUMNW=LVM_FIRST+95
LVM_GETSELECTEDCOUNT =(LVM_FIRST+50)
LVNI_SELECTED =2
LVM_GETNEXTITEM =(LVM_FIRST+12)
LVM_GETVIEW=LVM_FIRST+143
LV_VIEW_DETAILS=0x0001
LVM_GETSUBITEMRECT=LVM_FIRST+56
LV_VIEW_TILE=0x0004
    
#item mask flags
LVIF_TEXT=0x01
LVIF_IMAGE=0x02
LVIF_PARAM=0x04
LVIF_STATE=0x08
LVIF_INDENT=0x10
LVIF_GROUPID=0x100
LVIF_COLUMNS=0x200

#GETSUBITEMRECT flags
# Returns the bounding rectangle of the entire item, including the icon and label
LVIR_BOUNDS = 0
# Returns the bounding rectangle of the icon or small icon.
LVIR_ICON = 1
# Returns the bounding rectangle of the entire item, including the icon and label.
# This is identical to LVIR_BOUNDS.
LVIR_LABEL = 2



class FXEnabled:
	"""
	Plug-ins in question can not get focus otherwise, so force
	that from "Enabled" check box
	"""
	name = "Enabled"

	def script_focusNext(self, gesture):
		try:
			fx = self.parent.parent.getChild(7)
			if fx.childCount:
				fx = fx.children[0]
			if isinstance(fx, SIBINVDA):
				MoveFocusTo(fx.windowHandle)
				return
		except:
			pass
		gesture.send()

	__gestures = {
		"kb:tab": "focusNext",
	}
	
class ReaFXDialog(object):

	def _isFXChain(self):
		return self.windowText.startswith("FX");

	def _getFXList(self):
		try:
			if self.children[3].role == controlTypes.ROLE_LIST:
				return self.children[3]
			return self.children[4]
		except:
			log.info("FX List not found")
			return None
		
	def _afterFocus(self):
		"""
		Move focus to reasonable child
		"""
		focus = api.getFocusObject()
		if self == focus:
			if self._isFXChain():
				fxlist = self._getFXList()
				if fxlist:
					MoveFocusTo(self._getFXList().windowHandle);
				elif self.children:
					MoveFocusTo(self.children[0].windowHandle);
			elif self.children:
				MoveFocusTo(self.children[0].windowHandle);


	def getFXName(self):
		if self._isFXChain():
			fxlist = self._getFXList()
			if fxlist:
				for item in fxlist.children:
					if controlTypes.STATE_SELECTED in item.states:
						return item.name
			return ""
		return self.windowText
				
	def event_gainFocus(self):
		""" With some plug-ins, the focus is left on dialog... and nothing works (REAPER bug?)
		"""
		super(ReaFXDialog,self).event_gainFocus()
		# give some time, then check (beter way?)
		core.callLater(0.5, self._afterFocus)

gReaFadeInShapes = ( 
	(7, "Fast 2"),
	(12, "Fast 1"),
	(17, "Linear"),
	(18, "Smooth 1"), 
	(21, "Slow 1"), 
	(22, "Smooth 2"), 
	(27, "Slow 2"), 
)		

gReaFadeOutShapes = (  
	(8, "Slow 2"), 
	(13, "Slow 1"), 
	(17, "Linear"),
	(19, "Smooth 1"), 
	(21, "Fast 1"),
	(22, "Smooth 2"), 
	(27, "Fast 2"),
)		

gReaCrossFadeShapes = (
	(4, "Slow to fast 2"),
	(5, "Smooth 2"),
	(6, "Slow to fast 1"),
	(9, "Smooth 1"),
	(10, "Linear"),
	(15, "Fast to slow 1"),
	(21, "Fast to slow 2"),
)

class ReaFadeInShapeBtn(Window):
	def _get_value(self):
		i = FindInYDown(self.windowHandle, 20, 5, (0xb5b5b5, 0x000000), (0xf0f0f0,))
		for type in gReaFadeInShapes:
			if i <= type[0]:
				return type[1]
		return ""

class ReaFadeInForItemShapeBtn(ReaFadeInShapeBtn):
	name = "Fade in shape"
				
class ReaFadeOutForItemShapeBtn(Window):
	name = "Fade out shape"

	def _get_value(self):
		i = FindInYDown(self.windowHandle, 38, 5, (0xb5b5b5, 0x000000), (0xf0f0f0,))
		for type in gReaFadeOutShapes:
			if i <= type[0]:
				return type[1]
		return ""

class ReaCrossFadeShapeBtn(Window):
	def _get_value(self):
		i = FindInYDown(self.windowHandle, 24, 6, (0x888888, 0x000000), (0xf0f0f0, 0xb5b5b5))
		# return "%d" % i
		for type in gReaCrossFadeShapes:
			if i <= type[0]:
				return type[1]
		return ""

class ReaFadeShapeColorMatcherObj(ColorMatcherObj):
	def __init__(self, colors, name, items):
		super(ReaFadeShapeColorMatcherObj,self).__init__(colors, name)
		self.items = items
		
gReaFadeShapeMenuType = (
	ReaFadeShapeColorMatcherObj( (0xb5b5b5, 0xf0f0f0), "Fade in", ("Linear", "Fast 1", "Slow 1", "Fast 2", "Slow 2", "Smooth 1", "Smooth 2") ),
	ReaFadeShapeColorMatcherObj( (0xf0f0f0, 0xb5b5b5), "Fade out", ("Linear", "Slow 1", "Fast 1", "Slow 2", "Fase 2", "Smooth 1", "Smooth 2") ),
	ReaFadeShapeColorMatcherObj( (0xb5b5b5, 0xb5b5b5), "Cross fade", ("Liner", "Slow to fast 1", "Fast to slow 1", "Slow to fast 2", "Fast to slow 2", "Smooth 1", "Smooth 2" ) )
)
		
class ReaFadeShapeMenuItem(Window):
	
	def initOverlayClass(self):
		"""
		We have to find what this item really is and set its name
		"""
		# first detect In/Out/Cross
		obj = ColorMatcher(self.windowHandle, (XY(39,13), XY(74,13)), gReaFadeShapeMenuType).match()
		if obj is None:
			self.name = "Not matched"
			return
		# use ChildID as an index
		if len(obj.items) >= self.IAccessibleChildID:
			self.name = obj.items[self.IAccessibleChildID - 1]

class ReaTransient(object):
	""" can be used to workaround navigation, if I find how... """
	pass


# from sysListView32.py, NOTE: can be internal implementation dependent

from NVDAObjects.IAccessible import sysListView32
from ctypes import *
import watchdog

class ReaList(sysListView32.List):
	""" A workaround for OSError: [WindowError 8]
		This workaround tries to avoid virtualAllocEx in single column lists
	"""

	def _get__columnOrderArray(self):
		#log.info("ReaList: columnOrderArray")
		ccount = self.columnCount
		if ccount < 2:
			coa=(c_int *ccount)()
			coa[0] = 0
			return coa
		return super(ReaList, self)._get__columnOrderArray()


	def _getColumnLocationRaw(self,index):
		#log.info("L:getColumnLocationRaw")
		return super(ReaList, self)._getColumnLocationRay(index)
		
	def _getColumnContent(self, column):
		#log.info("L:getColumnContent")
		return super(ReaList, self)._getColumnContent(column)
	
	def _getColumnImageIDRaw(self, index):
		#log.info("L:getColumnImageIDRaw")
		return super(ReaList, self)._getColumnImageIDRaw(index)
	
	def _getColumnHeaderRaw(self, index):
		#log.info("L:getColumnHeaderRaw")
		return super(ReaList, self)._getColumnHeaderRaw(index)
	
	def _get_name(self):
		#log.info("L:getName");
		return super(ReaList, self)._get_name()
		
MPM_LVITEMTEXT = 0x8001
MPM_PROXY = 0x8002

class MP_LVItemText(Structure):
	_fields_=[
		('iItem',c_int),
		('iSubItem',c_int),
		('cchTextMax',c_int),
		('pad1',c_int),
		('hWnd',c_ulonglong),
		('pszText',c_ulonglong),
	]

class MPM_Msg(Structure):
  _fields_=[
	('hWnd',c_ulonglong),
	('uMsg',c_ulonglong),
	('wParam',c_ulonglong),
	('lParam',c_ulonglong)
  ]

class MProxy(object):
	hWnd = None
	hProcess = None
	
	def __init__(self):
		self.hWnd = winUser.user32.FindWindowW("SIBMP", None)
		if self.hWnd:
			id = winUser.getWindowThreadProcessID(self.hWnd)[0]
			if id:
				self.hProcess = winKernel.openProcess(winKernel.PROCESS_ALL_ACCESS, False, id)
	
	def __del__(self):
		if self.hProcess:
			winKernel.closeHandle(self.hProcess)

	def dupTo(self, data):
		proxyData = winKernel.virtualAllocEx(self.hProcess,None,sizeof(data),winKernel.MEM_COMMIT,winKernel.PAGE_READWRITE)
		try:
			winKernel.writeProcessMemory(self.hProcess, proxyData, byref(data), sizeof(data), None)
		except:
			winKernel.virtualFreeEx(self.hProcess, proxyData, 0, winKernel.MEM_RELEASE)
			proxyData = None
		return proxyData
	
	def copyFrom(self, data, proxyData):
		if proxyData:
			winKernel.readProcessMemory(self.hProcess, proxyData, byref(data), sizeof(data), None)

	def virtualFree(self, proxyData):
		if proxyData:
			winKernel.virtualFreeEx(self.hProcess, proxyData, 0, winKernel.MEM_RELEASE)
	
	def sendMessage(self, msg):
		proxyMsg = self.dupTo(msg)
		if not proxyMsg:
			return 0
		ret = 0
		try:
			ret = watchdog.cancellableSendMessage(self.hWnd, MPM_PROXY, 0, proxyMsg)
		finally:
			self.virtualFree(proxyMsg)
		return ret
			
class ReaListItem(sysListView32.ListItem):
	""" A workaround for OSError: [WindowError 8]
		It can only work stable when all methods using virtualAllocEx goes throw proxy,
		so it should be in sync with used NVDA version
	"""
	mproxy = None
	
	def initOverlayClass(self):
		if self.appModule.is64BitProcess:
			self.mproxy = MProxy()
			if not self.mproxy.hProcess:
				self.mproxy = None
	
	def _getColumnLocationRaw(self,index):
		# log.info("getColumnLocationRaw")
		if not self.mproxy:
			return super(ReaListItem, self)._getColumnLocationRaw(index)
		# log.info("using proxy")
		localRect = RECT(left=LVIR_LABEL, top=index)
		proxyRect = self.mproxy.dupTo(localRect)
		res = 0
		try:
			if proxyRect:
				res = self.mproxy.sendMessage(MPM_Msg(hWnd = self.windowHandle, uMsg = LVM_GETSUBITEMRECT, wParam = self.IAccessibleChildID-1, lParam = proxyRect))
			if res:
				self.mproxy.copyFrom(localRect, proxyRect)
		finally:
			self.mproxy.virtualFree(proxyRect)
		if res == 0:
			return None
		left = localRect.left
		top = localRect.top
		right = localRect.right
		bottom = localRect.bottom
		if left > right:
			left = right
		if top > bottom:
			top = bottom
		# log.info("we got something prom proxy: %d %d %d %d" % (left, top, right, bottom))
		return RectLTRB(left, top, right, bottom).toScreen(self.windowHandle).toLTWH()

			
	def _getColumnContent(self, column):
		# log.info("getColumnContent")
		return super(ReaListItem, self)._getColumnContent(column)
	
	def _getColumnImageIDRaw(self, index):
		# log.info("getColumnImageIDRaw")
		return super(ReaListItem, self)._getColumnImageIDRaw(index)
	
	def _getColumnHeaderRaw(self, index):
		# log.info("getColumnHeaderRaw")
		return super(ReaListItem, self)._getColumnHeaderRaw(index)
	
	def _get_name(self):
		# log.info("getName");
		return super(ReaListItem, self)._get_name()
	# def _get_name(self):
		# name = ""
		# try:
			# name = super(ReaListItem, self)._get_name()
		# except:
			# pass
		# if name:
			# return name
		# hProxyWnd = winUser.user32.FindWindowW("SIBMP", None)
		# if not hProxyWnd:
			# return name
		# buffer=create_unicode_buffer(512)
		# item = MP_LVItemText(iItem=self.IAccessibleChildID-1, iSubItem=0, cchTextMax=512, hWnd = self.windowHandle, pszText= addressof(buffer))
		# try:
			# len = watchdog.cancellableSendMessage(hProxyWnd, MPM_LVITEMTEXT, windll.kernel32.GetCurrentProcessId(), byref(item))
			# if len:
				# name = buffer.value
		# except:
			# pass
		# return name

	
# remember all windows
gKnownHandles = {}

def clearKnownCache(cache):
	for hwnd in list(cache):
		if not winUser.user32.IsWindow(hwnd):
			del cache[hwnd]

gDialogs = { "Program Browser": SIProgramBrowser, "Transient Detection Settings": ReaTransient, }
			
gCWPlugs = { "VSTi: SI-Electric": SIElectricPiano, "VSTi: SI-Drum": SIDrumKit, "VSTi: SI-Bass": SIBassGuitar, "VSTi: SI-String": SIStringSection, "VSTi: Dimension Pro": DimensionPro, "VSTi: SessionDrummer": SessionDrummer }

gWavesPlugs = { "VST3: GTR Stomp 2": GTRStomp2, "VST3: GTR Stomp 4": GTRStomp4, "VST3: GTR Stomp 6": GTRStomp6, "VST3: GTR Tool Rack": GTRRack }

gReaFadeButtons = { 1056: ReaFadeInForItemShapeBtn, 1057: ReaFadeOutForItemShapeBtn, 1127: ReaFadeInShapeBtn, 1128: ReaCrossFadeShapeBtn } 


gJUCEPlugs = { "VSTi: Addictive Drums 2": AD2, "VSTi: Addictive Keys": AK, "VSTi: Pure Synth Platinum": PSP, "VSTi: VSCO2": VSCO2, "VST3: STL Tonality - Ho": SHB }
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

gOtherPlugs1 = { "VSTi: Zampler": Zampler, "VST: EZmix": EZMix, "VSTi: EZdrummer": EZDrummer, "VSTi: sforzando": Sforzando, "VSTi: SynthMaster" : { "Player": SMP } }

class AppModule(appModuleHandler.AppModule):

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		"""
		Important that we get 2 (!) objects per window, one with ROLE_WINDOW and another with ROLE_XXX
		I am not sure what is better... but lets ignore ROLE_WINDOW
		"""		
		#if obj.role == controlTypes.ROLE_WINDOW:
		#	return
		# DIALOG, PANE, BUTTON only...
		if obj.windowClassName == "Button":
			if obj.windowText != "": # accessible by itself
				return
			if obj.windowControlID in gReaFadeButtons:
				clsList.insert(0, gReaFadeButtons[obj.windowControlID])
				return
		elif obj.windowClassName == "#32768": # Context menu
			if obj.role == controlTypes.ROLE_MENUITEM and obj.displayText == "": # not accessible menu items
				# (Cross)Fade shapes menu is 118x182, items 118x26
				if obj.location.width == 118 and obj.location.height == 26:
					clsList.insert(0, ReaFadeShapeMenuItem)
		elif obj.windowClassName.startswith("JUCE_"): # JUCE dialogs, the parent is not plug-in...
			if obj.windowText != "" and obj.windowText in gJUCEDialogs:
				gKnownHandles[obj.windowHandle] = gJUCEDialogs[obj.windowText]["cls"]
		elif obj.windowClassName in ("Static", "ComboBox"):
			return
		elif obj.windowClassName == "SysListView32":
			if obj.role == controlTypes.ROLE_LIST:
				clsList.insert(0, ReaList)
			elif obj.role == controlTypes.ROLE_LISTITEM:
				clsList.insert(0, ReaListItem)
			return
		if obj.role == controlTypes.ROLE_CHECKBOX:
			if obj.windowHandle not in gKnownHandles:
				gKnownHandles[obj.windowHandle] = False
				# detect enabled check box in the FX window
				try:
					dlg = obj.parent.parent
					if dlg.role == controlTypes.ROLE_PROPERTYPAGE and dlg.getChild(6) == obj: # and dlg.getChild(7).role == controlTypes.ROLE_UNKNOWN:
						gKnownHandles[obj.windowHandle] = FXEnabled
						clearKnownCache(gKnownHandles)
				except:
					pass
		elif obj.role == controlTypes.ROLE_DIALOG:
			if obj.windowHandle not in gKnownHandles:
				if obj.windowClassName == "#32770":
					if (obj.windowText.startswith("FX:") or obj.windowText.startswith("VST")):
						gKnownHandles[obj.windowHandle] = ReaFXDialog
						clearKnownCache(gKnownHandles)
					elif obj.windowText in gDialogs:
						gKnownHandles[obj.windowHandle] = gDialogs[obj.windowText]
						clearKnownCache(gKnownHandles)
					elif obj.windowText != "":
						gKnownHandles[obj.windowHandle] = False # empty window text can be transient, but once set we are sure what it is
					else:
						return # we are not yet sure what it is
				else:
					gKnownHandles[obj.windowHandle] = False # some unknown dialog, not interesting
		elif obj.role == controlTypes.ROLE_PANE:
			if obj.windowHandle not in gKnownHandles:
				# gKnownHandles[obj.windowHandle] = False  # do NOT cache "not found" till absolutely sure
				try:
					prt = obj.parent
					while prt:
						if prt.role == controlTypes.ROLE_DIALOG:
							break
						prt = prt.parent
				except:
					prt = None
				if  prt and isinstance(prt, ReaFXDialog):
					if obj.windowClassName.startswith("REAPER"):
						gKnownHandles[obj.windowHandle] = False  # many REAPER elements are panes
						return
					fxname = prt.getFXName()
					if fxname == "":
						return # a bug in NVDA prevents list traversal, hope it is transient
					if "JUCE_" in obj.windowClassName:
						for key, cls in gJUCEPlugs.items():
							if fxname.startswith(key):
								gKnownHandles[obj.windowHandle] = cls
								break
					elif "NIVSTChildWindow" in obj.windowClassName:
						if "VST: Guitar Rig" in fxname :
							gKnownHandles[obj.windowHandle] = GuitarRig5
						elif "VSTi: Absynth 5" in fxname:
							gKnownHandles[obj.windowHandle] = Absynth5
					elif obj.windowText == "External VST Window":
						for key in gCWPlugs:
							if key in fxname:
								gKnownHandles[obj.windowHandle] = gCWPlugs[key]
								break
					elif obj.windowClassName.startswith("WaveShell"):
						for key in gWavesPlugs:
							if fxname.startswith(key):
								gKnownHandles[obj.windowHandle] = gWavesPlugs[key]
								break
					elif obj.windowClassName.startswith("GWinClass_"):
						if fxname.startswith("VST: GTune"):
							gKnownHandles[obj.windowHandle] = GTune
					elif obj.windowClassName.startswith("VSTGUI"):
						for key in gVSTGUIPlugs:
							if fxname.startswith(key):
								gKnownHandles[obj.windowHandle] = gVSTGUIPlugs[key]
								break
					elif obj.windowClassName.startswith("Plugin"):
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
					if obj.windowHandle in gKnownHandles:
						clearKnownCache(gKnownHandles)
					else:
						log.info("Unknown pane: %d %d %s %s" % (obj.windowHandle, obj.role, obj.windowClassName, obj.windowText))
						return # we do not know what it is, at least at the moment
				elif prt and prt.windowText != "":
					gKnownHandles[obj.windowHandle] = False # not in ReaFXDialog (for sure) == not plug-in
				else:
					return # ReaFXDialog sometimes delay window text set and so is not detected in time
		else:
			return # so far there was no evidence other roles are interesting, no reason to cache
		
		if gKnownHandles[obj.windowHandle]:
			clsList.insert(0, gKnownHandles[obj.windowHandle])
			
			

	def script_focusNextTopWindow(self, gesture):
		focusNextTopWindowFor(self.processID, ["RPTopmostButton",])

	__gestures = {
		"kb:NVDA+control+tab": "focusNextTopWindow"
	}
	

	def __init__(self, *args, **kwargs):
		super(AppModule, self).__init__(*args, **kwargs)
		#log.info("REAPER loaded")
	
	def terminate(self):
		pass
		#log.info("REAPER terminated")
