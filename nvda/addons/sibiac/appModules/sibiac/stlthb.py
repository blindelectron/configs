# Single Image Blob Interface Accessible Control
# 
# STL Tonality Howard Benson overlay
#
# AZ (www.azslow.com), 2018 - 2020
from ctypes import *
from ctypes.wintypes import *
import time
from logHandler import log

from . import *

class SHB_List(Control):
	def __init__(self, name, sibi, check_pt, opt = None):
		super(SHB_List,self).__init__(name, sibi, opt)
		self.check_pt = sibi.ptScale(check_pt)
		self.item = sibi.TextBox(Pt(16,3), Pt(0,16), self._getItemShift)
		self.item_h = sibi.yScale(16)
		self.cur_mark = sibi.ptScale(Pt(12,3))
		self.box = None
		self.idx = 0
		self.n = 0
	
	def _getItemShift(self):
		return (0, self.idx * self.item_h)
	
	def isFocusable(self):
		hwnd = WindowFromPoint(self.sibi.hwnd, self.check_pt.x, self.check_pt.y)
		if hwnd == self.sibi.hwnd:
			self.box = None
			self.n = 0
			self.idx = 0
			return False
		if self.box is not None and self.box.hwnd == hwnd:
			return True
		# recalculate 3 + 16*i (+13 inclusive gives hightlited box)
		self.box = GetClientRect(hwnd)
		self.item.hwnd = hwnd
		self.item.right = self.box.right - 4
		self.n = self.box.bottom // self.item_h # +4 pixels for border
		# find current by check symbol
		dy = FindInYDown(hwnd, self.cur_mark.x, self.cur_mark.y, (0x323e44, 0x181f22), (0xf8f8f9,) )
		if dy > 0:
			self.idx = (dy + self.cur_mark.y) // self.item_h
		return True
	
	def getTextInfo(self):
		if not self.isFocusable():
			text = "Closed"
		else:
			text = self.item.getText()
		return (self.type, self.name, text)
	
	def onEscape(self):
		self.speakFocusAfter()
		return False
		
	def onUp(self):
		if self.idx > 0:
			self.idx -= 1
		self.speakAfter(0)
		return True
		
	def onDown(self):
		if self.idx < self.n - 1:
			self.idx += 1
		self.speakAfter(0)
		return True
	
	def onEnter(self):
		self.item.getBox().leftClick()
		self.speakFocusAfter()
		return True
		
	def onHelp(self, prefix = "", suffix = ""):
		self.speak(prefix + "Use up and down arrows to select and enter to load. Press escape to close the list without selection." + suffix)
		return True

class SHB_TopBtn(MXY):
	def __init__(self, sibi, pt, name, dialog = False):
		xy = sibi.ptScale(pt)
		super(SHB_TopBtn,self).__init__(sibi.hwnd, xy.x, xy.y)
		self.name = name
		self.dialog = dialog
		
	def isEnabled(self):
		return self.FindNearestColor((0x696969, 0xffffff)) == 1

class SHB_Knob(Control):
	def __init__(self, name, sibi, ref_pt, getShift = None):
		super(SHB_Knob,self).__init__(name, sibi, None)
		self.type = "Knob"
		self.label = sibi.TextBox(Pt(ref_pt.x, ref_pt.y), Pt(ref_pt.x + 45, ref_pt.y + 13), getShift)		
		self.ctl_pt = sibi.MXY(Pt(ref_pt.x + 23, ref_pt.y - 37), getShift)
		self.drag = 0 # auto delay and aggregate shift+arrows

	def isFocusable(self):
		return False
		
	def reactionTime(self):
		return 500 # we can produce unexpected double clicks otherwise
	
	def getValue(self):
		text = "unknown"
		label = self.label.getBox()
		if self.drag:
			self.ctl_pt.leftDrag(0, self.drag, 0.1, False)
			self.drag = 0
		else:
			self.ctl_pt.leftDown()
		try:
			time.sleep(0.1)
			text = label.getText()
		except:
			pass
		self.ctl_pt.leftUp()
		return text
		
	def getTextInfo(self):
		return (self.type, self.name, self.getValue())
	
	def getNameInGroup(self):
		return (self.name, self.getValue())
	
	def onUp(self):
		self.ctl_pt.mouseScroll(120)
		self.speakAfter()
		return True

	def onDown(self):
		self.ctl_pt.mouseScroll(-120)
		self.speakAfter()
		return True
	
	def onShiftUp(self):
		self.drag -= 1 
		self.speakAfter()
		return True

	def onShiftDown(self):
		self.drag += 1 
		self.speakAfter()
		return True
	
	def onEnter(self):
		self.ctl_pt.leftDblClick()
		self.speakAfter()
		return True		
	
	def onHelp(self, prefix = "", suffix = ""):
		self.speak(prefix + "Use up and down arrows to change the value. Do this with Shift for fine adjustment. Press enter to reset the value to default." + suffix)

class SHB_Knob2(SHB_Knob):
	__text_replace = { "1116" : "1/16", "118": "1/8", "1I32": "1/32", "114" : "1/4", "112" : "1/2", "20" : "2 D" }

	def __init__(self, name, sibi, ref_pt, getShift = None):
		super(SHB_Knob2,self).__init__(name, sibi, ref_pt, getShift)
		self.ctl_pt = sibi.MXY(Pt(ref_pt.x + 23, ref_pt.y + 37), getShift)	
	
	def getValue(self):
		text = super(SHB_Knob2,self).getValue()
		if text in self.__text_replace:
			text = self.__text_replace[text]
		return text
	
class SHB_Switch(Control):
	def __init__(self, name, sibi, ctl_pt, colors, states, getShift = None):
		super(SHB_Switch,self).__init__(name, sibi, None)
		self.type = "Switch"
		self.ctl_pt = sibi.MXY(ctl_pt, getShift)
		self.colors = Color2Tuple(colors)
		self.states = Arg2Tuple(states)
		
	def getState(self):
		i = self.ctl_pt.FindNearestColor(self.colors)
		return self.states[i]
		
	def getTextInfo(self):
		return (self.type, self.name, self.getState())
	
	def getNameInGroup(self):
		return (self.name, self.getState())
	
	def onUp(self):
		return self.onEnter()

	def onDown(self):
		return self.onEnter()
		
	def onEnter(self):
		self.ctl_pt.leftClick()
		self.speakAfter()
		return True		
	
	def onHelp(self, prefix = "", suffix = ""):
		self.speak(prefix + "Use enter or up and down arrows to switch." + suffix)

class SHB_Selector(Control):
	def __init__(self, name, sibi, btns):
		super(SHB_Selector,self).__init__(name, sibi, ("slow_reaction",))
		self.btns = []
		for btn in btns:
			self.btns.append( (sibi.MXY(btn[0]), btn[1]) )
		self.last_selected = None
		
	def __isSelected(self, idx):
		pt = self.btns[idx][0]
		return pt.FindNearestColor((0x616167, 0x2b2b2f)) == 1
	
	def __select(self,idx):
		if idx is None:
			return False
		if idx < 0:
			idx = len(self.btns) - 1
		elif idx >= len(self.btns):
			idx = 0
		self.btns[idx][0].leftClick()
		return True
	
	def getSelectedIdx(self):
		if self.last_selected is not None and self.__isSelected(self.last_selected):
			return self.last_selected
		for idx in range(len(self.btns)):
			if self.__isSelected(idx):
				self.last_selected = idx
				return idx
		return None
	
	def getSelected(self):
		idx = self.getSelectedIdx()
		if idx is None:
			return "Unknown"
		return self.btns[idx][1]
	
	def getTextInfo(self):
		return (self.type, self.name, self.getSelected())
		
	def onDown(self):
		self.getSelected()
		self.__select(self.last_selected + 1)
		self.speakAfter()
		return True

	def onUp(self):
		self.getSelected()
		self.__select(self.last_selected - 1)
		self.speakAfter()
		return True
		
	def onHelp(self, prefix = "", suffix = ""):
		self.speak(prefix + "Select %s using up or down arrow." % self.name + suffix);
		
class SHB_Switch_InGroup(SHB_Switch):

	def isFocusable(self):
		return False

class SHB_Switch2_InGroup(SHB_Switch_InGroup):

	def __init__(self, name, sibi, ctl_pt1, ctl_pt2, colors, states, getShift = None):
		super(SHB_Switch2_InGroup,self).__init__(name, sibi, ctl_pt1, colors, states, getShift)
		self.ctl_pt2 = sibi.MXY(ctl_pt2)
		
	def onEnter(self):
		i = self.ctl_pt.FindNearestColor(self.colors)
		if i == 0:
			self.ctl_pt2.leftClick()
		else:
			self.ctl_pt.leftClick()
		self.speakAfter()
		return True
	
class SHB_ListBtn(Label):
	def __init__(self, name, sibi, lt, rb):
		super(SHB_ListBtn,self).__init__(name, sibi, lt, rb, ("dynamic_text",))

	def isFocusable(self):
		return False
		
	def getNameInGroup(self):
		return (self.name, self.box.getBox().getText())
		
	def onEnter(self):
		box = self.box.getBox()
		box.leftClick()
		self.speakFocusAfter()
		return True
		
	def onHelp(self, prefix = "", suffix = ""):
		self.speak(prefix + "Press encter to open routing options" + suffix)
		return True
		
class SHB_IOGroup(Group):
	def __init__(self, sibi):
		super(SHB_IOGroup,self).__init__("Signal controls", sibi, None)
		
		self.add( SHB_ListBtn("Routing", sibi, Pt(833, 5), Pt(919, 18)) )
		self.add( SHB_Knob("Input", sibi, Pt(65, 551)) )
		self.add( SHB_Knob("Output", sibi, Pt(859, 551)) )
		self.add( SHB_Switch_InGroup("Gate", sibi, Pt(182, 480), (0x5f5f66, 0x38383d), ("disabled", "enabled")) )
		self.add( SHB_Knob("Gate level", sibi, Pt(172, 563)) )

class SHB_Page(Container):
	def __init__(self, name, sibi, active_pt, power_pt):
		super(SHB_Page,self).__init__(name, sibi, None)
		self._active_pt = sibi.MXY(active_pt)
		self._active_colors = (0x191918, 0x886335)
		self._power_pt = sibi.MXY(power_pt)
		self._power_colors = (0x1d1d1e, 0xefc07e) # ("No", "Yes")
		
	def _getPower(self):
		if self._power_pt.FindNearestColor(self._power_colors) == 1:
			return "On"
		return "Off"
	
	def isActive(self):
		return self._active_pt.FindNearestColor(self._active_colors) == 1
		
	def activate(self):
		self._active_pt.leftClick()

	def onEnter(self):
		self._power_pt.leftClick()
		self.speakAfter()
		return True

	def onHelp(self, prefix = "", suffix = ""):
		return self.speak(prefix + "Press enter to toggle power." + suffix)
		
class SHB_StopmsPage(SHB_Page):
	def __init__(self, sibi):
		super(SHB_StopmsPage,self).__init__("Pedals", sibi, Pt(612, 514), Pt(611, 552))

		screamer = Group("Screamer", sibi, None)
		self.add(screamer)
		screamer.add( SHB_Switch_InGroup("Power", sibi, Pt(200, 309), (0x44464c, 0xffeb85), ("Off", "On")) )
		screamer.add( SHB_Knob2("Gain", sibi, Pt(178, 160)) )
		screamer.add( SHB_Knob2("Drive", sibi, Pt(123, 227)) )
		screamer.add( SHB_Knob2("Tone", sibi, Pt(229, 226)) )

		delay = Group("Delay", sibi, None)
		self.add(delay)
		delay.add( SHB_Switch_InGroup("Power", sibi, Pt(492, 309), (0x44464c, 0xffeb85), ("Off", "On")) )
		delay.add( SHB_Knob2("Level", sibi, Pt(469, 160)) )
		delay.add( SHB_Knob2("Feedback", sibi, Pt(520, 226)) )
		delay.add( SHB_Switch_InGroup("BPM sync", sibi, Pt(511, 160), (0x656565, 0xbecc3), ("Off", "On")) )
		delay.add( SHB_Knob2("Delay time", sibi, Pt(414, 226)) )	
		delay.add( SHB_Switch2_InGroup("", sibi, Pt(491, 253), Pt(491, 276), (0xf2b95e, 0x181818), ("Pre", "Post")) )

		delay = Group("Reverb", sibi, None)
		self.add(delay)
		delay.add( SHB_Switch_InGroup("Power", sibi, Pt(780, 309), (0x44464c, 0xffeb85), ("Off", "On")) )
		delay.add( SHB_Knob2("Decay", sibi, Pt(758, 160)) )
		delay.add( SHB_Knob2("Mix", sibi, Pt(703, 227)) )
		delay.add( SHB_Knob2("Pre delay", sibi, Pt(809, 226)) )
		delay.add( SHB_Switch2_InGroup("", sibi, Pt(779, 253), Pt(779, 276), (0xf2b95e, 0x181818), ("Pre", "Post")) )
		
	def getNameInGroup(self):
		return (self.name, self._getPower())

class SHB_FileSelect(Label):
	def __init__(self, name, sibi, lt, rb, select_pt):
		super(SHB_FileSelect,self).__init__(name, sibi, lt, rb, ("dynamic_text",))
		self.select_pt = sibi.MXY(select_pt)
		
	def onEnter(self):
		self.expectFocusReturn()
		self.select_pt.leftClick()
		return True
		
	def onHelp(self, prefix = "", suffix = ""):
		self.speak(prefix + "Press enter to select file" + suffix)
		return True

		
class SHB_CabsPage(SHB_Page):
	__cabtype_btns = [
		(Pt(42, 148), "CAB-1"),
		(Pt(42, 184), "CAB-2"),
		(Pt(42, 220), "CAB-3"),
		(Pt(42, 258), "CAB-4"),
		(Pt(42, 295), "CAB-5"),
		(Pt(42, 332), "EXT CAB"),
	]

	def __init__(self, sibi):
		super(SHB_CabsPage,self).__init__("Cabinets", sibi, Pt(751, 514), Pt(735, 552))
		self.add( SHB_Selector("Cabinet type", sibi, self.__cabtype_btns) )
		self.add( SHB_FileSelect("Load IR", sibi, Pt(34, 371), Pt(215, 390), Pt(128, 341)) )
		self.add( SHB_Switch("Lock cab", sibi, Pt(768, 551), (0x101010, 0xf0bf7e), ("Off", "On")) )
		
class SHB_AmpSelector(SHB_Page):
	__amptype_btns = [
		(Pt(410, 477), "AMP-1"),
		(Pt(453, 477), "AMP-2"),
		(Pt(498, 477), "AMP-3"),
		(Pt(430, 524), "AMP-4"),
		(Pt(476, 524), "AMP-5"),
	]

	def __init__(self, sibi):
		super(SHB_AmpSelector,self).__init__("Amplifier", sibi, Pt(682, 514), Pt(682, 552))
		self.type_selector = SHB_Selector("Amplifier type", sibi, self.__amptype_btns)
		self.add( SHB_AmpControls(self.type_selector) )
	
	def getNameInGroup(self):
		return (self.name, "%s, %s" % (self.type_selector.getSelected(), self._getPower()))
		
	def onUp(self):
		return self.type_selector.onUp()

	def onDown(self):
		return self.type_selector.onDown()
		
	def onHelp(self, prefix = "", suffix = ""):
		return self.type_selector.onHelp(prefix, " Press enter to toggle power." + suffix)

class SHB_AmpControls(GroupBase):
	__amp_defs_conf = [
		( # AMP-1
		  ( "Knob", "Bass", Pt(220, 381) ), ( "Knob", "Middle", Pt(298, 381) ), ( "Knob", "Treble", Pt(377, 381) ), ( "Knob", "Presence", Pt(460, 381) ), 
		  ( "Knob", "Contour", Pt(544, 381) ), ( "Knob", "Gain", Pt(625, 381) ), ("Knob", "Master", Pt(708, 381) ) 
		),
		( # AMP-2
		  ( "Knob", "Bass", Pt(266, 381) ), ( "Knob", "Mids", Pt(342, 381) ), ( "Knob", "Treble", Pt(423, 381) ), ( "Knob", "Presence", Pt(507, 381) ), 
		  ( "Knob", "Gain", Pt(585, 381) ), ("Knob", "Master", Pt(662, 381) ), ( "Switch", "Channel", Pt(750, 339), (0xf4e1c2, 0x0c0a07), ("Lead", "Clean") ) 
		),
		( # AMP-3
		  ( "Knob", "Bass", Pt(266, 381) ), ( "Knob", "Mids", Pt(342, 381) ), ( "Knob", "Treble", Pt(423, 381) ), ( "Knob", "Presence", Pt(507, 381) ), 
		  ( "Knob", "Gain", Pt(585, 381) ), ("Knob", "Master", Pt(662, 381) ), ( "Switch", "Mode", Pt(750, 339), (0xf4e1c2, 0x0c0a07), ("Bright", "Normal") ) 
		),
		( # AMP-4
		  ( "Knob2", "Bass", Pt(244, 256) ), ( "Knob2", "Mids", Pt(328, 256) ), ( "Knob2", "Treble", Pt(414, 256) ), ( "Knob2", "Presence", Pt(502, 256) ), 
		  ( "Knob2", "Gain", Pt(588, 256) ), ("Knob2", "Master", Pt(672, 256) ), ( "Switch", "Mode", Pt(798, 299), (0xf4e1c2, 0x0c0a07), ("Bright", "Normal") ) 
		),
		( # AMP-4
		  ( "Knob2", "Bass", Pt(244, 256) ), ( "Knob2", "Mids", Pt(328, 256) ), ( "Knob2", "Treble", Pt(414, 256) ), ( "Knob2", "Presence", Pt(502, 256) ), 
		  ( "Knob2", "Gain", Pt(588, 256) ), ("Knob2", "Master", Pt(672, 256) ), ( "Switch", "Mode", Pt(798, 299), (0xf4e1c2, 0x0c0a07), ("Bright", "Normal") ) 
		),
	]

	def __init__(self, amp_type_selector):
		sibi = amp_type_selector.sibi
		super(SHB_AmpControls,self).__init__("Amp controls", sibi, None)
		self._amp_type_selector = amp_type_selector
		self._knob = SHB_Knob("", sibi, Pt(0,0), self._getControlsShift)
		self._knob2 = SHB_Knob2("", sibi, Pt(0,0), self._getControlsShift)
		self._switch = SHB_Switch_InGroup("", sibi, Pt(0,0), (0, 0), ("",""), self._getControlsShift)
		
		self._amp_idx = None    # last seen amp type, to reset ctrl_idx
		self._amp_def = []
		self._ctrl_idx = None
		self._shift = XY(0,0)
		
		self._amp_defs = []
		for adef in self.__amp_defs_conf:
			amp = []
			for a in adef:
				if a[0] == "Knob" or a[0] == "Knob2":
					amp.append( (a[0], a[1], sibi.MXY(a[2])) )
				elif a[0] == "Switch":
					amp.append( (a[0], a[1], sibi.MXY(a[2]), a[3], a[4]) )
			self._amp_defs.append( amp )
	
	def _getControlsShift(self):
		return (self._shift.x, self._shift.y)
	
	def _getAmpIdx(self):
		return self._amp_type_selector.getSelectedIdx()
	
	def __getCurrentCtrl(self):
		if self._amp_def is None or self._ctrl_idx is None:
			return None
		adef = self._amp_def[self._ctrl_idx]
		self._shift = adef[2]
		if adef[0] == "Knob":
			self._knob.name = adef[1]
			return self._knob
		elif adef[0] == "Knob2":
			self._knob2.name = adef[1]
			return self._knob2
		elif adef[0] == "Switch":
			self._switch.name = adef[1]
			self._switch.colors = adef[3]
			self._switch.states = adef[4]
			return self._switch
		return None
		
	def _getCurrent(self):
		idx = self._getAmpIdx()
		if idx is None or idx >= len(self._amp_defs):
			log.error(len(self._amp_defs))
			self._amp_idx = None
			self._amp_def = []
			self._ctrl_idx = None
			return None
		if self._ctrl_idx is None or self._amp_idx is None or self._amp_idx != idx:
			self._ctrl_idx = 0
			self._amp_idx = idx
			self._amp_def = self._amp_defs[idx]
		return self.__getCurrentCtrl()
	
	def onLeft(self):
		idx = self._ctrl_idx
		if idx is not None:
			if idx > 0:
				self._ctrl_idx = idx - 1
			else:
				self._ctrl_idx = len(self._amp_def) - 1
			self.__getCurrentCtrl() # to update parameters
		self.sibi.speakInFocusAfter(self.reactionTime())
		return True

	def onRight(self):
		idx = self._ctrl_idx
		if idx is not None:
			if idx < len(self._amp_def) - 1:
				self._ctrl_idx = idx + 1
			else:
				self._ctrl_idx = 0
			self.__getCurrentCtrl() # to update parameters
		self.sibi.speakInFocusAfter(self.reactionTime())
		return True
	
class SHB_TopSection(Control):
	def __init__(self, name, sibi, lt, rb, btns):
		super(SHB_TopSection,self).__init__(name, sibi, None)
		lt = sibi.ptScale(lt)
		rb = sibi.ptScale(rb)
		self.label = TextBox(sibi.hwnd, lt.x, lt.y, rb.x, rb.y)
		self.btns = []
		for btn in btns:
			self.btns.append( SHB_TopBtn(sibi, *btn) )
		self.idx = 0
	
	def onHelp(self, prefix = "", suffix = ""):
		text = "Press enter to open %ss list. Use left and right arrows to select an operation with %s, and then press enter to perform that operation." % (self.name, self.name)
		if self.name == "Preset":
			text = text + " You can return to this control pressing Control+P."
		self.speak(prefix + text + suffix)
		return True
	
	def getTextInfo(self):
		if self.idx == 0:
			return (self.type, self.name, self.label.getBox().getText())
		btn = self.btns[self.idx - 1]
		status = ""
		if not btn.isEnabled():
			status = ", disabled"
		return (self.type, self.name, btn.name + status)
		
	def focusSet(self):
		super(SHB_TopSection,self).focusSet()
		self.idx = 0
	
	def onLeft(self):
		if self.idx > 0:
			self.idx -= 1
		else:
			self.idx = len(self.btns)
		self.speakAfter()
		return True

	def onRight(self):
		if self.idx >= len(self.btns):
			self.idx = 0
		else:
			self.idx += 1
		self.speakAfter()
		return True
	
	def onEnter(self):
		if self.idx == 0:
			self.label.getBox().leftClick()
			self.speakFocusAfter()
		else:
			btn = self.btns[self.idx - 1]
			btn.leftClick()
			if not btn.dialog:
				self.speakFocusAfter()
			else:
				self.expectFocusReturn()
		return True

class SHB_Dialog(Window):
	def initOverlayClass(self):
		if self.windowText == "Edit preset name":
			self.name = "Enter preset name and press Enter. Press Escape to cancel."
			self.enter_pt = MXY(self.windowHandle, 123, 167)
		elif self.windowText == "Create preset":
			self.name = "Enter new preset name and press Enter. Press Escape to cancel."
			self.enter_pt = MXY(self.windowHandle, 123, 167)
		elif self.windowText == "Remove preset":
			self.name = "Confirm removing preset by pressing Enter. Press Escape to cancel."
			self.enter_pt = MXY(self.windowHandle, 144, 122)
		if self.windowText == "Edit bank name":
			self.name = "Enter bank name and press Enter. Press Escape to cancel."
			self.enter_pt = MXY(self.windowHandle, 123, 167)
		elif self.windowText == "Create new Bank":
			self.name = "Enter new bak name and press Enter. Press Escape to cancel."
			self.enter_pt = MXY(self.windowHandle, 123, 167)
		elif self.windowText == "Remove bank":
			self.name = "Confirm removing bank by pressing Enter. Press Escape to cancel."
			self.enter_pt = MXY(self.windowHandle, 144, 122)
		else: # Error
			self.name = self.windowText + ". " + TextBox(self.windowHandle, 72, 56, 438, 103).getText()
			self.enter_pt = MXY(self.windowHandle, 228, 134)
	
	def script_enter(self, gesture):
		self.enter_pt.leftClick()
	
	__gestures = {
		"kb:enter" : "enter",
		# Escape is processed by plug-in
	}

class SHB(SIBINVDA):
	sname = "Howard Benson"
	displayText = ""
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 975, 600, None)
		sibi = self.sibi

		banks_btns = [ 
			(Pt(254, 11), "Edit name", True), 
			(Pt(88, 11), "New", True), 
			(Pt(109, 12), "Delete", True) 
		]
		sibi.add( SHB_TopSection("Bank", sibi, Pt(122, 5), Pt(219, 18), banks_btns) )
		
		preset_btns = [ 
			(Pt(632, 11), "Save", False), 
			(Pt(672, 12), "Revert", False), 
			(Pt(705, 8), "Copy", False), 
			(Pt(741, 8), "Paste", False), 
			(Pt(597, 11), "Edit name", True), 
			(Pt(363, 11), "New", True), 
			(Pt(383, 12), "Delete", True) 
		]
		sibi.add( SHB_TopSection("Preset", sibi, Pt(396, 4), Pt(563, 19), preset_btns) )

		sibi.add( SHB_IOGroup(sibi) )
		
		pages = TabGroup("Page", sibi, ("slow_reaction",))
		sibi.add( pages )
		
		pages.add( SHB_AmpSelector(sibi) )		
		pages.add( SHB_StopmsPage(sibi) )
		pages.add( SHB_CabsPage(sibi) )
		
		sibi.addDialog( SHB_List("Preset list", sibi, Pt(487, 23), ("slow_reaction",)) )
		sibi.addDialog( SHB_List("Bank list", sibi, Pt(170, 23), ("slow_reaction",)) )
		sibi.addDialog( SHB_List("Routing options", sibi, Pt(888, 23), ("slow_reaction",)) )

	def script_onCtrlP(self, gesture):
		if self._skipGesture(gesture):
			return
		if self.sibi.last_dlg:
			return # we are in dialog
		self.sibi.focusFirst()
		self.sibi.focusNext()# preset control
		self.speakFocusAfter(0)

	def script_onShiftUp(self, gesture):
		self._onScript("onShiftUp", True, gesture)

	def script_onShiftDown(self, gesture):
		self._onScript("onShiftDown", True, gesture)
		
	__gestures = {
		"kb:control+p": "onCtrlP",
		"kb:shift+upArrow": "onShiftUp",
		"kb:shift+downArrow": "onShiftDown",
	}