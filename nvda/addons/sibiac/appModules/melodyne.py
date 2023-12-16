# -*- coding: utf-8 -*-
# Single Image Blob Interface Accessible Control
#
# SIBIAC based NVDA Application Module
# Melodyne
#
# AZ (www.azslow.com), 2018 - 2020
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

from .sibiac import *

class Fader(Label):
	def __init__(self, name, sibi, left_top, right_bottom, scroll_pt, opt):
		super(Fader,self).__init__( name, sibi, left_top, right_bottom, opt )
		self.fader = Box(sibi.hwnd, sibi.xScale(scroll_pt), sibi.yScale(scroll_pt))
		self.type = "Fader"

	def onUp(self):
		self.fader.getBox().moveTo()
		MouseScroll(120)
		self.sibi.speakAfter(self.reactionTime())
		return True

	def onDown(self):
		self.fader.getBox().moveTo()
		MouseScroll(-120)
		self.sibi.speakAfter(self.reactionTime())
		return True


class MelodyneEditText(Label):
	def __init__(self, name, sibi, left_top, right_bottom, opt = None, getShift = None):
		super(MelodyneEditText,self).__init__(name, sibi, left_top, right_bottom, ("dynamic_text",), getShift)
	
	def _isEditing(self):
		box = self.box.getBox()
		i = FindNearestColor(box.hwnd, box.right, box.bottom, (0xc2c2c2, 0xf4c295))
		return i == 1
	
	def onEnter(self):
		if self._isEditing():
			self.sibi.speakAfter(self.reactionTime())
			return False
		box = self.box.getBox()
		box.leftClick()
		box.leftClick()
		speech.cancelSpeech()
		speech.speakMessage("Text edit")
		return True
	
	def focusLost(self):
		if self._isEditing():
			box = self.box.getBox()
			outside = Box(box.hwnd, box.right + 5, box.bottom + 5)
			outside.leftClick()

class MelodyneToolEdit(MelodyneEditText):

	def isFocusable(self):
		i = self.sibi.tools._getCurrentTab()
		return i > 0 and i != 7
	
	def getTextInfo(self):
		i = self.sibi.tools._getCurrentTab()
		if i == 0 or i == 7:
			return ("", "", "")
		type, name, text = super(MelodyneToolEdit,self).getTextInfo()
		if i == 1:
			name = "Selected note"
		else:
			name = self.sibi.tools.ctrl[i].name
		return (type, name, text)

class MelodyneToolEdit2(MelodyneEditText):

	def isFocusable(self):
		i = self.sibi.tools._getCurrentTab()
		return i == 1
	
	def getTextInfo(self):
		i = self.sibi.tools._getCurrentTab()
		if i < 1 or i > 2:
			return ("", "", "")
		type, name, text = super(MelodyneToolEdit2,self).getTextInfo()
		name = "Pitch deviation"
		return (type, name, text)
		
class MelodyneMain(Control):

	def onEnter(self):
		type, name, text1 = self.sibi.tool_edit.getTextInfo()
		if text1 != "":
			type, name, text2 = self.sibi.tool_edit2.getTextInfo()
			speech.cancelSpeech()
			speech.speakMessage(text1 + " " + text2)
		return True
		
	def onUp(self):
		return False

	def onDown(self):
		return False

	def onShiftUp(self):
		return False

	def onShiftDown(self):
		return False

	def onLeft(self):
		return False

	def onRight(self):
		return False

class MelodyneOpenBtn(OpenBtn):
	"""
	Welcome buttons in Melodyne take quite some time
	"""
	def reactionTime(self):
		return 1000

class MelodyneSimpleTab(Container):

	def __init__(self, name, sibi, pt, getShift = None):
		super(MelodyneSimpleTab,self).__init__(name, sibi, None)
		self.box = Box(sibi.hwnd, sibi.xScale(pt), sibi.yScale(pt), getShift)
			
	def isActive(self):
		box = self.box.getBox()
		i = FindNearestColor(box.hwnd, box.left, box.top, (0xdadada, 0xfdfdfd))
		return i == 1
		
	def isFocusable(self):
		return self.isActive()

	def getTabName(self):
		return self.name

	def activate(self):
		self.box.getBox().leftClick()

class MelodynePitchTimeTab(MelodyneSimpleTab):

	def __init__(self, name, sibi, pt, black_pt1, black_pt2,  getShift = None):
		super(MelodynePitchTimeTab,self).__init__(name, sibi, pt)
		self.black_1 = Box(sibi.hwnd, sibi.xScale(black_pt1), sibi.yScale(black_pt1), getShift)
		self.black_2 = Box(sibi.hwnd, sibi.xScale(black_pt2), sibi.yScale(black_pt2), getShift)
			
	def isActive(self):
		if not super(MelodynePitchTimeTab,self).isActive():
			return False
		box = self.black_1.getBox()
		if FindNearestColor(box.hwnd, box.left, box.top, (0xfdfdfd, 0x000000)) != 1:
			return False
		box = self.black_2.getBox()
		return FindNearestColor(box.hwnd, box.left, box.top, (0xfdfdfd, 0x000000)) == 1

class MelodynePitchModulationTab(MelodynePitchTimeTab):

	def activate(self):
		box = self.box.getBox()
		drag = Box(box.hwnd, box.left, box.top, box.left, box.top + 24)
		drag.leftDrag()
		
class MelodynePitchDriftAttackTab(MelodynePitchModulationTab):
	
	def activate(self):
		box = self.box.getBox()
		drag = Box(box.hwnd, box.left, box.top, box.left, box.top + 48)
		drag.leftDrag()		
	
class Melodyne(SIBINVDA):
	name = "Melodyne"
	
	def _get_value(self):
		return self.windowText
	
	def isEssencial(self):
		pt = Pt(153, 58)
		x = self.sibi.xScale(pt)
		y = self.sibi.yScale(pt)
		i = FindNearestColor(self.sibi.hwnd, x, y, ( 0xbfbfbf, 0xdcdcdc, 0xfdfdfd))
		return i == 0
		
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, self.location.width, self.location.height, None)
		self.sibi.add( Clickable("Top menu", self.sibi, Pt(4,-10), Pt(4,-10), None) )
		tools = FixedTabControl("Editing tools", self.sibi, None)
		tools.type = ""
		self.sibi.tools = tools
		essential = self.isEssencial()

		tools.add( MelodyneSimpleTab("Main", self.sibi, Pt(84, 59)) )
		if not essential :
			tools.add( MelodynePitchTimeTab("Pitch", self.sibi, Pt(100, 60), Pt(105, 66), Pt(107, 62)) )
			tools.add( MelodynePitchModulationTab("Pitch modulation", self.sibi, Pt(100, 60), Pt(107, 68), Pt(107, 62)) )
			tools.add( MelodynePitchDriftAttackTab("Pitch drift", self.sibi, Pt(100, 60), Pt(111, 62), Pt(112, 69)) )
			tools.add( MelodyneSimpleTab("Formant", self.sibi, Pt(129, 59)) )
			tools.add( MelodyneSimpleTab("Amplitude", self.sibi, Pt(153, 60)) )
			tools.add( MelodynePitchTimeTab("Time", self.sibi, Pt(175, 60), Pt(171, 67), Pt(182, 67)) )
			tools.add( MelodynePitchDriftAttackTab("AttackSpeed", self.sibi, Pt(175, 60), Pt(177, 62), Pt(177, 72)) )		
		self.sibi.add( tools )
		
		self.sibi.add( MelodyneMain("Editor", self.sibi, None) )

		self.sibi.tool_edit =  MelodyneToolEdit("Edit1", self.sibi, Pt(220, 61), Pt(287, 74))
		self.sibi.add( self.sibi.tool_edit )
		self.sibi.tool_edit2 =  MelodyneToolEdit2("Edit2", self.sibi, Pt(300, 61), Pt(367, 74))
		self.sibi.add( self.sibi.tool_edit2 )
		
		xcenter = self.location.width//2
		dlg = Dialog("Information dialog", self.sibi, Pt(xcenter + 74, 44), 0x5c5c5c, 0x000000, set())
		dlg.ctrl[0].type = ""
		self.sibi.addDialog(dlg)
		dlg.add( Label("", self.sibi, Pt(xcenter - 536 + 274, 128), Pt(xcenter - 536 + 716, 142), ("dynamic_text",)) )
		dlg.add( MelodyneOpenBtn("", self.sibi, Pt(xcenter - 536 + 529, 219), Pt(xcenter - 536 + 616, 233), None, None, ("dynamic_text",)) )
		dlg.add( MelodyneOpenBtn("", self.sibi, Pt(xcenter - 536 + 701, 219), Pt(xcenter - 536 + 746, 233), None, None, ("dynamic_text",)) )

		dlg = Dialog("Startup dialog", self.sibi, Pt(xcenter, 3), 0x5c5c5c, 0xedecec, set())
		dlg.ctrl[0].type = ""
		self.sibi.addDialog(dlg)
		dlg.add( MelodyneOpenBtn("First step videos link", self.sibi, Pt(xcenter, 236), None, None, None, None) )
		dlg.add( CheckBtn("Show this window when Melodyne opens", self.sibi, Pt(0, 0), None, Pt(xcenter - 325, 538), None, 0xcdcdcd, 0x181818, None) )
		dlg.add( MelodyneOpenBtn("Close", self.sibi, Pt(xcenter + 291, 536), None, None, None, None) )

	def _speakTool(self):
		if self.sibi.info_queued > 1:
			self.sibi.info_queued -= 1
			return
		self.sibi.info_queued = 0
		i = self.sibi.tools._getCurrentTab()
		if i:
			tool = self.sibi.tools.ctrl[i]
			speech.cancelSpeech()
			speech.speakMessage(tool.name)

	def script_speakTool(self, gesture):
		self.sibi.info_queued += 1
		core.callLater(600, self._speakTool)
		gesture.send()
	
	__gestures = {
		"kb:f1": "speakTool",
		"kb:f2": "speakTool",
		"kb:f3": "speakTool",
		"kb:f4": "speakTool",
		"kb:f5": "speakTool",
	}
		
class MelodyneCombo(Label):
	"""
	It displays GNWindowMenu, but it does not move system focus there. The position is dynamic
	Arrows work there
	"""
	def __init__(self, name, sibi, box_left_top, box_right_bottom,
				 disabled_color, enabled_color, list_active_color, list_inactive_color, opt = None, getShift = None ):
		super(MelodyneCombo,self).__init__(name, sibi, box_left_top, box_right_bottom, opt, getShift)
		# self.type = "Combo"
		self.opt.add('dynamic_text')
		self.disabled_colors = Color2Tuple(disabled_color)
		self.enabled_colors = Color2Tuple(enabled_color)
		self.list_active_colors = Color2Tuple(list_active_color)
		self.list_inactive_colors = Color2Tuple(list_inactive_color)
		self.list_colors = self.list_active_colors + self.list_inactive_colors
		self.list_pt = XY(self.box.left, self.box.top + (self.box.bottom - self.box.top)//2)
		self.list_hwnd = None
		self.getShift = getShift
		self.list_item = TextBox(self.sibi.hwnd, 0, 0) # hash holder, real one will be self.list_hwnd with dynamic coordinates
		if getShift is None:
			self.getShift = lambda : (0, 0)
		

	def _isEnabled(self):
		dx,dy = self.getShift()
		if self.disabled_colors is None:
			return (dx, dy)
		i = FindNearestColor(self.sibi.hwnd, self.box.left + dx, self.box.top + dy, self.disabled_colors + self.enabled_colors + self.list_colors)
		if i < len(self.disabled_colors):
			return None
		return (dx, dy)
	
	def _isOpen(self, shift):
		dx,dy = shift
		self.list_hwnd = WindowFromPoint(self.sibi.hwnd, self.list_pt.x + dx, self.list_pt.y + dy)
		#self.list_up, self.list_down = FindInYRange(self.sibi.hwnd, self.list_right_pt.x + dx, self.list_right_pt.y + dy,
		#											self.not_list_colors, self.list_colors)
		if not self.list_hwnd or self.list_hwnd == self.sibi.hwnd:
			self.list_hwnd = None
		return self.list_hwnd is not None

	def _getListItemText(self):
		r = RECT()
		user32.GetClientRect(self.list_hwnd, byref(r))
		sel_y = FindInYDown(self.list_hwnd, r.right - 2, 1, self.list_inactive_colors, self.list_active_colors)
		if sel_y < 1: # assume the first is selected
			sel_y = 1
		else:
			sel_y += 1
		sel_h = FindInYDown(self.list_hwnd, r.right - 2, sel_y, self.list_active_colors, self.list_inactive_colors) # note reverset bg/fg
		if sel_h < 5:
			return ""
		# Fill the box
		self.list_item.hwnd = self.list_hwnd
		self.list_item.left = 15 # currently selected item is marked at the beginning
		self.list_item.top = sel_y
		self.list_item.right = r.right - 2
		self.list_item.bottom = sel_y + sel_h - 1
		return self.list_item.getText()

	def isFocusable(self):
		shift = self._isEnabled()
		return shift is not None
		
	def getTextInfo(self):
		shift = self._isEnabled()
		if shift is None:
			return (self.type, self.name, "disabled")
		if not self._isOpen(shift):
			return super(MelodyneCombo,self).getTextInfo()
		return (self.type, "list item", self._getListItemText())
		
	def onUp(self):
		return self.onDown()
	
	def onDown(self):
		shift = self._isEnabled()
		if shift is None:
			self.sibi.speakAfter(self.reactionTime())
			return True
		if not self._isOpen(shift):
			self.box.leftClick()
			self.sibi.speakFocusAfter(self.reactionTime())
			return True
		self.sibi.speakAfter(self.reactionTime())
		return False

	def focusLost(self):
		shift = self._isEnabled()
		if shift is None:
			return
		if self._isOpen(shift):
			self.box.leftClick()

	def onEnter(self):
		shift = self._isEnabled()
		if shift is None:
			return True
		if self._isOpen(shift):
			self.sibi.speakFocusAfter(self.reactionTime())
			return False
		self.sibi.speakAfter(self.reactionTime())
		return True
		
	def onEscape(self):
		shift = self._isEnabled()
		if shift is None:
			return True
		if self._isOpen(shift):
			self.sibi.speakFocusAfter(self.reactionTime())
			return False
		return True
		
class MelodyneCorrectPitch(SIBINVDA):
	
	def _get_value(self):
		return "" # it includes all text labels...

	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, self.location.width, self.location.height, None)
		self.sibi.add( Fader("Pitch center", self.sibi, Pt(380, 11), Pt(420, 28), Pt(249, 18), ("dynamic_text", )) )		
		self.sibi.add( Fader("Pitch drift", self.sibi, Pt(377, 36), Pt(423, 51), Pt(254, 45), ("dynamic_text", )) )
		self.sibi.add( CheckBtn("", self.sibi, Pt(21, 69), Pt(173, 82), Pt(15, 77), None, 0xcacaca, 0x0e0e0e, ("dynamic_text",)) )
		self.sibi.add( CheckBtn("Include notes fine-tuned manually", self.sibi, Pt(15, 95), Pt(15, 95), Pt(15, 95), 0xbdbdbd, 0xcacaca, 0x0e0e0e, None) )
		self.sibi.add( Clickable("Ok", self.sibi, Pt(360, 89), Pt(360, 89), ("silent_action",)) )
		self.sibi.cancel = Clickable("Cancel", self.sibi, Pt(268, 93), Pt(268, 93), ("silent_action",))
		self.sibi.add( self.sibi.cancel )

	def script_Close(self, gesture):
		self.sibi.cancel.onEnter()
		return True
	
	__gestures = {
		"kb:escape": "Close",
	}

class MelodyneRadioBtn(CheckBtn):
	def __init__(self, name, sibi, left_top, right_bottom, check_pt, disabled_color, off_color, on_color, opt, getShift = None):
		super(MelodyneRadioBtn,self).__init__( name, sibi, left_top, right_bottom, check_pt, disabled_color, off_color, on_color, opt, getShift )
		self.off_text = "Not selected"
		self.on_text = "Selected"
		
	def isFocusable(self):
		return self.getStateText() != "Disabled"
		
class MelodyneExport(SIBINVDA):
	
	def _get_value(self):
		return "" # it includes all text labels...

	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, self.location.width, self.location.height, None)
		self.sibi.add( MelodyneCombo("Format", self.sibi, Pt(102, 12), Pt(246, 25), 0xc3c3c3, 0xdbdbdb, 0x9c9c9c, 0xf5f5f5) )
		self.sibi.add( MelodyneCombo("Sample rate", self.sibi, Pt(312, 12), Pt(395, 25), 0xc3c3c3, 0xdbdbdb, 0x9c9c9c, 0xf5f5f5) )
		self.sibi.add( MelodyneCombo("Bit depth", self.sibi, Pt(422, 12), Pt(502, 25), 0xc3c3c3, 0xdbdbdb, 0x9c9c9c, 0xf5f5f5) )

		self.sibi.add( MelodyneCombo("Range", self.sibi, Pt(102, 35), Pt(278, 49), 0xc3c3c3, 0xdbdbdb, 0x9c9c9c, 0xf5f5f5) )
		self.sibi.add( MelodyneRadioBtn("", self.sibi, Pt(115, 59), Pt(282, 74), Pt(107, 66), 0xbfbfbf, 0xc1c1c1, 0xffffff, ("dynamic_text",)) )
		self.sibi.add( MelodyneRadioBtn("", self.sibi, Pt(114, 81), Pt(275, 94), Pt(107, 88), 0xbfbfbf, 0xc1c1c1, 0xffffff, ("dynamic_text",)) )
		self.sibi.add( MelodyneRadioBtn("Include tails", self.sibi, Pt(0,0), None, Pt(428, 67), 0xbfbfbf, 0xcacaca, 0x0e0e0e, None) )
		
		self.sibi.add( Clickable("Export", self.sibi, Pt(472, 98), None, ("silent_action",)) )
		self.sibi.cancel = Clickable("Cancel", self.sibi, Pt(357, 96), None, ("silent_action",))	
		self.sibi.add( self.sibi.cancel )

	def script_Close(self, gesture):
		self.sibi.cancel.onEnter()
		return True
	
	__gestures = {
		"kb:escape": "Close",
	}

class MelodyneTextWithBtn(Label):
	def __init__(self, name, sibi, left_top, right_bottom, btn_pt, opt = None, getShift = None):
		super(MelodyneTextWithBtn,self).__init__(name, sibi, left_top, right_bottom, ("dynamic_text",), getShift)
		self.btn = Box(self.sibi.hwnd, sibi.xScale(btn_pt), sibi.yScale(btn_pt), None, None, getShift)
	
	def onEnter(self):
		self.btn.getBox().leftClick()
		return True

class VListOfBoxes(Control):
	"""
	Virtically scrollable list of boxes. Items are separated by bg color (there should be no lines without fg color inside items)
	Scroll bar is used to scroll list when possible.
	"""
	def __init__(self, name, sibi, list_lt, list_rb, list_bg, list_fg, scroll_lt, scroll_rb, scroll_bg, scroll_fg, opt = None, getShift = None):
		super(VListOfBoxes,self).__init__(name, sibi, opt)
		self.type = "Listbox"
		self.list = sibi.Box(list_lt, list_rb, getShift)
		self.list_bg = Color2Tuple(list_bg)
		self.list_fg = Color2Tuple(list_fg)
		lt = sibi.ptScale(scroll_lt)
		rb = sibi.ptScale(scroll_rb)
		self.scroll = ScrollV(sibi.hwnd, lt.x, lt.y, rb.x, rb.y, scroll_bg, scroll_fg, getShift)
				
		self.valid = False # Set to True as long as we think the list can not change
		self.boxes = []    # Current list of items (boxes)
		self.idx = None    # Current index in items
	
	def _getItemID(self, idx):
		""" should return something which identify the item on the screen. <idx> is checked to be within <items> range """
		return None
	
	def _getRows(self):
		""" return the list of row boxes """
		rows = []
		box = self.list.getBox()
		y = box.top
		while y < box.bottom:
			height, top = FindRow(box.hwnd, (box.left, y, box.right, box.bottom), self.list_bg, self.list_fg)
			if height < 5:
				break
			y = top + height # so, bg line after
			rows.append(Box(box.hwnd, box.left, top - 1, box.right, y))
		return rows

	def _checkIdx(self):
		""" check that <idx> is correct """
		count = len(self.boxes)
		if count == 0:
			self.idx = None
		elif self.idx is None or self.idx < 0:
			self.idx = 0
		elif self.idx >= count:
			self.idx = count - 1
	
	def _update(self, force = False):
		if not self.valid or force:
			self.valid = True
			self.boxes = self._getRows()
			self._checkIdx()
		return self.idx is not None
		
	def _getBox(self, idx = None):
		if not self._update():
			return (None, None)
		if idx is None:
			idx = self.idx
		return (self.idx, self.boxes[idx])

	def _tryMoveUp(self):
		if not self._update():
			return (False, False)
		if self.idx > 0:
			self.idx -= 1
			return (True, False)
		old_id = self._getItemID(0)
		if not self.scroll.scrollUp():
			return (False, True)
		time.sleep(0.1)
		if not self._update(True):
			return (False, False)			
		count = len(self.boxes)
		idx = 1 # predicted position of current item
		if "scroll_by_two" in self.opt:
			idx = 2
		while idx >= count or (idx > 0 and self._getItemID(idx) != old_id):
			idx -= 1
		if idx > 0:
			idx = 0
		self.idx = idx
		return (True, False)

	def _tryMoveDown(self):
		if not self._update():
			return (False, False)
		count = len(self.boxes)
		if self.idx < count - 1:
			self.idx += 1
			return (True, False)
		old_id = self._getItemID(0)
		if not self.scroll.scrollDown():
			return (False, True)
		time.sleep(0.1)
		if not self._update(True):
			return (False, False)
		count = len(self.boxes)
		idx = count - 2 # predicted position of current item
		if "scroll_by_two" in self.opt:
			idx = count - 3
		while idx < 0 or (idx < count - 1 and self._getItemID(idx) != old_id):
			idx += 1
		if idx < count - 1:
			idx = count - 1
		self.idx = idx
		return (True, False)				
					
	def getTextInfo(self):
		idx, box = self._getBox()
		if idx is None:
			return (self.type, self.name, "Empty")
		text = "Item %d" % (idx + 1)
		return (self.type, self.name, text)
	
	def focusSet(self):
		self.valid = False
		super(VListOfBoxes,self).focusSet()
		
	def onUp(self):
		success, top = self._tryMoveUp()
		if not success and not top:
			self.speakInFocusAfter(0)
		elif not top:
			self.speakAfter(0)
		else:
			type, name, text = self.getTextInfo()
			self.speak(text + ", top")
		return True

	def onDown(self):
		success, bottom = self._tryMoveDown()
		if not success and not bottom:
			self.speakInFocusAfter(0)
		elif not bottom:
			self.speakAfter(0)
		else:
			type, name, text = self.getTextInfo()
			self.speak(text + ", bottom")
		return True
		
		
class MelodyneShortcuts(VListOfBoxes):
	def __init__(self, name, sibi, left_top, right_bottom):
		list_rb = Pt(359, right_bottom.y)
		super(MelodyneShortcuts,self).__init__(name, sibi, left_top, list_rb, (0xbfbfbf, 0xf9f9f9), 0x000000,
			Pt(371, 59), Pt(371, right_bottom.y - 6), 0xbfbfbf, 0x6c6c6c)
		self.type = "Treeview"
		
		self.item_open = sibi.TextBox(Pt(6,0), Pt(20,0))
		self.item_action = sibi.TextBox(Pt(21, 0), Pt(253, 0))
		self.item_shortcut = sibi.TextBox(Pt(254, 0), Pt(362, 0))

		self.__initial_idx()
		
	def __initial_idx(self):
		box = self.list.getBox()
		top, bottom = FindVRange(box.hwnd, box.left, box.top, box.bottom, 0xbfbfbf, 0xf9f9f9)
		if top < 0:
			return
		y = (top + bottom) // 2
		self._update()
		self._checkIdx()
		for idx, box in enumerate(self.boxes):
			if box.top < y and box.bottom > y:
				self.idx = idx
				break			
		
	def __updateItemPosition(self, box):
		self.item_open.top = box.top
		self.item_action.top = box.top
		self.item_shortcut.top = box.top
		self.item_open.bottom = box.bottom
		self.item_action.bottom = box.bottom
		self.item_shortcut.bottom = box.bottom

	def _isLearning(self, x, y):
		return FindNearestColor(self.sibi.hwnd, x, y, (0xbfbfbf, 0xf9f9f9)) == 1
			
	def _getItemID(self, idx):
		idx, box = self._getBox(idx)
		self.__updateItemPosition(box)
		return self.item_action.getText()
	
	def getTextInfo(self):
		idx, box = self._getBox()
		if idx is None:
			return (self.type, self.name, "Empty")
		self.__updateItemPosition(box)
		open = self.item_open.getText()
		action = self.item_action.getText()
		shortcut = self.item_shortcut.getText()
		
		if open == "v":
			return (self.type, self.name, "Open section " + action)
		if open != "":
			return (self.type, self.name, "Closed section " + action)
		if self._isLearning(box.left, box.top):
			action = "Learning " + action
		if shortcut != "":
			shortcut = ", current shortcut: " + shortcut
		return (self.type, self.name, action + shortcut)
	
	def onEnter(self):
		idx, box = self._getBox()
		if idx is None:
			return (self.type, self.name, "Empty")
		self.__updateItemPosition(box)
		open = self.item_open.getText()
		if open != "":
			self.item_open.leftClick()
		else:
			self.item_action.leftClick()
		self.speakAfter()
		return True
		
class MelodynePreferences(SIBINVDA):
	knownPages = { "Shortcuts": 0, "Recording": 1, "Audio": 2, "Chedc for Updates": 3, "User Interface": 4,
				   "Benutzeroberche": 4, "Interface ulisaheur": 4, "Inherfaz de usuario" : 4, "143 -'1' D? -71 -X" : 4, # OCR is not perfect there...
				   "Atajos de hedado": 0, u"Grabacin": 1, "Buscar actualizaciones": 3 # spanish
			     }
	def _get_value(self):
		return "" # it includes all text labels...

	def _getPageIdx(self):
		type, name, page_text = self.page_combo.getTextInfo()
		if page_text in self.knownPages:
			return self.knownPages[page_text]
		log.error("Unknown preferences page: '%s'" %(page_text))
		return None
		
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, self.location.width, self.location.height, None)
		self.page_combo = MelodyneCombo("Page", self.sibi, Pt(12, 12), Pt(263, 25), 0xc3c3c3, 0xdbdbdb, 0x9c9c9c, 0xf5f5f5)
		self.sibi.add( self.page_combo )
		getPageIdx = lambda : self._getPageIdx()
		pages = Pages("Preferences pages", self.sibi, getPageIdx)
		
		page = Container("Shortcuts", self.sibi, None)
		y = self.location.height - 19
		page.add( MelodyneShortcuts("Shortcuts", self.sibi, Pt(6, 48), Pt(371, self.location.height - 37)) )
		page.add( Clickable("Delete learning shortcut", self.sibi, Pt(130, y), None, ("silent_action",)) )
		page.add( Clickable("Reset", self.sibi, Pt(52, y), None, ("silent_action",)) )
		page.add( Clickable("Export", self.sibi, Pt(243, y), None, ("silent_action",)) )
		page.add( Clickable("Import", self.sibi, Pt(329, y), None, ("silent_action",)) )
		pages.add(page)

		page = Container("Recording", self.sibi, None)
		page.add( MelodyneTextWithBtn("Audio cache", self.sibi, Pt(148, 49), Pt(349, 64), Pt(360, 59)) )
		page.add( MelodyneEditText("Audio cache size", self.sibi, Pt(152, 74), Pt(222, 86)) )
		page.add( MelodyneCombo("Audio file format", self.sibi, Pt(152, 95), Pt(284, 108), 0xc3c3c3, 0xdbdbdb, 0x9c9c9c, 0xf5f5f5) )
		pages.add(page)

		page = Container("Audio", self.sibi, None)
		page.add( MelodyneCombo("Audio device", self.sibi, Pt(152, 50), Pt(332, 63), 0xc3c3c3, 0xdbdbdb, 0x9c9c9c, 0xf5f5f5) )
		page.add( MelodyneCombo("Sample rate", self.sibi, Pt(152, 73), Pt(355, 86), 0xc3c3c3, 0xdbdbdb, 0x9c9c9c, 0xf5f5f5) )
		page.add( MelodyneCombo("Buffer size", self.sibi, Pt(152, 96), Pt(353, 109), 0xc3c3c3, 0xdbdbdb, 0x9c9c9c, 0xf5f5f5) )
		page.add( MelodyneRadioBtn("Ignore buffer underruns", self.sibi, Pt(0,0), None, Pt(157, 126), 0xbfbfbf, 0xcacaca, 0x0e0e0e, None) )
		page.add( MelodyneCombo("Master output", self.sibi, Pt(152, 144), Pt(218, 157), 0xc3c3c3, 0xdbdbdb, 0x9c9c9c, 0xf5f5f5) )
		page.add( MelodyneCombo("Default input", self.sibi, Pt(152, 167), Pt(217, 180), 0xc3c3c3, 0xdbdbdb, 0x9c9c9c, 0xf5f5f5) )
		pages.add(page)

		page = Container("Updates", self.sibi, None)
		page.add( MelodyneCombo("Check for updates", self.sibi, Pt(152, 50), Pt(259, 63), 0xc3c3c3, 0xdbdbdb, 0x9c9c9c, 0xf5f5f5) )
		page.add( Clickable("Check Now", self.sibi, Pt(323, 57), None, ("silent_action",)) )
		page.add( Label("Status", self.sibi, Pt(146, 73), Pt(364, 89), ("dynamic_text",)) )
		page.add( Label("Last check", self.sibi, Pt(147, 95), Pt(255, 110), ("dynamic_text",)) )
		pages.add(page)

		page = Container("User interface", self.sibi, None)
		page.add( MelodyneCombo("Language", self.sibi, Pt(152, 50), Pt(333, 63), 0xc3c3c3, 0xdbdbdb, 0x9c9c9c, 0xf5f5f5) )
		page.add( MelodyneCombo("Pitch labels", self.sibi, Pt(152, 72), Pt(325, 85), 0xc3c3c3, 0xdbdbdb, 0x9c9c9c, 0xf5f5f5) )
		page.add( MelodyneCombo("Look", self.sibi, Pt(152, 94), Pt(327, 107), 0xc3c3c3, 0xdbdbdb, 0x9c9c9c, 0xf5f5f5) )
		page.add( MelodyneEditText("Default tuning", self.sibi, Pt(152, 116), Pt(220, 129)) )
		page.add( MelodyneCombo("Maximum undo levels", self.sibi, Pt(152, 138), Pt(208, 151), 0xc3c3c3, 0xdbdbdb, 0x9c9c9c, 0xf5f5f5) )
		page.add( MelodyneRadioBtn("Show tooltips", self.sibi, Pt(0,0), None, Pt(156, 168), None, 0xc1c1c1, 0x181818, None) )
		pages.add(page)
		
		self.sibi.add( pages)

gMelodyneDialogs = { "Correct Pitch": MelodyneCorrectPitch, "Export": MelodyneExport, "Preferences": MelodynePreferences,
					 u'Corregir afinaci\xf3n': MelodyneCorrectPitch, "Exportar": MelodyneExport } # Spanish
		
class AppModule(appModuleHandler.AppModule):

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		if obj.role == controlTypes.ROLE_WINDOW:
			return
		if obj.windowClassName == "GNWindowDoc":
			if obj.role == controlTypes.ROLE_PANE:
				clsList.insert(0, Melodyne)
		elif obj.windowClassName == "GNWindow":
			if obj.windowText in gMelodyneDialogs:
				clsList.insert(0, gMelodyneDialogs[obj.windowText])
			
