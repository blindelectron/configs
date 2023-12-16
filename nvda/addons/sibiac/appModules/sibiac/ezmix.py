# Single Image Blob Interface Accessible Control
# 
# EZMix 2 overlay
#
# AZ (www.azslow.com), 2018 - 2020
from ctypes import *
from ctypes.wintypes import *
import time
from logHandler import log
import winUser

from . import *

class EZMix_PresetCtl(Label):
	def __init__(self, name, sibi):
		super(EZMix_PresetCtl,self).__init__(name, sibi, Pt(174,353), Pt(346,362), ("dynamic_text", ))
		self.up_box = Box(self.sibi.hwnd, sibi.xScale(144), sibi.yScale(353))
		self.down_box = Box(self.sibi.hwnd, sibi.xScale(144), sibi.yScale(362))
		self.presets = EZMix_PresetTable(sibi)
		self.preset_text = "Empty"

	def storeText(self):
		type, name, text = super(EZMix_PresetCtl,self).getTextInfo()
		if text != "Empty":
			self.preset_text = text

	def updateText(self):
		self.sibi.knob1Ctl.storeText()
		self.sibi.knob2Ctl.storeText()
		self.sibi.presetCtl.storeText()
		
	def getTextInfo(self):
		if self.presets.col_idx == 0:
			self.updateText()
			type, name, text = super(EZMix_PresetCtl,self).getTextInfo()
			return (type, name, self.preset_text)
		col_name, text = self.presets.getCellText()
		return (self.type, self.name, "%s, %s" % (col_name, text))
		
	def onUp(self):
		self.presets.col_idx = 0
		box = self.up_box.getBox()
		i = FindNearestColor(box.hwnd, box.left, box.top, (0x5d5e5d , 0xa0a0a0))
		if i < 1:
			speech.speakMessage("First preset")
			return True
		box.leftClick()
		self.sibi.speakAfter(self.reactionTime())
		return True

	def onDown(self):
		self.presets.col_idx = 0
		box = self.down_box.getBox()
		i = FindNearestColor(box.hwnd, box.left, box.top, (0x5d5e5d , 0xa0a0a0))
		if i < 1:
			speech.speakMessage("Last preset")
			return True
		box.leftClick()
		self.sibi.speakAfter(self.reactionTime())
		return True
	
	def onLeft(self):
		col = self.presets.getColumns()
		if len(col) > 0:
			y0, y1 = self.presets._getSelectedRowPosition()
			if y0 >= 0: # sometimes not on screen, do not confuse the user
				col_idx = self.presets.col_idx - 1
				if col_idx < 0:
					col_idx = len(col) - 1
				self.presets.col_idx = col_idx
		self.sibi.speakAfter(0)
		return True

	def onRight(self):
		col = self.presets.getColumns()
		if len(col) > 0:
			y0, y1 = self.presets._getSelectedRowPosition()
			if y0 >= 0: # sometimes not on screen, do not confuse the user
				col_idx = self.presets.col_idx + 1
				if col_idx >= len(col):
					col_idx = 0
				self.presets.col_idx = col_idx
		self.sibi.speakAfter(0)
		return True

	def focusSet(self):
		self.presets.col_idx = 0
		super(EZMix_PresetCtl,self).focusSet()
		
class EZMix_FavBtn(Control):
	def __init__(self, name, sibi):
		super(EZMix_FavBtn,self).__init__(name, sibi, None)
		self.fav_box = Box(self.sibi.hwnd, sibi.xScale(375), sibi.yScale(357))

	def _isFavorite(self):
		box = self.fav_box.getBox()
		i = FindNearestColor(box.hwnd, box.left, box.top, (0x9c9c9c, 0xfef4ae))
		return i == 1

	def getTextInfo(self):
		if self._isFavorite():
			text = "Favourited, menu"
		else:
			text = "Not favourited"
		return (self.type, self.name, text)
		
	def onEnter(self):
		fav = self._isFavorite();
		box = self.fav_box.getBox()
		box.leftClick()
		if not fav:
			self.sibi.speakAfter(600)
		return True

class EZMix_Search(Clickable):
	def getTextInfo(self):
		type, name, text = super(EZMix_Search,self).getTextInfo()
		if text == "[m]": # the box is with gradient, that is recognized when empty
			text = "Empty"
		elif text == "Empty":
			try_text = self.box.getText()
			if try_text:
				text = try_text
		return (type, name, text)

	def _getStatus(self):
		box = self.status_box.getBox()
		i = FindNearestColor(self.sibi.hwnd, box.left, box.top, (0x72592b, 0x000, 0xa16608))		
		
class EZMix_Knob(Label):
	"""
	Turnable by mouse scroll knob, with popup label, dynamic name. Can be disabled.
	"""
	def __init__(self, name, sibi, name_left_top, name_right_bottom, ctl_pt, value_left_top, value_right_bottom):
		super(EZMix_Knob,self).__init__(name, sibi, name_left_top, name_right_bottom, ("dynamic_text",), None)
		self.enabled_box = Box(self.sibi.hwnd, self.box.left, self.box.top)
		self.ctl_box = Box(self.sibi.hwnd, self.sibi.xScale(ctl_pt), self.sibi.yScale(ctl_pt))
		self.popup_box = TextBox(self.sibi.hwnd, self.sibi.xScale(value_left_top), self.sibi.yScale(value_left_top),
								 self.sibi.xScale(value_right_bottom), self.sibi.yScale(value_right_bottom))
		self.name_text = "Empty"

	def storeText(self):
		type, name, name_text = super(EZMix_Knob,self).getTextInfo()
		if name_text != "Empty":
			self.name_text = name_text
			
	def getTextInfo(self):
		if not self.isDisabled():
			return (self.type, self.name, "Disabled")
		self.sibi.presetCtl.updateText()
		type, name, name_text = super(EZMix_Knob,self).getTextInfo()
		self.ctl_box.getBox().leftClick()
		time.sleep(0.05)
		if self.sibi.textout:
			value_text = self.popup_box.getBox().getTextOut()
		else:
			value_text = self.popup_box.getBox().getText()
		# value_text should end with "%"
		if value_text[-2:] == "96":
			value_text = value_text[:-2] + "%"
		
		return (type, self.name_text, value_text)
		
	def isDisabled(self):
		box = self.enabled_box.getBox()
		i = FindNearestColor(box.hwnd, box.left, box.top, (0x655b4b, 0xd0ac7c))
		return i == 1
		
	def onEnter(self):
		self.ctl_box.getBox().leftClick()
		self.sibi.speakFocusAfter(self.reactionTime())
		return True
	
	def onUp(self):
		self.ctl_box.getBox().moveTo()
		MouseScroll(12)
		self.sibi.speakAfter(self.reactionTime())
		return True

	def onDown(self):
		self.ctl_box.getBox().moveTo()
		MouseScroll(-120)
		self.sibi.speakAfter(self.reactionTime())
		return True

class EZMix_IO(Label):
	"""
	Turnable by mouse scroll level knob, with popup label. Enter audition level indicator.
	"""
	def __init__(self, name, sibi, ctl_pt, value_lt, value_rb, level_lt, level_rb):
		super(EZMix_IO,self).__init__(name, sibi, Pt(0,0), Pt(0,0), None)
		self.ctl_box = Box(sibi.hwnd, sibi.xScale(ctl_pt), sibi.yScale(ctl_pt))
		self.popup_box = TextBox(sibi.hwnd, sibi.xScale(value_lt), sibi.yScale(value_lt), sibi.xScale(value_rb), sibi.yScale(value_rb))
		self.level = Box(sibi.hwnd, sibi.xScale(level_lt), sibi.yScale(level_lt), sibi.xScale(level_rb), sibi.yScale(level_rb))
		self.clip_pt = Box(sibi.hwnd, self.level.right + sibi.xScale(1), self.level.top - sibi.yScale(2))

	def _getLevelText(self):
		clipping = ""
		i = FindNearestColor(self.sibi.hwnd, self.clip_pt.left, self.clip_pt.top, (0x230601, 0xe11310))
		if i == 1:
			clipping = ", had clipping"
			self.clip_pt.getBox().leftClick()
		x0, x1 = FindHRange(self.sibi.hwnd, self.level.left, self.level.right, self.level.top,
				   0x090909, 0x60420d)
		if x0 < 0:
			value = "No signal"
		else:
			length = x1 - x0
			value = "%.0f dB" % (1.05*length - 58.4)
			if length >= 56 and clipping != "":
				clipping = ", clipping"
		return "%s%s" % (value,clipping)
		
	def getTextInfo(self):
		""" it should be called after ctl_box operations only or when disabled """
		self.ctl_box.getBox().leftClick()
		time.sleep(0.05)
		if self.sibi.textout:
			value_text = self.popup_box.getBox().getTextOut()
		else:
			value_text = self.popup_box.getBox().getText()
		level_text = self._getLevelText()
		return (self.name, "%s, Level "%(value_text), level_text)
					
	def onEnter(self):
		speech.cancelSpeech()
		speech.speakMessage(self._getLevelText())
		return True
	
	def onUp(self):
		self.ctl_box.getBox().moveTo()
		MouseScroll(12)
		self.sibi.speakInFocusAfter(self.reactionTime())
		return True

	def onDown(self):
		self.ctl_box.getBox().moveTo()
		MouseScroll(-120)
		self.sibi.speakInFocusAfter(self.reactionTime())
		return True
	
class EZMix_FilterColumn(TextOutBox):
	def __init__(self, filter, left, right, old = None):
		super(EZMix_FilterColumn,self).__init__(filter.hwnd, left, filter.hdr_top, right, filter.hdr_bottom, None, old)
		self.filter = filter
		self.name = self.getText() # assume static name, we are recreated when it can be changed
		self.rows = None
		self.row_idx = None
		if isinstance(old, EZMix_FilterColumn) and old == self:
			self.rows = old.rows
			self.row_idx = old.row_idx
		self.update()

	def _getRow(self, idx):
		if self.rows and idx < len(self.rows):
			return self.rows[idx]
		return None

	def update(self, invalidate = False):
		if invalidate:
			self.rows = None
		rows = []
		idx = 0
		y = self.filter.col_top
		while y < self.filter.col_bottom:
			right = self.right - self.filter.col_scroll_width
			height, top = FindRow(self.hwnd, (self.left, y, right, self.filter.col_bottom), self.filter.col_bg, self.filter.col_fg)
			if height < 5:
				break
			y = top + height # so, bg line after
			rows.append(EZMix_FilterRow(self.hwnd, self.left, top - 1, right, y, None, self._getRow(idx)))
			idx += 1
		self.rows = rows
		if self.rows:
			if self.row_idx is None or self.row_idx >= len(self.rows):
				self.row_idx = 0
		else:
			self.row_idx = None
		return rows

	def getCurrentRow(self):
		if self.row_idx is None:
			return None
		return self.rows[self.row_idx]
		
	def getRows(self):
		if self.rows is None:
			return []
		return self.rows

class EZMix_FilterRow(TextOutBox):
	def __init__(self, hwnd, left, top, right, bottom, getShift = None, old = None):
		super(EZMix_FilterRow,self).__init__(hwnd, left, top, right, bottom, getShift, old)
		self.text = self.getTextOut() # semi-static text
		if self.text == "":
			self.text = None # try later

	def getText(self):
		if self.text is None:
			self.text = super(EZMix_FilterRow,self).getText()
		return self.text

class EZMix_FilterBtn(Control):
	def __init__(self, sibi):
		super(EZMix_FilterBtn,self).__init__("Filters", sibi, None)
		self.hwnd = sibi.hwnd
		self.pt = Box(self.hwnd, sibi.xScale(297), sibi.yScale(14))
		self.num = TextOutBox(sibi.hwnd, sibi.xScale(311), sibi.yScale(13), sibi.xScale(323), sibi.yScale(26))
		if self.__hasFilters():
			self.num.getText() # cache before search can destroy it
	
	def __hasFilters(self):
		i = FindNearestColor(self.hwnd, self.num.left, self.num.top, (0x0b0b0b, 0xfec74d))
		return i == 1

	def __isExpanded(self):
		i = FindNearestColor(self.hwnd, self.pt.left, self.pt.top, (0xd6d3d3, 0x3a3939))
		return i == 1
		
	def getTextInfo(self):
		if self.__isExpanded():
			text = "expanded"
		else:
			text = "collapsed"
		if self.__hasFilters():
			num = self.num.getText()
			if num != "":
				text += ", %s filters selected" % num
		return (self.type, self.name, text)
	
	def onEnter(self):
		self.sibi.presets.valid = False # collapsing the filter changes presets table size
		self.pt.leftClick()
		self.sibi.speakAfter(self.reactionTime())
	
	def onDelete(self):
		if self.__hasFilters():
			self.num.leftClick()
			self.pt.moveTo() # avoid mouse on num box, it changes the text...
		self.sibi.speakFocusAfter(self.reactionTime())			

class EZMix_FilterTable(Control):
	
	def __init__(self, sibi):
		super(EZMix_FilterTable,self).__init__("empty", sibi, None)
		self.hwnd = sibi.hwnd
		self.type = "Filter by"
		self.left = sibi.xScale(14) # horisontal reference
		self.right = sibi.xScale(508) # right border
		self.hdr_top = sibi.yScale(40) # header vertical reference, also for columns detection
		self.hdr_bottom = sibi.yScale(53) # for header labels
		self.hdr_bg = Color2Tuple(0x1f1f1f) #hdr background color
		self.hdr_border = Color2Tuple(0x0c0c0c) # border color
		
		self.col_top = sibi.yScale(57) # column top
		self.col_bottom = sibi.yScale(165) # column bottom
		self.col_scroll_width = sibi.xScale(8) # the width of scroll box
		self.col_bg = Color2Tuple((0x393939,0xbfb9a9))
		self.col_fg = Color2Tuple((0x1f4589,0xdcccac,0x641c30))
		self.col_scroll = ScrollV(self.hwnd, self.col_scroll_width//2, sibi.yScale(58), self.col_scroll_width//2, sibi.yScale(164), 
								  0x1b1b1b, 0x5f5f5f, lambda : self.__scrollShift())
		self.col_select_colors = (0x393939, 0xbfb9a9) # not selected, selected
		self.col_sel_shift = sibi.xScale(4) # shift from top right corner for cleaning current selection
		
		self.col = [] # an array of FilterColumns
		self.col_idx = None # current column index
		
		self.collapse_pt = self.sibi.ptScale(Pt(297, 14))

		self.border_x = sibi.xScale(10)
		self.border_top = self.hdr_top
		self.border_bg = Color2Tuple(0x0a0a0a)
		self.border_fg = Color2Tuple(0x2d2d2d)
		
		self.valid = False # to updateLayout when required
		sibi.filter = self # so other can invalidate us
		
	def isFocusable(self):
		i = FindNearestColor(self.hwnd, self.collapse_pt.x, self.collapse_pt.y, (0xd6d3d3, 0x3a3939))
		return i == 1

	def getBottom(self):
		if not self.isFocusable():
			return -1
		height = FindInYDown(self.hwnd, self.border_x, self.border_top, self.border_fg, self.border_bg)
		if height < 1:
			return -1
		return self.hdr_top + height - 1
		
	def __scrollShift(self):
		if self.col_idx is None:
			return (0, 0)
		return (self.col[self.col_idx].right - self.col_scroll_width, 0)
	
	def _getCol(self, idx):
		if self.col and idx < len(self.col):
			return self.col[idx]
		return None
	
	def getCurrentCol(self):
		if self.col_idx is None:
			return None
		return self.col[self.col_idx]
	
	def getCurrentRow(self):
		col = self.getCurrentCol()
		if col is None:
			return None
		return col.getCurrentRow()
	
	def _updateLayout(self):
		"""
		Invalidate on:
		when demo welcome is closed, when our sibi get focus, on filter expand and tab changes
		"""
		if self.valid:
			return
		self.col_bottom = self.getBottom()
		self.col_scroll.bottom = self.col_bottom - 1
		col = []
		idx = 0
		x = self.left
		while x < self.right:
			w = FindInXRight(self.hwnd, x, self.hdr_top, self.hdr_bg, self.hdr_border)
			if w < self.col_scroll_width:
				break # too narrow to be practical
			nc = EZMix_FilterColumn(self, x, x + w - 1, self._getCol(idx))
			if nc.name == "":
				break # not usefull without name
			col.append(nc)
			x += w + 1
			idx += 1
		self.col = col
		if len(self.col) > 0:
			if self.col_idx is None or self.col_idx >= len(self.col):
				self.col_idx = 0
		else:
			self.col_idx = None

		self.valid = True

	def _changeRow(self, delta):
		col = self.getCurrentCol()
		if col is None:
			return False
		rows = col.getRows()
		if len(rows) == 0:
			return False
		row_idx = col.row_idx + delta
		if row_idx >= 0 and row_idx < len(rows):
			col.row_idx = row_idx
			return True
		current_item_text = col.getCurrentRow().getText()
		if row_idx < 0:
			if not self.col_scroll.scrollUp():
				return False
		else:
			if not self.col_scroll.scrollDown():
				return False
		time.sleep(0.05)
		col.update(True) # we know there had to be changes
		rows = col.getRows()
		if len(rows) == 0: # something fishy
			return True # let audition the result
		# we expect scroll by 2, we just need to find old item
		if row_idx < 0:
			row_idx = 2
			if row_idx >= len(rows):
				row_idx = 0 # should not happened
			while row_idx > 0:
				new_item_text = rows[row_idx].getText()
				if new_item_text == current_item_text: # found!
					row_idx -= 1
					break
				row_idx -= 1
		else:
			row_idx -= 3
			if row_idx < 0 or row_idx >= len(rows):
				row_idx = len(rows) - 1 # should not happened
			else:
				while row_idx < len(rows) - 1:
					new_item_text = rows[row_idx].getText()
					if new_item_text == current_item_text: # found!
						row_idx += 1
						break
					row_idx += 1
		col.row_idx = row_idx # already validated
		return True
		
	def _getCellText(self):
		col = self.getCurrentCol()
		if col is None:
			return ("columns not set", "")
		row = col.getCurrentRow()
		if row is None:
			text = "no rows"
		else:
			text = row.getText()
			i = FindNearestColor(self.hwnd, row.left, row.top, self.col_select_colors)
			if i == 1:
				text += " selected"
		return ("%s," % col.name, text)
	
	def getTextInfo(self):
		self._updateLayout()
		name, text = self._getCellText()
		return (self.type, name, text)

	def onDown(self):
		self._updateLayout()
		if self._changeRow(1) or self.getCurrentRow() is None:
			self.sibi.speakAfter(0)
		else:
			name, text = self._getCellText()
			speech.speakMessage("%s, bottom" % text)
		return True

	def onUp(self):
		self._updateLayout()
		if self._changeRow(-1) or self.getCurrentRow() is None:
			self.sibi.speakAfter(0)
		else:
			name, text = self._getCellText()
			speech.speakMessage("%s, top" % text)
		return True
	
	def onLeft(self):
		self._updateLayout()
		if self.col_idx is not None:
			self.col_idx -= 1
			if self.col_idx < 0:
				self.col_idx = len(self.col) - 1
		self.sibi.speakInFocusAfter(0)
		return True

	def onRight(self):
		self._updateLayout()
		if self.col_idx is not None:
			self.col_idx += 1
			if self.col_idx >= len(self.col):
				self.col_idx = 0
		self.sibi.speakInFocusAfter(0)
		return True

	def onEnter(self):
		self._updateLayout()
		row = self.getCurrentRow()
		if row is None:
			self.sibi.speakFocusAfter(0)
			return True
		row.leftClick()
		self.valid = False
		self.sibi.speakFocusAfter(self.reactionTime())

	def onDelete(self):
		self._updateLayout()
		col = self.getCurrentCol()
		if col is not None:
			box = Box(self.hwnd, col.right - self.col_sel_shift, col.top + self.col_sel_shift)
			box.leftClick()
			col.moveTo() # move away from preset box
		self.valid = False
		self.sibi.speakInFocusAfter(self.reactionTime())

class EZMix_Effect(object):
	def __init__(self, name, pt, colors):
		self.name = name
		self.pt = pt
		self.colors = colors

class EZMix_Effects(object):
	__known_effects = (
		EZMix_Effect("Amplifier, Cabinet", Pt(728, 368), (0x16110c, 0x5a4733)),
		EZMix_Effect("Chorus", Pt(644, 437), (0x130f0d, 0x635046)),
		EZMix_Effect("Compressor", Pt(613, 209), (0x130f0d, 0x988359)),
		EZMix_Effect("De-esser", Pt(717, 185), (0x130f0d, 0x6a8288)),
		EZMix_Effect("Delay", Pt(614, 62), (0x130f0d, 0x53595f)),
		EZMix_Effect("Distortion", Pt(594, 432), (0x130f0d, 0xe56c2b)),
		EZMix_Effect("Equilizer", Pt(736, 136), (0x130f0d, 0xc4be7e)),
		EZMix_Effect("Exciter", Pt(725, 205), (0x130f0d, 0x696c76)),
		EZMix_Effect("Flanger", Pt(681, 432), (0x130f0d, 0x835380)),
		EZMix_Effect("Gate", Pt(600, 109), (0x130f0d, 0x8a8a8a)),
		EZMix_Effect("Overloud", Pt(725, 52), (0x130f0d, 0xbebcbd)),
		EZMix_Effect("Octaver", Pt(622, 430), (0x130f0d, 0x9b6332)),
		EZMix_Effect("Phaser", Pt(567, 433), (0x130f0d, 0x72ba54)),
		EZMix_Effect("Reverb", Pt(650, 115), (0x130f0d, 0xc5bbb3)),
		EZMix_Effect("Rotary", Pt(580, 335), (0x130f0d, 0x4d291b)),
		EZMix_Effect("Stereo Enhancer", Pt(623, 131), (0x130f0d, 0xafadad)),
		EZMix_Effect("Tape simulator", Pt(623, 351), (0x130f0d, 0xc7bdaa)),
		EZMix_Effect("Transient", Pt(606, 196), (0x130f0d, 0x3b443a)),
		EZMix_Effect("Tremolo", Pt(751, 436), (0x130f0d, 0x46947f)),
		EZMix_Effect("Vibrato", Pt(707, 431), (0x130f0d, 0x466a87)),
		EZMix_Effect("Wah-Wah", Pt(535, 429), (0x130f0d, 0xaeafa7)),
	)
	
	def __init__(self, sibi):
		self.sibi = sibi
		self.hwnd = sibi.hwnd
		self._effects = []
		for e in self.__known_effects:
			self._effects.append(EZMix_Effect(e.name, sibi.ptScale(e.pt), e.colors))
	
	def getCurrentEffects(self):
		elist = []
		for e in self._effects:
			if FindNearestColor(self.hwnd, e.pt.x, e.pt.y, e.colors) == 1:
				elist.append(e.name)
		if len(elist) == 0:
			return ""
		return ": %s" % (", ".join(elist))

class EZMix_PresetColumn(TextOutBox):
	def __init__(self, presets, left, right, old = None):
		super(EZMix_PresetColumn,self).__init__(presets.hwnd, left, presets.hdr_top, right, presets.hdr_top + presets.hdr_height - 1, None, old)
		self.presets = presets
		self.name = self.getText() # assume static name, we are recreated when it can be changed, can clash with sort box
	
class EZMix_PresetTable(object):
	
	def __init__(self, sibi):
		self.sibi = sibi
		self.hwnd = sibi.hwnd
		self.type = "Filter by"
		self.left = sibi.xScale(14) # horisontal reference
		self.right = sibi.xScale(500) # right, excluding following scrollbox
		self.hdr_top = sibi.yScale(175) # header vertical reference, also for columns detection, dynamic
		self.hdr_height = sibi.yScale(14) # for header labels
		self.hdr_bg = Color2Tuple((0x1f1f1f, 0x2e2e2e)) #hdr background colors, including sorting
		self.hdr_border = Color2Tuple(0x0c0c0c) # border color
		
		self.rows_top_dy = sibi.yScale(17) # rows top, relative to hdr_top
		self.bottom = sibi.yScale(319) # bottom, absolute
		self.scroll_x = sibi.xScale(505) # the center of scroll box
		self.rows_not_selected_bg = Color2Tuple((0x393939, 0x454545, 0x1c1c1c)) # all colors except selected
		self.rows_selected_bg = Color2Tuple(0xbfb9a9) # bg color when the row is selected
		self.scroll = ScrollV(self.hwnd, self.scroll_x, self.hdr_top + self.rows_top_dy, self.scroll_x, sibi.yScale(318), 
								  0x1b1b1b, 0x5f5f5f) # top is dynamic
		
		self.col = [] # an array of FilterColumns
		self.col_idx = 0 # current column index, 0 is always name

		self.border_x = sibi.xScale(10)
		self.border_bg = Color2Tuple(0x0a0a0a)
		self.border_fg = Color2Tuple(0x2d2d2d)
		
		self.valid = False # to updateLayout when required
		sibi.presets = self # so other can invalidate us
		
		self.effects = EZMix_Effects(sibi)

	def updateTop(self):
		""" re-evaluate table top """
		filter_bottom = self.sibi.filter.getBottom()
		if filter_bottom < 1:
			self.hdr_top = 40
		else:
			height = FindInYDown(self.hwnd, self.border_x, filter_bottom + 1, self.border_bg, self.border_fg)
			if height < 1:
				self.hdr_top = 40
		self.scroll.top = self.hdr_top + self.rows_top_dy
	
	def _getCol(self, idx):
		if self.col and idx < len(self.col):
			return self.col[idx]
		return None
	
	def _getCurrentCol(self):
		if self.col_idx is None or self.col is None:
			return None
		return self.col[self.col_idx]

	def getColumns(self):
		self._updateLayout()
		if self.col is None:
			return []
		return self.col
		
	def _getSelectedRowPosition(self):
		return FindVRange(self.hwnd, self.left, self.hdr_top + self.rows_top_dy - 1, self.bottom,
						  self.rows_not_selected_bg, self.rows_selected_bg)
		
	def _updateLayout(self):
		"""
		Invalidate on:
		when demo welcome is closed, when our sibi get focus, on filter expand and tab changes
		"""
		if self.valid:
			return
		self.updateTop()
		col = []
		idx = 0
		x = self.left
		while x < self.right:
			w = FindInXRight(self.hwnd, x, self.hdr_top, self.hdr_bg, self.hdr_border)
			if w < 8:
				break # too narrow to be practical
			nc = EZMix_PresetColumn(self, x, x + w - 1, self._getCol(idx))
			if nc.name == "":
				break # not usefull without name
			col.append(nc)
			x += w + 1
			idx += 1
		self.col = col
		if len(self.col) > 0:
			if self.col_idx is None or self.col_idx >= len(self.col):
				self.col_idx = 0
		else:
			self.col_idx = 0

		self.valid = True
		
	def getCellText(self):
		self._updateLayout()
		col = self._getCurrentCol()
		if col is None:
			return ("columns not set", "")
		y0, y1 = self._getSelectedRowPosition()
		if y0 < 0:
			return (col.name, "not available")
		effects = ""
		colname = col.name+","
		if colname.startswith("F"):
			colname = "Effects "
			effects = self.effects.getCurrentEffects()
		box = TextOutBox(self.hwnd, col.left, y0, col.right, y1)
		return ("%s" % colname, "%s%s" % (box.getText(), effects))

class EZMix_DemoCloseBtn(CloseBtn):

	def onEnter(self):
		self.sibi.filter.valid = False
		self.sibi.presets.valid = False
		super(EZMix_DemoCloseBtn,self).onEnter()

class EZMix_TabControl(FixedTabControl):
	def nextTab(self):
		self.sibi.filter.valid = False
		return super(EZMix_TabControl,self).nextTab()

	def previousTab(self):
		self.sibi.filter.valid = False
		return super(EZMix_TabControl,self).previousTab()
		
class EZMix(SIBINVDA):
	sname = "EZ Mix"
	displayText = ""
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 776, 465, ("block_arrows","textout"))
		#self.sibi = SIBI(self.windowHandle, 776, 465, ("block_arrows",))
		tab = EZMix_TabControl("List", self.sibi, None)
		tab.add( FixedTab("All presets", self.sibi, Pt(142, 24), Pt(185, 34), Pt(165, 23), 0x717070, 0xd7d4d4, None) )
		tab.add( FixedTab("Favorites", self.sibi, Pt(216, 22),  Pt(269, 35), Pt(236, 23), 0x717070, 0xd7d4d4, None) )
		self.sibi.add( tab )
		
		self.sibi.add( EZMix_FilterBtn(self.sibi) )
		self.sibi.add( EZMix_FilterTable(self.sibi) )
		
		self.sibi.add( EZMix_Search("Search", self.sibi, Pt(333, 15), Pt(488, 25), ("dynamic_text",)) )
		
		self.sibi.presetCtl = EZMix_PresetCtl("Preset", self.sibi)
		self.sibi.add( self.sibi.presetCtl )
		
		self.sibi.knob1Ctl = EZMix_Knob("Control 1", self.sibi, Pt(43, 394), Pt(130, 403), Pt(193, 405), Pt(179,444), Pt(211,457))
		self.sibi.add( self.sibi.knob1Ctl )
		self.sibi.knob2Ctl = EZMix_Knob("Control 2", self.sibi, Pt(379, 394), Pt(470, 403), Pt(328, 406), Pt(312,444), Pt(344,457))
		self.sibi.add( self.sibi.knob2Ctl )

		self.sibi.add( EZMix_FavBtn("", self.sibi) )

		self.sibi.add( EZMix_IO("Input",  self.sibi, Pt(104, 355), Pt(83,324), Pt(125,337), Pt(25, 349), Pt(81, 349)) )
		self.sibi.add( EZMix_IO("Output", self.sibi, Pt(417, 355), Pt(394,324), Pt(441,337), Pt(439, 349), Pt(495, 349)) )
		
		self.sibi.add( Clickable("Menu", self.sibi, Pt(740, 17), Pt(740, 17), ("silent_action",)) )
		
		dlg = Dialog("Demo session time is expired", self.sibi, Pt(508, 377), 0x4f2813, 0xb9b9b9, None)
		self.sibi.addDialog(dlg)

		dlg = Dialog("Session will expire in 15 minutes", self.sibi, Pt(508, 333), 0x4f2813, 0xfefefe, None)
		dlg.add( EZMix_DemoCloseBtn("Close", self.sibi, Pt(582, 104), Pt(582, 104), None) )
		self.sibi.addDialog(dlg)

	def event_gainFocus(self):
		self.sibi.filter.valid = False
		self.sibi.presets.valid = False
		time.sleep(0.1) # give time to redraw
		super(EZMix,self).event_gainFocus()
		
	__gestures = {
		"kb:shift+upArrow": "up",
		"kb:shift+downArrow": "down",
	}
