# Single Image Blob Interface Accessible Control
# 
# Waves GTR3 overlay
#
# AZ (www.azslow.com), 2018
from ctypes import *
from ctypes.wintypes import *
import time
from logHandler import log
import winUser
import math

from . import *

class GTRSlotChooser(Control):
	"""
	Choose GTX slots with arrows, open slot menu on enter
	"""

	def __init__(self, name, sibi, opt):
		super(GTRSlotChooser,self).__init__(name, sibi, opt);
		self.slots = [] # flat list of controls
		self.slot_idx = 0 # current control index, the first is default
		self.chooser = True # indicate we have focus

	def add(self, ctrl):
		self.slots.append( ctrl )
			
	def getFocused(self):
		if self.chooser:
			return self
		focused = self.slots[self.slot_idx].getFocused()
		if focused:
			return focused
		self.chooser = True
		return self

	def getTextInfo(self):
		"""
		Should return (CtrlType, CtrlName, CtrlText) tuple
		"""
		if self.chooser and not self.slots:
			return (self.type, self.name, 'has no slots')
		(type, name, text) = self.slots[self.slot_idx].getTextInfo()
		return  (self.type, self.name, "%s, %s" % (name, text))
		
	def focusSet(self):
		"""
		Inform the control it got the focus
		"""
		focused = self.getFocused()
		if focused == self:
			super(GTRSlotChooser,self).focusSet()
			return
		focused.focusSet()

	def focusFirst(self):
		"""
		Set focus to the first element
		"""
		self.chooser = True
		self.focusSet()
		return True

	def focusLast(self):
		"""
		Set focus to the last element
		"""
		if self.slots:
			if self.slots[self.slot_idx].focusLast():
				return True
		return self.focusFirst()
				
	def focusNext(self):
		"""
		Containers process Tab as long as they can focus corresponsing control.
		"""
		if self.chooser:
			if not self.slots:
				return False
			self.chooser = False
			return self.slots[self.slot_idx].focusFirst()
		return self.slots[self.slot_idx].focusNext()
		
	def focusPrevious(self):
		"""
		Containers process Tab as long as they can focus corresponsing control.
		"""
		if self.chooser:
			return False
		focused = self.slots[self.slot_idx].focusPrevious()
		if focused:
			return focused
		return self.focusFirst()

	def onUp(self):
		if self.slots:
			if self.slot_idx > 0:
				self.slot_idx -= 1
			else:
				self.slot_idx = len(self.slots) - 1
		self.sibi.speakAfter(self.reactionTime())
		return True
	
	def onLeft(self):
		return self.onUp()
		
	def onDown(self):
		if self.slots:
			if self.slot_idx < len(self.slots) - 1:
				self.slot_idx += 1
			else:
				self.slot_idx = 0
		self.sibi.speakAfter(self.reactionTime())
		return True

	def onRight(self):
		return self.onDown()
		
	def onEnter(self):
		if self.slots:
			self.slots[self.slot_idx].openMenu()
		return True
		
class GTRMKnob(object):
	def __init__(self, name, pt, popup_pt):
		self.name = name
		self.pt = pt
		self.popup_pt = popup_pt.copy()
		# adjust for the border
		self.popup_pt.x += 2
		self.popup_pt.y += 1

class GTRMEQFader(object):
	def __init__(self, name, pt):
		self.name = name
		self.pt = pt

class GTRAKnob(GTRMKnob):
	pass
		
class GTRMDualKnob(GTRMKnob):
	def __init__(self, name, pt, popup_pt, test_pt, inactive, active, inactive_idx):
		super(GTRMDualKnob,self).__init__(name, pt, popup_pt)
		self.test_pt = test_pt
		self.inactive = inactive
		self.active = active
		self.inactive_idx = inactive_idx

class GTRMBypass(object):
	def __init__(self, name, pt, led_pt, off_color, on_color):
		""" Ignore scaling, it will not work in any case """
		self.name = name
		self.pt = pt
		self.led_pt = led_pt
		self.off_color = off_color
		self.on_color = on_color

class GTRABypass(GTRMBypass):
	pass

class GTRCBypass(GTRMBypass):
	pass

class GTRMOnOff(object):
	def __init__(self, name, pt, off_color, on_color):
		self.name = name
		self.pt = pt
		self.off_color = off_color
		self.on_color = on_color
		
class GTRModule(object):
	
	def __init__(self, colors, name, controls = None):
		self.match_colors = colors
		self.name = name
		self.ctrl = controls		

		
gGTRStompDef = [
	GTRModule( [0x313234, 0x2c2f30, 0x404042], 'Empty' ),
	GTRModule( [0xb9becf, 0x9ca2bb, 0x393e53], 'AxxPress', [
		GTRMBypass("Bypass", Pt(108, 304), Pt(109, 105), 0x3f3f3f,  0xf0a0a ),
		GTRMKnob("Output", Pt(146, 192), Pt(132,216)),
		GTRMKnob("Press", Pt(110,142), Pt(132,138)),
		GTRMKnob("Attack", Pt(71, 193), Pt(57,216)),
	] ),
	GTRModule( [0x92a2ae, 0x708192, 0x4a5358], 'Buzz', [
		GTRMBypass("Bypass", Pt(110, 296), Pt(110, 107), 0x505050, 0xff0000 ),
		GTRMKnob("Drive", Pt(69, 137), Pt(56,163)),
		GTRMKnob("Tone", Pt(110, 191), Pt(95,218)),
		GTRMKnob("Level", Pt(147, 133), Pt(132,163)),
	]	),
	GTRModule( [0x373933, 0x385182, 0x526b9d], 'Chorus', [
		GTRMBypass("Bypass", Pt(107, 299), Pt(108, 109), 0x4c4c4c, 0xff0000 ),
		GTRMKnob("Depth", Pt(71, 129), Pt(58,157)),
		GTRMKnob("Rate", Pt(146, 130), Pt(130,157)),
	]  ),
	GTRModule( [0xd3beb6, 0xc0a297, 0x573e35], 'Compressor', [
		GTRMBypass("Bypass", Pt(109, 302), Pt(110, 104), 0x717171, 0xff0000 ),
		GTRMKnob("Comp", Pt(110, 145), Pt(133,137)),
		GTRMKnob("Attack", Pt(71, 192), Pt(56,215)),
		GTRMKnob("Release", Pt(148, 192), Pt(132,215)),
	] ),
	
	GTRModule( [0x4a8bb6, 0x3487c7, 0x1c3953], 'Delay', [
		GTRMBypass("Bypass", Pt(108, 291), Pt(109, 105), 0x656565,  0xff0000 ),
		GTRMKnob("Mix", Pt(67, 130), Pt(93,117)),
		GTRMDualKnob("Manual Time", Pt(150, 131), Pt(95,133), Pt(95, 153), 0xc8a820, 0x828282, 3),
		GTRMDualKnob("Sync Time", Pt(150, 131), Pt(95,133), Pt(95, 153), 0x828282, 0xc8a820, 2),
		GTRMOnOff("Sync", Pt(95, 153), 0x828282, 0xc8a820),
		GTRMKnob("Hi cut", Pt(70, 190), Pt(93,173)),
		GTRMKnob("Feedback", Pt(149, 187), Pt(95,189)),
		GTRMKnob("Spread", Pt(110, 233), Pt(136,232)),
	] ),
	GTRModule( [0x9a8227, 0xb89a2d, 0x0], 'Distortion', [
		GTRMBypass("Bypass", Pt(109, 302), Pt(110, 107), 0x606060,  0xff0a0a ),
		GTRMKnob("Drive", Pt(75, 132), Pt(59,159)),
		GTRMKnob("Level", Pt(144, 134), Pt(128,159)),
		GTRMKnob("Contour", Pt(75, 197), Pt(59,224)),
		GTRMKnob("Tone", Pt(143, 195), Pt(128,224)),
	] ),
	GTRModule( [0xc43f1c, 0xa02f12, 0x4b1304], 'Doubler', [
		GTRMBypass("Bypass", Pt(107, 300), Pt(109, 107), 0x535353,  0xffbe00 ),
		GTRMKnob("Mix", Pt(67, 133), Pt(54,154)),
		GTRMKnob("Detune", Pt(149, 130), Pt(134,154)),
		GTRMKnob("Delay", Pt(68, 187), Pt(54,210)),
		GTRMKnob("Feedback", Pt(148, 188), Pt(134,210)),
		GTRMKnob("Spread", Pt(107, 232), Pt(93,255)),
	] ),
	GTRModule( [0x263c40, 0x395257, 0x1d2f32], 'EQ', [
		GTRMBypass("Bypass", Pt(109, 291), Pt(109, 108), 0x484848,  0xff0000 ),
		GTRMEQFader("125 Hz", Pt(59, 135)),
		GTRMEQFader("250 Hz", Pt(79, 135)),
		GTRMEQFader("500 Hz", Pt(99, 135)),
		GTRMEQFader("1 kHz", Pt(119, 135)),
		GTRMEQFader("2 kHz", Pt(139, 135)),
		GTRMEQFader("3 kHz", Pt(159, 135)),
    ] ),
	GTRModule( [0xa362ad, 0x814a8b, 0x3b203f], 'Flanger', [
		GTRMBypass("Bypass", Pt(107, 299), Pt(109, 107), 0x535353,  0xff0016 ),
		GTRMKnob("Depth", Pt(69, 129), Pt(54,154)),
		GTRMDualKnob("Manual Rate", Pt(149, 134), Pt(134,154), Pt(96, 152), 0xa9ae41, 0x757575, 3 ),
		GTRMDualKnob("Sync Rate", Pt(149, 134), Pt(134,154), Pt(96, 152), 0x757575, 0xa9ae41, 2 ),
		GTRMOnOff("Sync", Pt(96, 152), 0x757575, 0xa9ae41),
		GTRMKnob("Delay", Pt(64, 188), Pt(54,210)),
		GTRMKnob("Feedback", Pt(147, 186), Pt(134,210)),
		GTRMKnob("Spread", Pt(111, 231), Pt(93,255)),
	] ),
	GTRModule( [0x363637, 0x989ca0, 0x151515], 'Fuzz', [
		GTRMBypass("Bypass", Pt(108, 284), Pt(109, 109), 0x5c5c5c,  0xff0000 ),
		GTRMKnob("Sustain", Pt(68, 127), Pt(56,156)),
		GTRMKnob("Tone", Pt(110, 172), Pt(94,140)),
		GTRMKnob("Level", Pt(147, 127), Pt(134,156)),
	] ),
	GTRModule( [0xbacacc, 0x9db4b7, 0x3b4e51], 'Gate', [
		GTRMBypass("Bypass", Pt(106, 299), Pt(109, 104), 0x6d6d6d,  0xff0000 ),
		GTRMKnob("Thresh", Pt(70, 134), Pt(56,156)),
		GTRMKnob("Hold", Pt(146, 133), Pt(132,156)),
		GTRMKnob("Attack", Pt(71, 194), Pt(56,215)),
		GTRMKnob("Release", Pt(145, 192), Pt(132,215)),
	] ),
	GTRModule( [0xcecabb, 0xb9b49d, 0x524d3b], 'Gate-Comp', [
		GTRMBypass("Bypass", Pt(108, 301), Pt(109, 104), 0x6d6d6d,  0xff0000 ),
		GTRMKnob("Gate Thresh", Pt(75, 135), Pt(56,157)),
		GTRMKnob("Comp Thresh", Pt(147, 135), Pt(132,157)),
	] ),
	GTRModule( [0xd56491, 0xd25e89, 0xdd7da4], 'LayD', [
		GTRMBypass("Bypass", Pt(110, 298), Pt(110, 106), 0x696969,  0xff270a ),
		GTRMKnob("Mix", Pt(59, 127), Pt(49,147)),
		GTRMDualKnob("Manual Time", Pt(154, 125), Pt(140,148), Pt(96, 126), 0xaa9155, 0x999999, 3),
		GTRMDualKnob("Sync Time", Pt(154, 125), Pt(140,148), Pt(96, 126), 0x999999, 0xaa9155, 2),
		GTRMOnOff("Sync", Pt(96, 126), 0x999999, 0xaa9155 ),
		GTRMOnOff("Pitch direction", Pt(97, 149), 0x939393, 0xa38f4a),
		GTRMKnob("Pitch", Pt(62, 180), Pt(49,202)),
		GTRMKnob("Feedback", Pt(155, 182), Pt(140,202)),
		GTRMKnob("Spread", Pt(107, 180), Pt(94,202)),
	] ),
	GTRModule( [0x212121, 0x373737, 0x464646], 'Metal', [
		GTRMBypass("Bypass", Pt(111, 299), Pt(110, 109), 0x686868,  0xff0000 ),
		GTRMKnob("Dist", Pt(69, 122), Pt(93,123)),
		GTRMKnob("Level", Pt(151, 125), Pt(95,136)),
		GTRMKnob("Low Gain", Pt(71, 179), Pt(93,158)),
		GTRMKnob("High Gain", Pt(153, 181), Pt(95,174)),
		GTRMKnob("Middle Gain", Pt(149, 237), Pt(95,244)),
		GTRMKnob("Middle Freq", Pt(70, 238), Pt(93,211)),
	] ),
	GTRModule( [0xc8c8c8, 0xe2e2e2, 0x9e9e9e], 'Octaver', [
		GTRMBypass("Bypass", Pt(105, 293), Pt(109, 107), 0x606060,  0xff0000 ),
		GTRMKnob("Direct", Pt(111, 236), Pt(134,231)),
		GTRMKnob("Octave 1", Pt(72, 139), Pt(92,125)),
		GTRMKnob("Octave 2", Pt(152, 135), Pt(94,138)),
		GTRMKnob("Pan 1", Pt(72, 189), Pt(53,210)),
		GTRMKnob("Pan 2", Pt(150, 188), Pt(136,210)),
	] ),
	GTRModule( [0x4c9033, 0x539d3b, 0x2f5920], 'Overdrive', [
		GTRMBypass("Bypass", Pt(110, 297), Pt(109, 106), 0x595959,  0xff0000 ),
		GTRMKnob("Drive", Pt(66, 134), Pt(51,160)),
		GTRMKnob("Level", Pt(153, 134), Pt(136,160)),
		GTRMKnob("Tone", Pt(109, 185), Pt(94,209)),
	] ),
	GTRModule( [0x5098a6, 0x3391a5, 0x236370], 'Panner', [
		GTRMBypass("Bypass", Pt(109, 292), Pt(110, 109), 0x696969,  0xff0000 ),
		GTRMDualKnob("Manual Rate", Pt(73, 138), Pt(57,159), Pt(97, 159), 0xb5972a, 0x828282, 2),
		GTRMDualKnob("Sync Rate", Pt(73, 138), Pt(57,159), Pt(97, 159), 0x828282, 0xb5972a, 1),
		GTRMKnob("Width", Pt(146, 136), Pt(133,159)),
		GTRMOnOff("Sync", Pt(97, 159), 0x828282, 0xb5972a),
		GTRMKnob("Shape", Pt(107, 221), Pt(0, 0)),
	] ),
	GTRModule( [0xcccfd2, 0xa9acaf, 0x171717], 'Phaser', [
		GTRMBypass("Bypass", Pt(110, 281), Pt(109, 103), 0x4c4c4c,  0xff0000 ),
		GTRMKnob("Depth", Pt(71, 126), Pt(56,156)),
		GTRMDualKnob("Manual Rate", Pt(147, 130), Pt(134,156), Pt(96, 149), 0xa9ae41, 0x757575, 3),
		GTRMDualKnob("Sync Rate", Pt(147, 130), Pt(134,156), Pt(96, 149), 0x757575, 0xa9ae41, 2),
		GTRMOnOff("Sync", Pt(96, 149), 0x757575, 0xa9ae41),
		GTRMKnob("Spread", Pt( 112, 185), Pt(133,181)),
	] ),
	GTRModule( [0xb8c7d6, 0x99a4b2, 0x636a71], 'Pitcher Bass', [
		GTRMBypass("Bypass", Pt(107, 294), Pt(109, 108), 0x666666,  0xed1c1c ),
		GTRMKnob("Mix", Pt(105, 137), Pt(95,161)),
		GTRMKnob("Min Pitch", Pt(77, 189), Pt(61,159)),
		GTRMKnob("Max Pitch", Pt(141, 187), Pt(129,159)),
		GTRMKnob("Pitch", Pt(107, 228), Pt(0, 0)),
	] ),
	GTRModule( [0xec2414, 0xb91c0f, 0x701109], 'Pitcher', [
		GTRMBypass("Bypass", Pt(108, 287), Pt(109, 108), 0x5d5d5d,  0xedbe1c ),
		GTRMKnob("Mix", Pt(112, 132), Pt(95,161)),
		GTRMKnob("Min Pitch", Pt(78, 192), Pt(61,159)),
		GTRMKnob("Max Pitch", Pt(142, 191), Pt(129,159)),
		GTRMKnob("Pitch", Pt(104, 231), Pt(0, 0)),
	] ),
	GTRModule( [0xb099a2, 0xc8b7bf, 0xd8ccd1], 'Reverb', [
		GTRMBypass("Bypass", Pt(110, 289), Pt(109, 105), 0x6c6c6c,  0xff0000 ),
		GTRMKnob("Mix", Pt(75, 129), Pt(60,166)),
		GTRMKnob("Time", Pt(143, 129), Pt(131,166)),
		GTRMKnob("Pre delay", Pt(73, 201), Pt(60,236)),
		GTRMKnob("Tone", Pt(143, 201), Pt(131,236)),
	] ),
	GTRModule( [0xe3dbd7, 0x3e3635, 0x414040], 'Spring reverb', [
		GTRMBypass("Bypass", Pt(107, 297), Pt(110, 106), 0x5f5f5f,  0xff0000 ),
		GTRMKnob("Mix", Pt(103, 149), Pt(95,174)),
		GTRMKnob("Pre delay", Pt(78, 219), Pt(56,246)),
		GTRMKnob("Time", Pt(151, 225), Pt(133,246)),
	] ),
	GTRModule( [0x33434c, 0x4e616a, 0x2c454a], 'Tone', [
		GTRMBypass("Bypass", Pt(111, 295), Pt(110, 105), 0x626262,  0xff0000 ),
		GTRMKnob("Low", Pt(81, 184), Pt(61,208)),
		GTRMKnob("Mid", Pt(110, 143), Pt(95,167)),
		GTRMKnob("High", Pt(142, 187), Pt(130,208)),
		GTRMKnob("Hi Pass", Pt(78, 242), Pt(61,262)),
		GTRMKnob("Lo Pass", Pt(142, 232), Pt(130,262)),
	] ),
	GTRModule( [0x6c5033, 0xa46828, 0x774c1d], 'Vibrolo', [
		GTRMBypass("Bypass", Pt(108, 286), Pt(110, 109), 0x5f5f5f,  0xff0000 ),
		GTRMKnob("Tremolo", Pt(148, 127), Pt(134,154)),
		GTRMDualKnob("Manual Rate", Pt(109, 174), Pt(132,168), Pt(97, 126), 0x9f7d37, 0x707070, 3),
		GTRMDualKnob("Sync Rate", Pt(109, 174), Pt(132,168), Pt(97, 126), 0x707070, 0x9f7d37, 2),
		GTRMKnob("Vibrato", Pt(70, 131), Pt(55,154)),
		GTRMOnOff("Sync", Pt(97, 126), 0x9f7d37, 0x707070),
		GTRMKnob("Shape", Pt(109, 234), Pt(0, 0)),
	] ),
	GTRModule( [0x777a7c, 0x3a3e43, 0x434343], 'Volume', [
		GTRMBypass("Bypass", Pt(104, 296), Pt(109, 104), 0x696969,  0xff0000 ),
		GTRMKnob("Min", Pt(67, 128), Pt(49,148)),
		GTRMKnob("Max", Pt(153, 122), Pt(139,148)),
		GTRMKnob("Volume", Pt(160, 250), Pt(0, 0)),
		GTRMOnOff("Log mode", Pt(55, 256), 0x0d0d0d, 0xb2b3b4 ),
	] ),
	GTRModule( [0xb1b3b4, 0xa0a4a9, 0x090909], 'WahWah', [
		GTRMBypass("Bypass", Pt(110, 295), Pt(109, 103), 0x696969,  0xff0000 ),
		GTRMKnob("Sens", Pt(66, 115), Pt(49,143)),
		GTRMKnob("Speed", Pt(158, 120), Pt(139,143)),
		GTRMKnob("Range", Pt(109, 143), Pt(94,162)),
		GTRMOnOff("Manual Mode", Pt(56, 237), 0x070707, 0xa1a1a1),
		GTRMKnob("Wah", Pt(160, 251), Pt(0, 0)),
	] ),
]
	
class GTRCtrl(Control):
	"""
	Turnable by mouse scroll knobs, with popup label, enter produce right button click.
	On/Off switches.
	The existence of this controls and positions for the knob and the label are dynamic
	
	NOTE: Shift+Click is used by plug-in itself, so Shift+Tab "leaks" and does not work corrently
	"""
	def __init__(self, slot):
		super(GTRCtrl,self).__init__("Module controls,", slot.sibi, None)
		self.slot = slot
		self.n = 0
		self.text_hash = {}

	def _getCtrl(self):
		module = self.slot.matcher.match()
		if module is None or module.ctrl is None or len(module.ctrl) == 0:
			return None
		if self.n >= len(module.ctrl):
			self.n = 0 # that should not happened without focus changes, so silent
		ctrl = module.ctrl[self.n]
		if isinstance(ctrl, GTRMDualKnob):
			pt = self.sibi.ptScale(ctrl.test_pt, self.getShift())
			i = FindNearestColor(self.sibi.hwnd, pt.x, pt.y, [ctrl.inactive, ctrl.active])
			if i != 1:
				self.n = ctrl.inactive_idx
		return module.ctrl[self.n]

	def _getCtrlBox(self, ctrl):
		if ctrl is None:
			return None
		pt = self.sibi.ptScale(ctrl.pt, self.getShift())
		return Box(self.sibi.hwnd, pt.x, pt.y)
		
	def getShift(self):
		return self.slot.getShift()
		
	def _getCtrlCount(self):
		module = self.slot.matcher.match()
		if module is None or module.ctrl is None:
			return 0
		return len(module.ctrl)
	
	def getTextInfo(self):
		ctrl = self._getCtrl()
		if ctrl is None:
			return (self.type, self.name, "Does not exist")
		if isinstance(ctrl, GTRMKnob):
			type = ""
			text = "\t"
			if isinstance(ctrl, GTRAKnob):
				w = self.sibi.xScale(32)
				h = self.sibi.yScale(10)
				name_prefix = ""
			else:
				name_prefix = "Control %d, " % (self.n)
				w = self.sibi.xScale(28)
				h = self.sibi.yScale(8)			
			if ctrl.popup_pt.x != 2 or ctrl.popup_pt.y != 1:
				pt = self.sibi.ptScale(ctrl.popup_pt, self.getShift())
				box = Box( self.sibi.hwnd, pt.x, pt.y, pt.x + w, pt.y + h ) # Some are 30x8, but I do not care...
				crc32 = box.getCRC()
				if crc32 in self.text_hash:
					text = self.text_hash[crc32]
				else:
					text = box.getText()
					self.text_hash[crc32] = text
				if text == "":
					text = "Empty"
			name = "%s%s," % (name_prefix, ctrl.name)
		elif isinstance(ctrl, GTRMEQFader):
			name = "Control %d, %s" % (self.n, ctrl.name)
			box = self._getCtrlBox(ctrl)
			i = FindInYDown(box.hwnd, box.left, box.top, (0x2e474d, 0x000000), (0xa6a6a6,))
			if i < 0:
				text = "could not detect"
			else:
				log.error(i)
				# i = 0...111 -> -6...+6db
				text = "%.1f dB" % (6. - 12.*i/110)
		elif isinstance(ctrl, GTRMBypass):
			pt = self.sibi.ptScale(ctrl.led_pt, self.getShift())
			i = FindNearestColor(self.sibi.hwnd, pt.x, pt.y, [ctrl.off_color, ctrl.on_color])
			if i == 1:
				text = "inactive"
			else:
				text = "active"
			if isinstance(ctrl, GTRABypass):
				name = "Amp controls, bypass is"
			elif isinstance(ctrl, GTRCBypass):
				name = "Cam and Mic controls, bypass is"
			else:
				name = "Module controls, bypass is"
		elif isinstance(ctrl, GTRMOnOff):
			pt = self.sibi.ptScale(ctrl.pt, self.getShift())
			i = FindNearestColor(self.sibi.hwnd, pt.x, pt.y, [ctrl.off_color, ctrl.on_color])
			if i == 1:
				text = "on"
			else:
				text = "off"
			name = "Control %d, %s is" % (self.n, ctrl.name)
		else:
			name = "Unknown control type,"
			text = ""
		return ( self.type, name, text )

	def isFocusable(self):
		return self._getCtrl() is not None
		
	def focusSet(self):
		ctrl = self._getCtrl()
		if ctrl is None:
			return
		if isinstance(ctrl, GTRMKnob) or isinstance(ctrl, GTRMDualKnob):
			box = self._getCtrlBox(ctrl)
			box.moveTo()		
			MouseSlowLeftClick()
		super(GTRCtrl,self).focusSet()
		
	def onEnter(self):
		""" That opens context menu """
		ctrl = self._getCtrl()
		if ctrl is None:
			return
		box = self._getCtrlBox(ctrl)
		box.rightClick()
		return True
	
	def onUp(self):
		ctrl = self._getCtrl()
		if ctrl is None:
			return
		box = self._getCtrlBox(ctrl)
		if isinstance(ctrl, GTRMKnob) or isinstance(ctrl, GTRMEQFader):
			box.moveTo()
			MouseScroll(120)
		elif isinstance(ctrl, GTRMBypass) or isinstance(ctrl, GTRMOnOff):
			box.leftClick()
			time.sleep(0.05)
		self.sibi.speakAfter(self.reactionTime())
		return True

	def onDown(self):
		ctrl = self._getCtrl()
		if ctrl is None:
			return
		box = self._getCtrlBox(ctrl)
		if isinstance(ctrl, GTRMKnob) or isinstance(ctrl, GTRMEQFader):
			box.moveTo()
			MouseScroll(-120)
		elif isinstance(ctrl, GTRMBypass) or isinstance(ctrl, GTRMOnOff):
			box.leftClick()
		self.sibi.speakAfter(self.reactionTime())
		return True

	def onLeft(self):
		ctrl = self._getCtrl()
		if ctrl is None:
			return True
		if self.n > 0:
			self.n -= 1
		else:
			self.n = self._getCtrlCount() - 1
		n = self.n
		self._getCtrl() # that can change self.n !
		if n != self.n:
			# we hit inactive dual control, assume it is not last or first
			self.n = n - 1
		self.focusSet()
		self.sibi.speakFocusAfter(self.reactionTime())
		return True
		
	def onRight(self):
		ctrl = self._getCtrl()
		if ctrl is None:
			return
		self.n += 1
		if self.n >= self._getCtrlCount():
			self.n = 0
		n = self.n
		self._getCtrl() # that can change self.n !
		if n != self.n:
			# we hit inactive dual control, assume it is not last or first
			self.n = n + 1
		self.focusSet()
		self.sibi.speakFocusAfter(self.reactionTime())
		return True
		
	def focusFirst(self):
		self.n = 0
		self.focusSet()
		return True
		
	def focusLast(self):
		self.n = 0
		self.focusSet()
		return True
		
class GTRSlot(Container):
	"""
	Stomps slot for Stomps VST3
	
	All coordinates are for the first slot in Stomps VSTs. getShift should return correction.
	"""

	def __init__(self, sibi, n):
		global gGTRStompDef

		super(GTRSlot,self).__init__("Slot %d" % (n + 1), sibi, None)
		self.getShift = lambda : self._getShift()
		self.n = n
		self.amp_idx = -1

		self.box = Box(self.sibi.hwnd, self.sibi.xScale(111), self.sibi.yScale(360), None, None, self.getShift)
		
		match_points = [ sibi.ptScale(Pt(51, 98)), sibi.ptScale(Pt(154, 95)), sibi.ptScale(Pt(101, 322)) ] 
		self.matcher = ColorMatcher(self.sibi.hwnd, match_points, gGTRStompDef, self.getShift)
		
		self.add( GTRCtrl(self) )
		self.add( PushBtn("Module preset", self.sibi, Pt(45, 78), Pt(101, 88), None, None, ("dynamic_text",), self.getShift) )
		self.add( OpenBtn("Load module preset", self.sibi, Pt(106, 78), Pt(137, 88), None, None, set(), self.getShift) )
		self.add( PushBtn("Save module preset", self.sibi, Pt(142, 78), Pt(174, 88), None, None, set(), self.getShift) )

		# for Stomps VST3, the position is fixed
		slot_shift = 187 - 44
		self._dx = self.sibi.xScale(slot_shift * n)

	def _getShift(self):
		return (self._dx, 0)

	def _updateAmpIdx(self):
		pass
		
	def openMenu(self):
		"""
		Called by chooser on Enter, should open context menu
		"""
		self.box.getBox().leftClick()
		
	def _isEmpty(self):
		self._updateAmpIdx()
		if self.n == self.amp_idx:
			return True
		module = self.matcher.match()
		return module is None or module.name == 'Empty' 

	def getFocused(self):
		if self._isEmpty():
			return None
		return super(GTRSlot,self).getFocused();
		
	def getTextInfo(self):
		"""
		Should return (CtrlType, CtrlName, CtrlText) tuple
		"""
		self._updateAmpIdx()
		if self.n == self.amp_idx:
			return ("Slot", self.name, "Amps")
		module = self.matcher.match()
		if module is not None:
			name = module.name
		else:
			name = "Empty"
		# self.matcher.log()
		return ("Slot", self.name, name)
		

	def isFocusable(self):
		"""
		Should return True in case the control can get focus.
		That is not always the case, f.e. if under some conditions the control does not exist...
		
		General container get no focus by itself, but other containers can
		"""
		return not self._isEmpty()
		
	def focusSet(self):
		"""
		Inform the control it got the focus
		"""
		if not self._isEmpty():
			super(GTRSlot,self).focusSet();

	def focusFirst(self):
		"""
		Set focus to the first element
		"""
		if self._isEmpty():
			return False
		return super(GTRSlot,self).focusFirst();

	def focusLast(self):
		"""
		Set focus to the last element
		"""
		if self._isEmpty():
			return False
		return super(GTRSlot,self).focusLast();
				
	def focusNext(self):
		"""
		Containers process Tab as long as they can focus corresponsing control.
		"""
		if self._isEmpty():
			return False
		return super(GTRSlot,self).focusNext();
		
	def focusPrevious(self):
		"""
		Containers process Tab as long as they can focus corresponsing control.
		"""
		if self._isEmpty():
			return False
		return super(GTRSlot,self).focusPrevious();
		
class GTRRackSlot(GTRSlot):

	def __init__(self, sibi, n):
		super(GTRRackSlot,self).__init__(sibi, n)
		
		# these have different position (11,312), take dynamic shift into account
		self.box = Box(self.sibi.hwnd, self.sibi.xScale(111 - 3), self.sibi.yScale(312 + 35 ), None, None, self.getShift)
		self.amp_idx = 0
		
		self.amp_pt = self.sibi.ptScale(Pt(67, 310))
		self.slot_shift = self.sibi.xScale(187 - 47) # x distance to the next stomp slot
		self.rack_shift_pt = self.sibi.ptScale(Pt(3, -35)) # difference between Stomps and Rack
		self.amp_shift_pt = self.sibi.ptScale(Pt(-29, -35)) # slot point difference between Stomp and Amps
		self.amp_stomp_diff = self.sibi.xScale(-64) # difference in width between Stomp and Amp
		
	def _updateAmpIdx(self):
		for n in range(0,7):
			i = FindNearestColor(self.sibi.hwnd, self.amp_pt.x + n*self.slot_shift, self.amp_pt.y, [0x202020, 0xb4aea9])
			if i == 1:
				self.amp_idx = n
				return
		self.amp_idx = -1
		
	def _getShift(self):
		self._updateAmpIdx()
		if self.n == self.amp_idx:
			return (self.rack_shift_pt.x + self.slot_shift*self.n + self.amp_shift_pt.x, self.amp_shift_pt.y) # correction for slot point only
		if self.amp_idx >= 0 and self.n > self.amp_idx:
			amp_stomp_diff = self.amp_stomp_diff
		else:
			amp_stomp_diff = 0
		return (self.rack_shift_pt.x + self.slot_shift*self.n + amp_stomp_diff, self.rack_shift_pt.y)
	
class GTRStomp2(SIBINVDA):
	sname = "GTR Stomp 2 slots"
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 444, 410, ("block_arrows",))
		
		cpreset = SpinLabel("Preset", self.sibi, Pt(115, 22), Pt(223, 34), Pt(240, 28), Pt(267, 28), ("click_on_enter","load_reaction"))
		cpreset.type = "Current"
		self.sibi.add( cpreset )
		self.sibi.add( Clickable("Save A to B", self.sibi, Pt(302, 29), Pt(302, 29), None) )
		self.sibi.add( Clickable("Load preset", self.sibi, Pt(340, 31), Pt(341, 32), ("silent_action",)) )
		self.sibi.add( Clickable("Save preset", self.sibi, Pt(371, 31), Pt(372, 32), ("silent_action",)) )
		
		slots = GTRSlotChooser("Choose", self.sibi, set())
		for n in range(0, 2):
			slots.add( GTRSlot(self.sibi, n) )
		self.sibi.add(slots)

class GTRStomp4(SIBINVDA):
	sname = "GTR Stomp 4 slots"
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 731, 410, ("block_arrows",))
		
		cpreset = SpinLabel("Preset", self.sibi, Pt(131, 22), Pt(466, 34), Pt(487, 28), Pt(520, 28), ("click_on_enter","load_reaction"))
		cpreset.type = "Current"
		self.sibi.add( cpreset )
		self.sibi.add( Clickable("Save A to B", self.sibi, Pt(558, 29), Pt(558, 29), None) )
		self.sibi.add( Clickable("Load preset", self.sibi, Pt(604, 30), Pt(605, 31), ("silent_action",)) )
		self.sibi.add( Clickable("Save preset", self.sibi, Pt(652, 30), Pt(653, 31), ("silent_action",)) )
		
		slots = GTRSlotChooser("Choose", self.sibi, set())
		for n in range(0, 4):
			slots.add( GTRSlot(self.sibi, n) )
		self.sibi.add(slots)

class GTRStomp6(SIBINVDA):
	sname = "GTR Stomp 6 slots"
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 1015, 410, ("block_arrows",))
		
		cpreset = SpinLabel("Preset", self.sibi, Pt(131, 22), Pt(750, 34), Pt(773, 28), Pt(805, 28), ("click_on_enter","load_reaction"))
		cpreset.type = "Current"
		self.sibi.add( cpreset )
		self.sibi.add( Clickable("Save A to B", self.sibi, Pt(841, 28), Pt(841, 28), None) )
		self.sibi.add( Clickable("Load preset", self.sibi, Pt(890, 30), Pt(891, 31), ("silent_action",)) )
		self.sibi.add( Clickable("Save preset", self.sibi, Pt(937, 30), Pt(938, 31), ("silent_action",)) )
		
		slots = GTRSlotChooser("Choose", self.sibi, set())
		for n in range(0, 6):
			slots.add( GTRSlot(self.sibi, n) )
		self.sibi.add(slots)

class GTRAmpSlotChooser(Control):
	"""
	Choose GTX amp slots in rack with arrows
	"""

	def __init__(self, name, sibi, opt):
		super(GTRAmpSlotChooser,self).__init__(name, sibi, None);
		self.slots = [] # flat list of controls
		self.slot_idx = 0 # current control index, the first is default
		self.chooser = True # indicate we have focus

	def add(self, ctrl):
		self.slots.append( ctrl )
			
	def getFocused(self):
		if self.chooser:
			return self
		focused = self.slots[self.slot_idx].getFocused()
		if focused:
			return focused
		self.chooser = True
		return self

	def getTextInfo(self):
		"""
		Should return (CtrlType, CtrlName, CtrlText) tuple
		"""
		if self.chooser and not self.slots:
			return (self.type, self.name, 'has no slots')
		(type, name, text) = self.slots[self.slot_idx].ctrl[0].getTextInfo() # preset name
		return  (self.type, self.name, "Amp %d, %s" % (self.slot_idx + 1, text) )
		
	def focusSet(self):
		"""
		Inform the control it got the focus
		"""
		focused = self.getFocused()
		if focused == self:
			super(GTRAmpSlotChooser,self).focusSet()
			return
		focused.focusSet()

	def focusFirst(self):
		"""
		Set focus to the first element
		"""
		self.chooser = True
		self.focusSet()
		return True

	def focusLast(self):
		"""
		Set focus to the last element
		"""
		if self.slots:
			if self.slots[self.slot_idx].focusLast():
				return True
		return self.focusFirst()
				
	def focusNext(self):
		"""
		Containers process Tab as long as they can focus corresponsing control.
		"""
		if self.chooser:
			if not self.slots:
				return False
			self.chooser = False
			return self.slots[self.slot_idx].focusFirst()
		return self.slots[self.slot_idx].focusNext()
		
	def focusPrevious(self):
		"""
		Containers process Tab as long as they can focus corresponsing control.
		"""
		if self.chooser:
			return False
		focused = self.slots[self.slot_idx].focusPrevious()
		if focused:
			return focused
		return self.focusFirst()

	def onUp(self):
		if self.slots:
			if self.slot_idx > 0:
				self.slot_idx -= 1
			else:
				self.slot_idx = len(self.slots) - 1
		self.sibi.speakAfter(self.reactionTime())
		return True
	
	def onLeft(self):
		return self.onUp()
		
	def onDown(self):
		if self.slots:
			if self.slot_idx < len(self.slots) - 1:
				self.slot_idx += 1
			else:
				self.slot_idx = 0
		self.sibi.speakAfter(self.reactionTime())
		return True

	def onRight(self):
		return self.onDown()

class GTRCombo(Label):

	def __init__(self, name, sibi, left_top, right_bottom, click_pt, opt, getShift = None):
		super(GTRCombo,self).__init__(name, sibi, left_top, right_bottom, opt, getShift)
		self.click_box = Box(sibi.hwnd, sibi.xScale(click_pt), sibi.yScale(click_pt), None, None, getShift)
		self.type = "Combo"

	def onEnter(self):
		self.click_box.getBox().leftClick()
		if not "silent_action" in self.opt:
			self.sibi.speakAfter(self.reactionTime())
		return True

class GTRTempo(Label):
	"""
	Read currently set tempo, click for "Tap". Can be disabled
	"""
	def __init__(self, name, sibi, left_top, right_bottom, click_pt, disable_color, enable_color, getShift = None):
		super(GTRTempo,self).__init__("Tempo", sibi, left_top, right_bottom, ("slow_reaction", "dynamic_text"), getShift)
		self.click_box = Box(sibi.hwnd, sibi.xScale(click_pt), sibi.yScale(click_pt), None, None, getShift)
		self.colors = [disable_color, enable_color] 

	def getTextInfo(self):
		box = self.box.getBox()
		i = FindNearestColor(self.sibi.hwnd, box.left, box.top, self.colors)
		if i != 1:
			return (self.type, self.name, "disabled")
		return super(GTRTempo,self).getTextInfo()
		
	def onEnter(self):
		self.click_box.getBox().leftClick()
		if not "silent_action" in self.opt:
			self.sibi.speakAfter(self.reactionTime())
		return True
		
gGTRAmpDef = GTRModule( None, 'Flanger', [
		GTRABypass("Bypass", Pt(111, 154), Pt(111, 154), 0xf0f0f0, 0xe0e0e0 ),
		GTRAKnob("Drive", Pt(197, 120), Pt(178,70)),
		GTRAKnob("Bass", Pt(253, 116), Pt(238,70)),
		GTRAKnob("Mid", Pt(316, 119), Pt(298,70)),
		GTRAKnob("Treble", Pt(372, 118), Pt(358,70)),
		GTRAKnob("Presence", Pt(434, 123), Pt(418,70)),
	] )

gGTRCabMicDef = GTRModule( None, 'Flanger', [
		GTRCBypass("Bypass", Pt(260, 212), Pt(260, 212), 0xf0f0f0, 0xe0e0e0 ),
		GTRMOnOff("Phase", Pt(260, 263), 0xf0f0f0, 0xe0e0e0 ),
		GTRAKnob("Air", Pt(109, 263), Pt(90,292)),
		GTRAKnob("Delay", Pt(175, 264), Pt(160,292)),
		GTRAKnob("Volume", Pt(373, 265), Pt(356,292)),
		GTRAKnob("Pan", Pt(441, 266), Pt(425,292)),
	] )
	
class GTRAmpSlot(Container):

	def __init__(self, name, sibi, dx, dy):
		super(GTRAmpSlot,self).__init__(name, sibi, None)
		
		dx = sibi.xScale(dx)
		dy = sibi.yScale(dy)
		
		self.getShift = lambda : (dx, dy)
		if dx == 0:
			cpreset = SpinLabel("Amp Preset", self.sibi, Pt(42, 40), Pt(252, 52), Pt( 271, 47), Pt(308, 45), ("click_on_enter","slow_reaction"))
			cpreset.type = "Current"
			self.add( cpreset )
			self.add( Clickable("Load amp preset", self.sibi, Pt(390, 48), Pt(390, 48), ("silent_action",)) )
			self.add( Clickable("Save amp preset", self.sibi, Pt(434, 45), Pt(434, 45), ("silent_action",)) )
		else:
			cpreset = SpinLabel("Amp Preset", self.sibi, Pt(524, 40), Pt(736, 52), Pt( 756, 47), Pt(786, 45), ("click_on_enter","slow_reaction"))
			cpreset.type = "Current"
			self.add( cpreset )
			self.add( Clickable("Load amp preset", self.sibi, Pt(876, 48), Pt(876, 48), ("silent_action",)) )
			self.add( Clickable("Save amp preset", self.sibi, Pt(926, 45), Pt(926, 45), ("silent_action",)) )
		
		self.add( GTRCombo("Amp type", self.sibi, Pt(217 + dx, 150 + dy), Pt(404 + dx, 164 + dy), Pt(415 + dx, 158 + dy), ("dynamic_text", "silent_action")) )
		self.add( GTRAmpCtrl("Amp controls,", self, gGTRAmpDef) )

		self.add( GTRCombo("Cabinet", self.sibi, Pt(82 + dx, 208 + dy), Pt(207 + dx, 221 + dy), Pt(218 + dx, 215 + dy), ("dynamic_text", "silent_action")) )
		self.add( GTRCombo("Microphone", self.sibi, Pt(326 + dx, 208 + dy), Pt(451 + dx, 221 + dy), Pt(462 + dx, 215 + dy), ("dynamic_text", "silent_action")) )
		self.add( GTRAmpCtrl("Cab and mic controls,", self, gGTRCabMicDef) )

class GTRAmpCtrl(GTRCtrl):
	def __init__(self, name, slot, module):
		super(GTRAmpCtrl,self).__init__(slot)
		self.module = module
		
	def _getCtrl(self):
		if self.n > len(self.module.ctrl):
			self.n = 0
		ctrl = self.module.ctrl[self.n]
		if isinstance(ctrl, GTRMDualKnob):
			pt = self.sibi.ptScale(ctrl.test_pt, self.getShift())
			i = FindNearestColor(self.sibi.hwnd, pt.x, pt.y, [ctrl.inactive, ctrl.active])
			if i != 1:
				self.n = ctrl.inactive_idx
		return self.module.ctrl[self.n]
		
	def _getCtrlCount(self):
		return len(self.module.ctrl)
		
class GTRRack(SIBINVDA):
	sname = "GTR Tool Rack"
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 1002, 480, ("block_arrows",))
		
		cpreset = SpinLabel("Preset", self.sibi, Pt(97, 3), Pt(741, 15), Pt(763, 9), Pt(794, 10), ("click_on_enter","load_reaction"))
		cpreset.type = "Current"
		self.sibi.add( cpreset )
		self.sibi.add( Clickable("Save A to B", self.sibi, Pt(833, 8), Pt(833, 8), None) )
		self.sibi.add( Clickable("Load preset", self.sibi, Pt(881, 10), Pt(881, 10), ("silent_action",)) )
		self.sibi.add( Clickable("Save preset", self.sibi, Pt(919, 10), Pt(919, 10), ("silent_action",)) )
		self.sibi.add( SwitchBtn("HD", self.sibi, Pt(956, 10), Pt(956, 10), None, 0xaaaaaa, 0x49e300, ("slow_reaction",)) )
		
		tv = FixedTabControl("Views", self.sibi, ("delayed_reaction",))
		self.sibi.add( tv );
		
		stomps = FixedTab("Stomp", self.sibi, Pt(371, 355), Pt(371, 355), Pt(371, 355), 0x979797, 0x64f700, None)
		slots = GTRSlotChooser("Choose", self.sibi, set())
		for n in range(0,7):
			slots.add( GTRRackSlot(self.sibi, n) )
		stomps.add(slots)
		tv.add(stomps)
	
		amps = FixedTab("Amp", self.sibi, Pt(456, 357), Pt(456, 357), Pt(456, 357), 0xb7b7b7, 0x5ddf00, None)
		amps.add( SwitchBtn("Link amps", self.sibi, Pt(482, 46), Pt(482, 46), None, 0x232323, 0x49e300, ("slow_reaction",)) )
		ampslots = GTRAmpSlotChooser("Tab", self.sibi, None)
		ampslots.add( GTRAmpSlot("Amp 1", self.sibi, 0, 0) )
		ampslots.add( GTRAmpSlot("Amp 2", self.sibi, 616 - 166, 0) )
		amps.add( ampslots )
		tv.add(amps)
		
		
		tuner = FixedTab("Routing", self.sibi, Pt(524, 355), Pt(524, 355), Pt(524, 355), 0xbbbbbb, 0x64f700, None)
		tuner.add( SwitchBtn("Sync", self.sibi, Pt(202, 345), Pt(202, 345), None, 0xb4b4b4, 0x66ff00, ("slow_reaction",)) )
		tuner.add( GTRTempo("Tempo", self.sibi, Pt(108, 342), Pt(167, 353), Pt(280, 348), 0x000000, 0x2a2a2a) )
		tuner.add( Clickable("Pre routing", self.sibi, Pt(107, 363), Pt(168, 374), ("dynamic_text", "slow_reaction")) )
		tuner.add( Clickable("Post routing", self.sibi, Pt(248, 363), Pt(311, 374), ("dynamic_text", "slow_reaction")) )
		tv.add(tuner)
