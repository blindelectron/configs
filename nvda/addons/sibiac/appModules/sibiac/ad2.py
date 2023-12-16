# Single Image Blob Interface Accessible Control
# 
# SynthMaster Player overlay
#
# AZ (www.azslow.com), 2018
from ctypes import *
from ctypes.wintypes import *
import time
from logHandler import log
import winUser
import math

from . import *

		
class AD2KitPeaceEdit(Label):

	def __init__(self, name, sibi, ref_pt):
		x = ref_pt.x
		y = ref_pt.y
		super(AD2KitPeaceEdit,self).__init__(name, sibi, Pt(x - 6, y + 78), Pt(x + 140, y + 91), ("dynamic_text","slow_reaction"))
		
		self.prev_btn = Box(sibi.hwnd, sibi.xScale(Pt(x,y)), sibi.yScale(Pt(x,y)))
		self.next_btn = Box(sibi.hwnd, sibi.xScale(Pt(x,y + 25)), sibi.yScale(Pt(x, y + 25)))
		self.play_btn = Box(sibi.hwnd, sibi.xScale(Pt(x + 64, y + 7)), sibi.yScale(Pt(x + 64, y + 7)))
							   
	def onDown(self):
		self.next_btn.leftClick()
		self.sibi.speakAfter(self.reactionTime())
		return True			

	def onUp(self):
		self.prev_btn.leftClick()
		self.sibi.speakAfter(self.reactionTime())
		return True
	
	def onEnter(self):
		self.play_btn.leftClick()
		return True

def _AD2OutputFor(name, sibi, ref_pt):
		btn = PopupMenuButton(name, sibi, ref_pt, ref_pt,
								Pt(ref_pt.x+1,ref_pt.y-63), Pt(ref_pt.x+1,ref_pt.y+34), 0x2c313a, (0x807b68, 0x5d5a52),
								Pt(ref_pt.x+4, ref_pt.y-106), Pt(ref_pt.x+236,ref_pt.y-87), Pt(ref_pt.x+4, ref_pt.y-86),
								0x2c313a, 0x4d5361, 0x6b4933, ("skip_spacers","slow_open","strip_prefix") )
		btn.type = "Set"
		btn.menu.type = ""
		btn.menu.name += " to"
		return btn

def _AD2KitpeaceSelect(name, sibi, ref_pt, color_off, color_on):
	kpnames = ("Cymbal 1", "Cymbal 2", "Cymbal 3", "Cymbal 4", "Cymbal 5", "Cymbal 6",
			   "Tom 1", "Tom 2", "Tom 3", "Tom 4", "Ride 1", "Ride 2",
			   "Kick", "Snare", "HiHat", "Flexi 1", "Flexi 2", "Flexi 3")
	dx = 50
	dy = 18
	tc = FixedTabControl(name, sibi, ("slow_reaction",))
	for i, kp in enumerate(kpnames):
		row, col = divmod(i, 6)
		tab = FixedTab(kp, sibi, Pt(ref_pt.x + dx*col, ref_pt.y + dy*row), 
								 Pt(ref_pt.x + 45 + dx*col, ref_pt.y + 12 + dy*row), 
								 Pt(ref_pt.x + dx*col, ref_pt.y + 6 + dy*row), 
								 color_off, color_on, ())
		tc.add( tab )
	return tc

def _AD2LearnBtn(name, sibi, ref_pt):
	btn = SwitchBtn(name, sibi, ref_pt, ref_pt, 0x000000, 0x49505c, 0x88481b, ())
	btn.type = "Learn"
	btn.on_text = "Learning"
	btn.off_text = ""
	return btn

def _AD2ChannelTab(name, sibi, x):
	return FixedTab(name, sibi, Pt(x, 493), Pt(x, 493), Pt(x, 493), (0xbcb5a4, 0x524e45), 0xc5d3e3, ())

class AD2Knob(PopupLabel):
	"""
	Turnable by mouse scroll knob, with popup label
	"""
	def __init__(self, name, sibi, center_pt, opt):
		super(AD2Knob,self).__init__(name, sibi, Pt(center_pt.x, center_pt.y + 30), Pt(center_pt.x, center_pt.y + 47),
								(0xd0d0c0, 0xf1f2dc), 0xe1e2cc, opt)
		self.kbox = Box(sibi.hwnd, sibi.xScale(center_pt), sibi.yScale(center_pt))

	def getTextInfo(self):
		""" try to fix recognition problems on transparent background... """
		type, name, text = super(AD2Knob,self).getTextInfo()
		if "%" in self.opt:
			if text[-2] == '9':
				text = text[:-2]+"%"
			if text[-1] != "%" and text[-3] == '9':
				text = text[:-3]+"%"
		return type, name, text
	
	def focusSet(self):
		""" We need to click to make it visible """
		self.kbox.moveTo()
		MouseSlowLeftClick()
		super(AD2Knob,self).focusSet()
		
	def onEnter(self):
		self.kbox.moveTo()
		MouseSlowLeftClick()
		self.sibi.speakAfter(self.reactionTime())
		return True
	
	def onUp(self):
		self.kbox.moveTo()
		MouseScroll(120)
		self.sibi.speakAfter(self.reactionTime())
		return True

	def onDown(self):
		self.kbox.moveTo()
		MouseScroll(-120)
		self.sibi.speakAfter(self.reactionTime())
		return True
		
class GroupButton(PushBtn):
	def __init__(self, name, sibi, ref_pt, tabctl, bg, fg):
		super(GroupButton,self).__init__(name, sibi, ref_pt, ref_pt, bg, fg, ("say_enabled",))
		self.tabctl = tabctl
		self.type = "Group"
		
	def onLeft(self):
		if not self.tabctl:
			return False
		self.tabctl.previousTab()
		return True

	def onRight(self):
		if not self.tabctl:
			return False
		self.tabctl.nextTab()
		return True

class AD2XBar(PopupLabel):
	"""
	Horizontal bar control. Display label when changinging, but clicking is problematic...
	"""
	def __init__(self, name, sibi, left_top, right_bottom, bar_inactive, bar_active, popup_ctop, popup_cbottom, popup_not_bg, popup_bg):
		super(AD2XBar,self).__init__(name, sibi, popup_ctop, popup_cbottom,  popup_bg, popup_not_bg, None)
		self.bar = sibi.Box(left_top, right_bottom)
		self.bar_inactive = Color2Tuple(bar_inactive)
		self.bar_active = Color2Tuple(bar_active)
		
	def _getBarRange(self, bar = None):
		if bar is None:
			bar = self.bar.getBox()
		return FindHRange(bar.hwnd, bar.left, bar.right, bar.top, self.bar_inactive, self.bar_active)
		
	def _getBarText(self):
		"""
		Should be overwritten to be more informative
		"""
		bar = self.bar.getBox()
		x0, x1 = self._getBarRange(bar)
		if x0 != bar.left:
			return "0 %%"
		return "%.0f %%" % (100. * (x1 - x0) / (self.bar.right - self.bar.left))

	def getTextInfo(self):
		type, name, text = super(AD2XBar,self).getTextInfo()
		if text == "Not visible":
			text = self._getBarText()
		else:
			if text != "96" and text.endswith("96"):
				text = text[:-2] + "%"
		return (type, name, text)

	def onUp(self):
		self.bar.moveTo()
		MouseScroll(120)
		return self.speakAfter()

	def onDown(self):
		self.bar.moveTo()
		MouseScroll(-120)
		return self.speakAfter()

class AD2SnareBuzz(AD2XBar):

	def __init__(self, sibi):
		super(AD2SnareBuzz,self).__init__("Snare Buzz", sibi, Pt(49, 331), Pt(119, 331), 0x272d37, (0x02060d, 0xd5dce6), Pt(84, 350), Pt(84, 371), 0xd9dac2, 0x030913)

	def _getBarText(self):
		bar = self.bar.getBox()
		x0, x1 = self._getBarRange(bar)
		if x0 != bar.left:
			return "no buzz"
		return "around %.0f db" % (42. * (x1 - x0) / (119 - 49) - 32 )

	def isFocusable(self):
		return False # work as control in group
		
	def getNameInGroup(self):
		type, name, text = super(AD2SnareBuzz,self).getTextInfo()
		return (name, text)
		

class AD2SnareMicPos(AD2XBar):

	def __init__(self, sibi):
		super(AD2SnareMicPos,self).__init__("Snare mic top/bottom", sibi, Pt(48, 448), Pt(121, 448), 0x2a313e, 0xa5afc5, Pt(85, 466), Pt(85, 487), 0x7f6db, (0xa69a84, 0x000000))

	def _getBarText(self):
		bar = self.bar.getBox()
		x0, x1 = self._getBarRange(bar)
		if x0 < 0:
			return "not detected"
		pos = (x0 + x1) / 2;
		return "%.0f %%" % (200. * (pos - 49) / (120 - 49) - 100 )

	def isFocusable(self):
		return False # work as control in group

	def getNameInGroup(self):
		type, name, text = super(AD2SnareMicPos,self).getTextInfo()
		return (name, text)

class AD2KickMicPos(AD2XBar):

	def __init__(self, sibi):
		super(AD2KickMicPos,self).__init__("Kick mic beater/front", sibi, Pt(48, 448), Pt(121, 448), 0x2a313e, 0xa5afc5, Pt(85, 466), Pt(85, 487), 0x7f6db, (0xa69a84, 0x000000))

	def _getBarText(self):
		bar = self.bar.getBox()
		x0, x1 = self._getBarRange(bar)
		if x0 < 0:
			return "not detected"
		pos = (x0 + x1) / 2;
		return "%.0f %%" % (200. * (pos - 49) / (120 - 49) - 100 )

class AD2RoomMicPos(AD2XBar):

	def __init__(self, sibi):
		super(AD2RoomMicPos,self).__init__("Room mic distance", sibi, Pt(49, 448), Pt(119, 448), 0x272d37, (0x02060d, 0xd5dce6), Pt(84, 466), Pt(84, 487), 0x7f6db, (0xa69a84, 0x000000))

	def _getBarText(self):
		bar = self.bar.getBox()
		x0, x1 = self._getBarRange(bar)
		if x0 != bar.left:
			return "0 ft"
		x = (1. * (x1 - x0) / (119 - 49))
		y = 56.18098*math.pow(x, 1.965667)
		return "around %.0f ft" % (y)

class AD2MicPageDetector(object):
	def __init__(self, sibi):
		self._vis_pt1 = sibi.MXY(Pt(43,323))
		self._vis_pt2 = sibi.MXY(Pt(42, 456))
		self._vis_pt3 = sibi.MXY(Pt(45, 356))
	
	def getPageIdx(self):
		if self._vis_pt2.FindNearestColor((0x01040a, 0x373c47)) != 1:
			return None # no controls
		if self._vis_pt1.FindNearestColor((0x01040a, 0x373c47)) == 1:
			return 0 # Snare
		if self._vis_pt3.FindNearestColor((0x02060e, 0xa0aabf)) == 1:
			return 1 # Room
		return 2 # Kick
		
	
class AD2(SIBINVDA):
	sname = "Addictive drums"
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 1170, 700, ("block_arrows",))
		sibi = self.sibi
		
		sibi.add( SpinLabel("Preset", sibi, Pt(42, 28), Pt(262, 45), Pt(400, 30), Pt(400, 50), ("click_on_enter", "slow_reaction")) )
		sibi.add( SwitchBtn("Play", sibi, Pt(772, 39), Pt(772,39), 0x000000, 0x5c5142, 0xd7e4fe, set()) )		
		sibi.add( OpenBtn("Save preset", sibi, Pt(345, 47), Pt(345, 47), 0x000000, 0xa2acc1, ("slow_reaction",)) ) 
		
		sibi.add( PopupMenuButton("Help", sibi, Pt(1129, 38), Pt(1129, 38),
								Pt(994, 70), Pt(994, 260), 0x2a2f39, (0x000000, 0x333129),
								Pt(993, 41), Pt(1162, 60), Pt(993, 61),
								0x2c313a, 0x4d5361, 0x6b4933, ("skip_spacers","slow_open","slow_reaction") ) )

		tv = FixedTabControl("Views", sibi, ("delayed_reaction",))
		sibi.add( tv );
		
		kit = FixedTab("Kit", sibi, Pt(919, 31), Pt( 943, 46), Pt(919, 38), 0x202327, 0x7fa7f5, ())
		tv.add(kit)

		kit.add( _AD2OutputFor("Kick output", sibi, Pt(35,663)) )
		kit.add( _AD2OutputFor("Snare output", sibi, Pt(107,663)) )
		kit.add( _AD2OutputFor("Hi hat output", sibi, Pt(179,663)) )
		kit.add( _AD2OutputFor("Tom 1 output", sibi, Pt(251,663)) )
		kit.add( _AD2OutputFor("Tom 2 output", sibi, Pt(323,663)) )
		kit.add( _AD2OutputFor("Tom 3 output", sibi, Pt(395,663)) )
		kit.add( _AD2OutputFor("Tom 4 output", sibi, Pt(467,663)) )
		kit.add( _AD2OutputFor("Flexi 1 output", sibi, Pt(539,663)) )
		kit.add( _AD2OutputFor("Flexi 2 output", sibi, Pt(611,663)) )
		kit.add( _AD2OutputFor("Flexi 3 output", sibi, Pt(683,663)) )
		kit.add( _AD2OutputFor("Overhead output", sibi, Pt(778,663)) )
		kit.add( _AD2OutputFor("Room output", sibi, Pt(850,663)) )
		kit.add( _AD2OutputFor("Bus output", sibi, Pt(922,663)) )
		
		kit.add( AD2KitPeaceEdit("Kick", sibi, Pt(92, 371)) )
		kit.add( AD2KitPeaceEdit("Snare", sibi, Pt(262, 371)) )
		kit.add( AD2KitPeaceEdit("Hi Hat", sibi, Pt(432, 371)) )

		kit.add( AD2KitPeaceEdit("Tom 1", sibi, Pt(92, 243)) )
		kit.add( AD2KitPeaceEdit("Tom 2", sibi, Pt(262, 243)) )
		kit.add( AD2KitPeaceEdit("Tom 3", sibi, Pt(432, 243)) )
		kit.add( AD2KitPeaceEdit("Tom 4", sibi, Pt(602, 243)) )
		kit.add( AD2KitPeaceEdit("Ride 1", sibi, Pt(772, 243)) )
		kit.add( AD2KitPeaceEdit("Ride 2", sibi, Pt(942, 243)) )

		kit.add( AD2KitPeaceEdit("Flexi 1", sibi, Pt(602, 371)) )
		kit.add( AD2KitPeaceEdit("Flexi 2", sibi, Pt(772, 371)) )
		kit.add( AD2KitPeaceEdit("Flexi 3", sibi, Pt(942, 371)) )

		kit.add( AD2KitPeaceEdit("Cymbal 1", sibi, Pt(92, 115)) )
		kit.add( AD2KitPeaceEdit("Cymbal 2", sibi, Pt(262, 115)) )
		kit.add( AD2KitPeaceEdit("Cymbal 3", sibi, Pt(432, 115)) )
		kit.add( AD2KitPeaceEdit("Cymbal 4", sibi, Pt(602, 115)) )
		kit.add( AD2KitPeaceEdit("Cymbal 5", sibi, Pt(772, 115)) )
		kit.add( AD2KitPeaceEdit("Cymbal 6", sibi, Pt(942, 115)) )
		
		edit = FixedTab("Edit", sibi, Pt(968, 31), Pt( 999, 46), Pt(968, 38), 0x202327, 0x7fa7f5, ())
		tv.add(edit)
		edit.add( _AD2KitpeaceSelect("Kit", sibi, Pt(66, 211), (0x393642, 0x474d5c), 0x533e34) )
		edit.add( AD2KitPeaceEdit("", sibi, Pt(146, 113)) )
		
		ctrls = Container("Kitpeace controls", sibi, ("switchable",))
		edit.add( ctrls )
		ctrls.add( PushBtn("Kitpeace controls", sibi, Pt(393, 105), Pt(393, 105), 0x414754, 0xd7e4fe, ("say_enabled",)) )
		resp = Container("Response", sibi, ("switchable",))
		ctrls.add( resp )
		resp.add( PushBtn("Response", sibi, Pt(421, 105), Pt(421, 105), 0x3e2617, 0xfe8025, ("say_enabled",)) )
		resp.add( AD2Knob("Response volume", sibi, Pt(438, 146), ("%",)) )
		resp.add( AD2Knob("Response filter", sibi, Pt(438, 226), ("%",)) )
		resp.add( PushBtn("Alternate samples", sibi, Pt(471, 127), Pt(471, 127), 0x533e34, (0x49505f, 0x373d4a), ("say_enabled", "move_away")) )
		resp.add( YRange("Sample", sibi, Pt(478, 142), Pt(501, 230), 0x1a1f28, 0xb7c2da, ()) )

		pitch = Container("Pitch", sibi, ("switchable",))
		ctrls.add( pitch )
		pitch.add( PushBtn("Pitch", sibi, Pt(531, 105), Pt(531, 105), 0x3e2617, 0xfe8025, ("say_enabled",)) )
		pitch.add( AD2Knob("Main pitch", sibi, Pt(546, 146), ()) )
		pitch.add( AD2Knob("Overheads and room pitch offset", sibi, Pt(546, 226), ()) )

		tpd = FixedTabControl("Pitch envelope/Tone designer", sibi, ("skip_control",))
		ctrls.add( tpd )

		penv = FixedTab("Pitch envelope", sibi, Pt(725, 106), Pt( 725, 106), Pt(747, 236), 0x000000, 0x343a47, ("switchable",))
		tpd.add( penv )
		penv.add( GroupButton("Pitch envelope", sibi, Pt(588, 105), tpd, 0x3e2617, 0xfe8025) )
		penv.add( ScrollLabel("Start", sibi, Pt(584, 222), Pt(627, 236), ()))
		penv.add( ScrollLabel("Hold", sibi, Pt(640, 222), Pt(685, 236), ()))
		penv.add( ScrollLabel("Release", sibi, Pt(696, 222), Pt(741, 236), ()))
		penv.add( YBar("Pitch envelope velocity", sibi, Pt(757, 122), Pt(757, 234), 0x1a1f28, 0xb7c2da, ()) )

		td = FixedTab("Tone X designer", sibi, Pt(725, 106), Pt( 725, 106), Pt(747, 236), 0x343a47, 0x000000, ("switchable",))
		tpd.add( td )
		td.add( GroupButton("Tone designer", sibi, Pt(588, 105), tpd, 0x3e2617, 0xfe8025) )
		td.add( ScrollLabel("Start", sibi, Pt(584, 222), Pt(635, 236), ()))
		td.add( ScrollLabel("Decay", sibi, Pt(646, 222), Pt(699, 236), ()))
		td.add( ScrollLabel("End", sibi, Pt(710, 222), Pt(761, 236), ()))
		
		venv = Container("Volume envelope", sibi, ("switchable",))
		ctrls.add( venv )
		venv.add( PushBtn("Volume envelope", sibi, Pt(783, 105), Pt(783, 105), 0x3e2617, 0xfe8025, ("say_enabled",)) )
		venv.add( YBar("Volume envelope velocity", sibi, Pt(783, 122), Pt(783, 234), 0x1a1f28, 0xb7c2da, ()) )
		venv.add( ScrollLabel("Attack", sibi, Pt(799, 222), Pt(844, 236), ()))
		venv.add( ScrollLabel("Decay", sibi, Pt(854, 222), Pt(900, 236), ()))
		venv.add( ScrollLabel("Sustain level", sibi, Pt(911, 222), Pt(956, 236), ()))
		venv.add( ScrollLabel("Sustain time", sibi, Pt(967, 222), Pt(1011, 236), ()))
		venv.add( ScrollLabel("Release", sibi, Pt(1023, 222), Pt(1068, 236), ()))
		
		cut = Container("Cut", sibi, ("switchable",))
		ctrls.add( cut )
		cut.add( PushBtn("Cut", sibi, Pt(1090, 105), Pt(1090, 105), 0x3e2617, 0xfe8025, ("say_enabled",)) )

		ch = FixedTabControl("Channels", sibi, ("slow_reaction",))
		edit.add( ch )
		for name, x in ( ("Kick", 37), ("Snare", 108), ("HiHat", 179), ("Tom 1", 251), ("Tom 2", 324), ("Tom 3", 394), ("Tom 4", 467),
						 ("Flexi 1", 538), ("Flexi 2", 610), ("Flexi 3", 682), ("Overheads", 780), ("Room", 850), ("Bus", 927), ("Master", 1085) ):
			ch.add( _AD2ChannelTab(name, sibi, x) )

		# Mic section
		mic_page_detector = AD2MicPageDetector(sibi)
		mic_page_idx = lambda : mic_page_detector.getPageIdx()
		mic_pages = Pages("Mic controls", sibi, mic_page_idx)
		edit.add( mic_pages )
		
		mic_snare = Group("Snare controls", sibi)
		mic_pages.add( mic_snare )
		mic_snare.add( AD2SnareBuzz(sibi) )
		mic_snare.add( AD2SnareMicPos(sibi) )
		
		mic_pages.add( AD2RoomMicPos(sibi) )
		mic_pages.add( AD2KickMicPos(sibi) )
		
			
		sends = Container("Sends", sibi, ("switchable",))
		edit.add( sends )
		sends.add( PushBtn("Sends", sibi, Pt(1085, 426), Pt(1085, 426), 0x696b6e, 0xe4fefe, ("say_enabled",)) )

		inserts = Container("Inserts", sibi, ("switchable",))
		edit.add( inserts )
		inserts.add( PushBtn("Inserts", sibi, Pt(1085, 395), Pt(1085, 395), 0x696b6e, 0xe4fefe, ("say_enabled",)) )
		
		iv = FixedTabControl("Inserts view", sibi, ("skip_control",))
		inserts.add( iv )

		fiv = FixedTab("First inserts view", sibi, Pt(0, 0), Pt( 0, 0), Pt(43, 381), 0x585248, 0x03070e, ())
		iv.add( fiv )
		noise = Container("Noise", sibi, ("switchable",))
		fiv.add( noise )
		noise.add( PushBtn("Noise", sibi, Pt(159, 304), Pt(159, 304), 0x6a6b6f, 0xfee654, ("say_enabled",)) )
		cnd = Container("C&D", sibi, ("switchable",))
		fiv.add( cnd )
		cnd.add( PushBtn("Compression and distortion", sibi, Pt(229, 304), Pt(229, 304), 0x6a6b6f, 0xfee654, ("say_enabled",)) )
		eq = Container("EQ", sibi, ("switchable",))
		fiv.add( eq )
		eq.add( PushBtn("EQ", sibi, Pt(505, 304), Pt(505, 304), 0x6a6b6f, 0xfee654, ("say_enabled",)) )
		tns = Container("T&S", sibi, ("switchable",))
		fiv.add( tns )
		tns.add( PushBtn("Tape and shape", sibi, Pt(810, 304), Pt(810, 304), 0x6a6b6f, 0xfee654, ("say_enabled",)) )

		siv = FixedTab("Bus and master inserts view", sibi, Pt(0, 0), Pt( 0, 0), Pt(43, 381), 0x03070e, 0x585248, ())
		iv.add( siv )
		cnd = Container("C&D", sibi, ("switchable",))
		siv.add( cnd )
		cnd.add( PushBtn("Compression and distortion", sibi, Pt(88, 304), Pt(88, 304), 0x6a6b6f, 0xfee654, ("say_enabled",)) )
		eq = Container("EQ", sibi, ("switchable",))
		siv.add( eq )
		eq.add( PushBtn("EQ", sibi, Pt(363, 304), Pt(363, 304), 0x6a6b6f, 0xfee654, ("say_enabled",)) )
		tns = Container("T&S", sibi, ("switchable",))
		siv.add( tns )
		tns.add( PushBtn("Tape and shape", sibi, Pt(670, 304), Pt(670, 304), 0x6a6b6f, 0xfee654, ("say_enabled",)) )
		noise = Container("Noise", sibi, ("switchable",))
		siv.add( noise )
		noise.add( PushBtn("Noise", sibi, Pt(946, 304), Pt(946, 304), 0x6a6b6f, 0xfee654, ("say_enabled",)) )
		cut = Container("Cut", sibi, ("switchable",))
		siv.add( cut )
		cut.add( PushBtn("Cut", sibi, Pt(1016, 304), Pt(1016, 304), 0x6a6b6f, 0xfee654, ("say_enabled",)) )
		
		fx = FixedTab("FX", sibi, Pt(1026, 31), Pt( 1047, 46), Pt(1026, 38), 0x202327, 0x7fa7f5, ())
		tv.add(fx)

		fx1 = Container("FX1", sibi, ("switchable",))
		fx.add( fx1 )
		fx1.add( PushBtn("FX1", sibi, Pt(44, 102), Pt(44, 102), 0x696b6e, 0xe4fefe, ("say_enabled",)) )

		fx2 = Container("FX2", sibi, ("switchable",))
		fx.add( fx2 )
		fx2.add( PushBtn("FX2", sibi, Pt(44, 275), Pt(44, 275), 0x696b6e, 0xe4fefe, ("say_enabled",)) )
		
		#####
		dlg = Dialog("Map window", sibi, Pt(182, 46), 0x000000, 0x5d6576, set())
		sibi.addDialog(dlg)
		
		btn = PopupMenuButton("Preset", sibi, Pt(182,109), Pt(362, 126),
							Pt(220, 126), Pt(220, 0), 0x2b303a, (0x000000, 0x363c49),
							Pt(224, 126), Pt(317, 145), Pt(224, 146),
							0x2b303a, 0x4c5361, 0x6b4933, ("dynamic_text", "skip_spacers", "slow_open", "slow_reaction") )
		btn.type = ""
		btn.defineSubmenu( Pt(340, 0), 3, 6, 4, 0, 20, 20)
		dlg.add( btn )

		dlg.add( Clickable("Set as default", sibi, Pt(884,580), Pt(884,580), ("silent_action",)) ) 
		dlg.add( OpenBtn("OK", sibi, Pt(834, 621), Pt(834, 621), 0x000000, 0x687983, ("OK",)) )
		dlg.add( OpenBtn("Cancel", sibi, Pt(940, 625), Pt(940, 625), 0x000000, 0x687983, ("CANCEL",)) )

		dlg.add( Label("MIDI in", sibi, Pt(209, 174), Pt(360, 189), ("dynamic_text",)) )
		dlg.add( Label("MIDI out", sibi, Pt(209, 192), Pt(360, 207), ("dynamic_text",)) )

		dlg.add( SwitchBtn("HiHat CC reverse", sibi, Pt(410, 496), Pt(410, 496), 0x000000, 0x321b0f, 0xfe8025, ()) )
		dlg.add( _AD2LearnBtn("HiHat CC", sibi, Pt(380, 520)) )
		
		dlg.add( _AD2KitpeaceSelect("Kit", sibi, Pt(464, 206), (0x474d5a, 0x3e4552), 0x533e34) )

		learn = _AD2LearnBtn("", sibi, Pt(524, 278))
		list = OptionTable("Assignments", sibi, Pt(565, 264), Pt(710, 293), Pt(565, 296), Pt(565, 645), learn, ())
		list.setScroll(Pt(752, 265), Pt(752, 646), 0x000000, 0x3c424e)
		dlg.add(list)
		
		###
		dlg = Dialog("Save map", sibi, Pt(395, 274), 0x000000, 0xabb7d0, ())
		sibi.addDialog(dlg)
		dlg.add( Clickable("Name", sibi, Pt(426,262), Pt(783,282), ("dynamic_text", "silent_action")) ) 
		dlg.add( OpenBtn("OK", sibi, Pt(400, 369), Pt(400, 359), 0x131720, 0x687983, ("OK",)) )
		dlg.add( OpenBtn("Cancel", sibi, Pt(686, 361), Pt(686, 361), 0x000000, 0x687983, ("CANCEL",)) )

		###
		dlg = Dialog("Save preset", sibi, Pt(661, 166), 0x000000, 0x8a94aa, ())
		sibi.addDialog(dlg)
		dlg.add( Clickable("Name", sibi, Pt(414, 177), Pt(653, 195), ("dynamic_text", "silent_action")) ) 
		dlg.add( OpenBtn("OK", sibi, Pt(389, 364), Pt(389, 364), 0x131720, 0x687983, ("OK",)) )
		dlg.add( OpenBtn("Cancel", sibi, Pt(686, 361), Pt(686, 361), 0x000000, 0x687983, ("CANCEL",)) )
		
		######
		dlg = Dialog("Preset Browser", sibi, Pt(185, 72), 0x000000, 0x5d6677, set())
		sibi.addDialog(dlg)
		dlg.add( Clickable("Filter by name", sibi, Pt(863,109), Pt(976,125), ("dynamic_text", "silent_action")) ) 
		dlg.add( SpinLabel("Preset", sibi, Pt(733, 494), Pt(990, 513), Pt(431, 550), Pt(471, 550), ("slow_reaction",)) )
		dlg.add( Clickable("Set as startup", sibi, Pt(930,550), Pt(930,550), ("silent_action",)) ) 
		dlg.add( OpenBtn("OK", sibi, Pt(490, 600), Pt(490,600), 0x000000, 0x687983, ("OK",)) )
		dlg.add( OpenBtn("Cancel", sibi, Pt(600, 600), Pt(600,600), 0x000000, 0x687983, ("CANCEL",)) )

	__gestures = {
		"kb:shift+upArrow": "up",
		"kb:shift+downArrow": "down",
	}
