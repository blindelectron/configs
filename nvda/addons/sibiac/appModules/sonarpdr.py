# Single Image Blob Interface Accessible Control
#
# SIBIAC based NVDA Application Module
# Sonar X3
#
# AZ (www.azslow.com), 2017 - 2020
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

from .sibiac import *

class SIBGuitarRig5(SIBINVDA):

	name = "Guitar Rig dialog" # to avoid "None", plug-in window already has the text

	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 1018, 810)
		self.sibi.add( Label("Current preset", self.sibi, Pt(416, 57), Pt(565, 79), ("dynamic_text",) ) )
		self.sibi.add( VList("Tag", self.sibi, Pt(19,122), Pt(132,134), Pt(19,139), Pt(19,406,MOVE,2),
							0x242424, 0xfaaf2f, set(("click_on_enter",))) )
		self.sibi.add( VList("Subtag", self.sibi, Pt(138,122), Pt(251,134), Pt(138,139), Pt(138,406,MOVE,2),
							0x242424, 0xfaaf2f, set(("click_on_enter",))) )
		self.sibi.add( VList("Group", self.sibi, Pt(257,122), Pt(370,134), Pt(257,139), Pt(257,406,MOVE,2),
							0x242424, 0xfaaf2f, set(("click_on_enter",))) )
		self.sibi.add( VList("Preset", self.sibi, Pt(47,481,MOVE,2), Pt(193,497,MOVE,2), Pt(47,501,MOVE,2), Pt(47,765,MOVE),
							0x242424, 0xfaaf2f, set(("click_to_focus", "pass_enter"))) )

class Absynth5(SIBINVDA):

	name = "Absynth 5"
	
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

class SIElectricPiano(SIBINVDA):

	name = "\t"
	displayText = ""
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 827, 587)
		self.sibi.add( PushBtn("Program", self.sibi, Pt(53, 81), Pt(147, 97), 0x000000, 0x303031, ("dynamic_text","silent_action")) )

class SIDrumKit(SIBINVDA):

	name = "\t"
	displayText = ""
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 848, 588)
		self.sibi.add( PushBtn("Program", self.sibi, Pt(50, 82), Pt(146, 98), 0x000000, 0x303031, ("dynamic_text","silent_action")) )

class SIBassGuitar(SIBINVDA):

	name = "\t"
	displayText = ""
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 827, 587)
		self.sibi.add( PushBtn("Program", self.sibi, Pt(53, 81), Pt(147, 97), 0x000000, 0x303031, ("dynamic_text","silent_action")) )

class SIStringSection(SIBINVDA):

	name = "\t"
	displayText = ""
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 850, 570)
		self.sibi.add( PushBtn("Program", self.sibi, Pt(52, 82), Pt(146, 98), 0x000000, 0x303031, ("dynamic_text","silent_action")) )

class DimensionPro(SIBINVDA):

	name = "\t"
	displayText = ""
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 752, 503)
		self.sibi.add( PushBtn("Program", self.sibi, Pt(37, 11), Pt(378, 25), 0x000000, 0x170f2d, ("dynamic_text","silent_action")) )

class SDKitPiece(Label):
	"""
	LeftClick on enter,
	Right click on down
	"""
	# Control Interface

	def onEnter(self):
		self.box.leftClick()
		return True

	def onDown(self):
		self.box.rightClick()
		return True

class SDOutput(Label):
	"""
	Combo with accessible list on clickLabel, support "dynamic_text"  and "silent_action"
	"""
	def __init__(self, name, sibi, left_top, right_bottom, combo_point):
		super(SDOutput,self).__init__( name, sibi, left_top, right_bottom, ("dynamic_text",))
		self.combo_box = Box( sibi.hwnd, sibi.xScale(combo_point), sibi.yScale(combo_point), sibi.xScale(combo_point), sibi.yScale(combo_point) )

	def onEnter(self):
		self.combo_box.leftClick()
		return True
		
		
class SessionDrummer(SIBINVDA):

	name = "\t"
	displayText = ""
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 964, 570)
		
		tabs = FixedTabControl("", self.sibi, ());
		self.sibi.add( tabs );
		
		drumkit = FixedTab("Drumkit", self.sibi, Pt(738, 29), Pt(806, 42), Pt(743, 34), 0x5d646d, 0xf3f8fd, ())
		drumkit.add( PushBtn("Program", self.sibi, Pt(112, 468), Pt(422, 482), 0x000000, 0x170f2d, ("dynamic_text","silent_action")) )
		drumkit.add( SDKitPiece("Kick", self.sibi, Pt(450, 350), Pt(479, 378), set()) )
		drumkit.add( SDKitPiece("Snare", self.sibi, Pt(290, 313), Pt(305, 327), set()) )
		drumkit.add( SDKitPiece("Hi Hat", self.sibi, Pt(181, 264), Pt(196, 276), set()) )
		drumkit.add( SDKitPiece("Hi tom", self.sibi, Pt(374, 173), Pt(396, 191), set()) )
		drumkit.add( SDKitPiece("Mid tom", self.sibi, Pt(516, 171), Pt(548, 196), set()) )
		drumkit.add( SDKitPiece("Low tom", self.sibi, Pt(646, 332), Pt(670, 349), set()) )
		drumkit.add( SDKitPiece("Crash", self.sibi, Pt(248, 151), Pt(262, 167), set()) )
		drumkit.add( SDKitPiece("Ride", self.sibi, Pt(724, 231), Pt(751, 251), set()) )
		drumkit.add( SDKitPiece("FX 1", self.sibi, Pt(128, 339), Pt(136, 358), set()) )
		drumkit.add( SDKitPiece("FX 2", self.sibi, Pt(162, 338), Pt(170, 346), set()) )
		drumkit.add( SDKitPiece("FX 3", self.sibi, Pt(139, 381), Pt(151, 391), set()) )
		drumkit.add( SDKitPiece("FX 4", self.sibi, Pt(175, 371), Pt(184, 379), set()) )
		tabs.add( drumkit )

		mixer = FixedTab("Mixer", self.sibi, Pt(817, 29), Pt(884, 42), Pt(823, 34), 0x5d646d, 0xf3f8fd, ())
		mixer.add( SDOutput("Kick output", self.sibi, Pt(53, 530), Pt(77, 539), Pt(85, 534)) )
		mixer.add( SDOutput("Snare output", self.sibi, Pt(107, 530), Pt(147, 539), Pt(152, 534)) )
		mixer.add( SDOutput("Hi Hat output", self.sibi, Pt(176, 530), Pt(218, 539), Pt(223, 534)) )
		mixer.add( SDOutput("Hi tom output", self.sibi, Pt(248, 530), Pt(288, 539), Pt(295, 534)) )
		mixer.add( SDOutput("Mid tom output", self.sibi, Pt(319, 530), Pt(358, 539), Pt(363, 534)) )
		mixer.add( SDOutput("Low tom output", self.sibi, Pt(389, 530), Pt(429, 539), Pt(435, 534)) )
		mixer.add( SDOutput("Crash output", self.sibi, Pt(462, 530), Pt(498, 539), Pt(504, 534)) )
		mixer.add( SDOutput("Ride output", self.sibi, Pt(529, 530), Pt(568, 539), Pt(574, 534)) )
		mixer.add( SDOutput("FX 1 output", self.sibi, Pt(605, 530), Pt(638, 539), Pt(644, 534)) )
		mixer.add( SDOutput("FX 2 output", self.sibi, Pt(670, 530), Pt(708, 539), Pt(715, 534)) )
		mixer.add( SDOutput("FX 3 output", self.sibi, Pt(737, 530), Pt(778, 539), Pt(784, 534)) )
		mixer.add( SDOutput("FX 4 output", self.sibi, Pt(807, 530), Pt(848, 539), Pt(854, 534)) )
		tabs.add( mixer )
			
class SIBTL64(SIBINVDA):
	
	name = "Tube Leveler dialog" # to avoid "External VST"
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 774, 122)
		self.sibi.add( PushBtn("Power", self.sibi, Pt(40,58), Pt(59,85),
								0x0e0e0e, 0x2e2e2e, set()) )

class SIProgramBrowser(object):

	name = "Program browser"
	
	def event_gainFocus(self):
		try:
			self.firstChild.firstChild.setFocus()
		except:
			pass

class AD2PresetSpin(Label):
	
	def __init__(self, name, sibi, left_top, right_bottom, 
					prev_center, next_center, opt):
		super(AD2PresetSpin,self).__init__( name, sibi, left_top, right_bottom, opt)
		self.type = 'Choose'
		self.opt.add('dynamic_text')

		self.prev_btn = Box(sibi.hwnd, sibi.xScale(prev_center), sibi.yScale(prev_center),
									   sibi.xScale(prev_center), sibi.yScale(prev_center))
		self.next_btn = Box(sibi.hwnd, sibi.xScale(next_center), sibi.yScale(next_center),
									   sibi.xScale(next_center), sibi.yScale(next_center))
							   
	def onDown(self):
		self.next_btn.leftClick()
		self.sibi.speakAfter(self.reactionTime())
		return True			

	def onUp(self):
		self.prev_btn.leftClick()
		self.sibi.speakAfter(self.reactionTime())
		return True
	
	def onEnter(self):
		if "click_on_enter" not in self.opt:
			return True
		self.box.leftClick()
		self.sibi.speakFocusAfter(self.reactionTime())
		return True
		
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
		
class AD2(SIBINVDA):
	name = "\t"
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 1170, 700)

		self.sibi.add( AD2PresetSpin("Preset", self.sibi, Pt(42, 28), Pt(262, 45), Pt(400, 30), Pt(400, 50), ("click_on_enter","slow_reaction")) )
		self.sibi.add( SwitchBtn("Play", self.sibi, Pt(772, 39), Pt(772,39), 0x000000, 0x5c5142, 0xd7e4fe, set()) )		
		self.sibi.add( OpenBtn("Save preset", self.sibi, Pt(345, 47), Pt(345, 47), 0x000000, 0xa2acc1, ("slow_reaction",)) ) 
		
		self.sibi.add( PopupMenuButton("Help", self.sibi, Pt(1129, 38), Pt(1129, 38),
								Pt(994, 70), Pt(994, 260), 0x2a2f39, (0x000000, 0x333129),
								Pt(993, 41), Pt(1162, 60), Pt(993, 61),
								0x2c313a, 0x4d5361, 0x6b4933, ("skip_spacers","slow_open","slow_reaction") ) )

		tv = FixedTabControl("Views", self.sibi, ("delayed_reaction",))
		self.sibi.add( tv );
		
		kit = FixedTab("Kit", self.sibi, Pt(919, 31), Pt( 943, 46), Pt(919, 38), 0x202327, 0x7fa7f5, ())
		tv.add(kit)

		kit.add( _AD2OutputFor("Kick output", self.sibi, Pt(35,663)) )
		kit.add( _AD2OutputFor("Snare output", self.sibi, Pt(107,663)) )
		kit.add( _AD2OutputFor("Hi hat output", self.sibi, Pt(179,663)) )
		kit.add( _AD2OutputFor("Tom 1 output", self.sibi, Pt(251,663)) )
		kit.add( _AD2OutputFor("Tom 2 output", self.sibi, Pt(323,663)) )
		kit.add( _AD2OutputFor("Tom 3 output", self.sibi, Pt(395,663)) )
		kit.add( _AD2OutputFor("Tom 4 output", self.sibi, Pt(467,663)) )
		kit.add( _AD2OutputFor("Flexi 1 output", self.sibi, Pt(539,663)) )
		kit.add( _AD2OutputFor("Flexi 2 output", self.sibi, Pt(611,663)) )
		kit.add( _AD2OutputFor("Flexi 3 output", self.sibi, Pt(683,663)) )
		kit.add( _AD2OutputFor("Overhead output", self.sibi, Pt(778,663)) )
		kit.add( _AD2OutputFor("Room output", self.sibi, Pt(850,663)) )
		kit.add( _AD2OutputFor("Bus output", self.sibi, Pt(922,663)) )
		
		kit.add( AD2KitPeaceEdit("Kick", self.sibi, Pt(92, 371)) )
		kit.add( AD2KitPeaceEdit("Snare", self.sibi, Pt(262, 371)) )
		kit.add( AD2KitPeaceEdit("Hi Hat", self.sibi, Pt(432, 371)) )

		kit.add( AD2KitPeaceEdit("Tom 1", self.sibi, Pt(92, 243)) )
		kit.add( AD2KitPeaceEdit("Tom 2", self.sibi, Pt(262, 243)) )
		kit.add( AD2KitPeaceEdit("Tom 3", self.sibi, Pt(432, 243)) )
		kit.add( AD2KitPeaceEdit("Tom 4", self.sibi, Pt(602, 243)) )
		kit.add( AD2KitPeaceEdit("Ride 1", self.sibi, Pt(772, 243)) )
		kit.add( AD2KitPeaceEdit("Ride 2", self.sibi, Pt(942, 243)) )

		kit.add( AD2KitPeaceEdit("Flexi 1", self.sibi, Pt(602, 371)) )
		kit.add( AD2KitPeaceEdit("Flexi 2", self.sibi, Pt(772, 371)) )
		kit.add( AD2KitPeaceEdit("Flexi 3", self.sibi, Pt(942, 371)) )

		kit.add( AD2KitPeaceEdit("Cymbal 1", self.sibi, Pt(92, 115)) )
		kit.add( AD2KitPeaceEdit("Cymbal 2", self.sibi, Pt(262, 115)) )
		kit.add( AD2KitPeaceEdit("Cymbal 3", self.sibi, Pt(432, 115)) )
		kit.add( AD2KitPeaceEdit("Cymbal 4", self.sibi, Pt(602, 115)) )
		kit.add( AD2KitPeaceEdit("Cymbal 5", self.sibi, Pt(772, 115)) )
		kit.add( AD2KitPeaceEdit("Cymbal 6", self.sibi, Pt(942, 115)) )
		
		edit = FixedTab("Edit", self.sibi, Pt(968, 31), Pt( 999, 46), Pt(968, 38), 0x202327, 0x7fa7f5, ())
		tv.add(edit)
		edit.add( _AD2KitpeaceSelect("Kit", self.sibi, Pt(66, 211), (0x393642, 0x474d5c), 0x533e34) )
		edit.add( AD2KitPeaceEdit("", self.sibi, Pt(146, 113)) )
		
		ctrls = Container("Kitpeace controls", self.sibi, ("switchable",))
		edit.add( ctrls )
		ctrls.add( PushBtn("Kitpeace controls", self.sibi, Pt(393, 105), Pt(393, 105), 0x414754, 0xd7e4fe, ("say_enabled",)) )
		resp = Container("Response", self.sibi, ("switchable",))
		ctrls.add( resp )
		resp.add( PushBtn("Response", self.sibi, Pt(421, 105), Pt(421, 105), 0x3e2617, 0xfe8025, ("say_enabled",)) )
		resp.add( AD2Knob("Response volume", self.sibi, Pt(438, 146), ("%",)) )
		resp.add( AD2Knob("Response filter", self.sibi, Pt(438, 226), ("%",)) )
		resp.add( PushBtn("Alternate samples", self.sibi, Pt(471, 127), Pt(471, 127), 0x533e34, 0x49505f, ("say_enabled",)) )
		resp.add( YRange("Sample", self.sibi, Pt(478, 142), Pt(501, 230), 0x1a1f28, 0xb7c2da, ()) )

		pitch = Container("Pitch", self.sibi, ("switchable",))
		ctrls.add( pitch )
		pitch.add( PushBtn("Pitch", self.sibi, Pt(531, 105), Pt(531, 105), 0x3e2617, 0xfe8025, ("say_enabled",)) )
		pitch.add( AD2Knob("Main pitch", self.sibi, Pt(546, 146), ()) )
		pitch.add( AD2Knob("Overheads and room pitch offset", self.sibi, Pt(546, 226), ()) )

		tpd = FixedTabControl("Pitch envelope/Tone designer", self.sibi, ("skip_control",))
		ctrls.add( tpd )

		penv = FixedTab("Pitch envelope", self.sibi, Pt(725, 106), Pt( 725, 106), Pt(747, 236), 0x000000, 0x343a47, ("switchable",))
		tpd.add( penv )
		penv.add( GroupButton("Pitch envelope", self.sibi, Pt(588, 105), tpd, 0x3e2617, 0xfe8025) )
		penv.add( ScrollLabel("Start", self.sibi, Pt(584, 222), Pt(627, 236), ()))
		penv.add( ScrollLabel("Hold", self.sibi, Pt(640, 222), Pt(685, 236), ()))
		penv.add( ScrollLabel("Release", self.sibi, Pt(696, 222), Pt(741, 236), ()))
		penv.add( YBar("Pitch envelope velocity", self.sibi, Pt(757, 122), Pt(757, 234), 0x1a1f28, 0xb7c2da, ()) )

		td = FixedTab("Tone X designer", self.sibi, Pt(725, 106), Pt( 725, 106), Pt(747, 236), 0x343a47, 0x000000, ("switchable",))
		tpd.add( td )
		td.add( GroupButton("Tone designer", self.sibi, Pt(588, 105), tpd, 0x3e2617, 0xfe8025) )
		td.add( ScrollLabel("Start", self.sibi, Pt(584, 222), Pt(635, 236), ()))
		td.add( ScrollLabel("Decay", self.sibi, Pt(646, 222), Pt(699, 236), ()))
		td.add( ScrollLabel("End", self.sibi, Pt(710, 222), Pt(761, 236), ()))
		
		venv = Container("Volume envelope", self.sibi, ("switchable",))
		ctrls.add( venv )
		venv.add( PushBtn("Volume envelope", self.sibi, Pt(783, 105), Pt(783, 105), 0x3e2617, 0xfe8025, ("say_enabled",)) )
		venv.add( YBar("Volume envelope velocity", self.sibi, Pt(783, 122), Pt(783, 234), 0x1a1f28, 0xb7c2da, ()) )
		venv.add( ScrollLabel("Attack", self.sibi, Pt(799, 222), Pt(844, 236), ()))
		venv.add( ScrollLabel("Decay", self.sibi, Pt(854, 222), Pt(900, 236), ()))
		venv.add( ScrollLabel("Sustain level", self.sibi, Pt(911, 222), Pt(956, 236), ()))
		venv.add( ScrollLabel("Sustain time", self.sibi, Pt(967, 222), Pt(1011, 236), ()))
		venv.add( ScrollLabel("Release", self.sibi, Pt(1023, 222), Pt(1068, 236), ()))
		
		cut = Container("Cut", self.sibi, ("switchable",))
		ctrls.add( cut )
		cut.add( PushBtn("Cut", self.sibi, Pt(1090, 105), Pt(1090, 105), 0x3e2617, 0xfe8025, ("say_enabled",)) )

		ch = FixedTabControl("Channels", self.sibi, ("slow_reaction",))
		edit.add( ch );
		for name, x in ( ("Kick", 37), ("Snare", 108), ("HiHat", 179), ("Tom 1", 251), ("Tom 2", 324), ("Tom 3", 394), ("Tom 4", 467),
						 ("Flexi 1", 538), ("Flexi 2", 610), ("Flexi 3", 682), ("Overheads", 780), ("Room", 850), ("Bus", 927), ("Master", 1085) ):
			ch.add( _AD2ChannelTab(name, self.sibi, x) )
		
		sends = Container("Sends", self.sibi, ("switchable",))
		edit.add( sends )
		sends.add( PushBtn("Sends", self.sibi, Pt(1085, 426), Pt(1085, 426), 0x696b6e, 0xe4fefe, ("say_enabled",)) )

		inserts = Container("Inserts", self.sibi, ("switchable",))
		edit.add( inserts )
		inserts.add( PushBtn("Inserts", self.sibi, Pt(1085, 395), Pt(1085, 395), 0x696b6e, 0xe4fefe, ("say_enabled",)) )
		
		iv = FixedTabControl("Inserts view", self.sibi, ("skip_control",))
		inserts.add( iv )

		fiv = FixedTab("First inserts view", self.sibi, Pt(0, 0), Pt( 0, 0), Pt(43, 381), 0x585248, 0x03070e, ())
		iv.add( fiv )
		noise = Container("Noise", self.sibi, ("switchable",))
		fiv.add( noise )
		noise.add( PushBtn("Noise", self.sibi, Pt(159, 304), Pt(159, 304), 0x6a6b6f, 0xfee654, ("say_enabled",)) )
		cnd = Container("C&D", self.sibi, ("switchable",))
		fiv.add( cnd )
		cnd.add( PushBtn("Compression and distortion", self.sibi, Pt(229, 304), Pt(229, 304), 0x6a6b6f, 0xfee654, ("say_enabled",)) )
		eq = Container("EQ", self.sibi, ("switchable",))
		fiv.add( eq )
		eq.add( PushBtn("EQ", self.sibi, Pt(505, 304), Pt(505, 304), 0x6a6b6f, 0xfee654, ("say_enabled",)) )
		tns = Container("T&S", self.sibi, ("switchable",))
		fiv.add( tns )
		tns.add( PushBtn("Tape and shape", self.sibi, Pt(810, 304), Pt(810, 304), 0x6a6b6f, 0xfee654, ("say_enabled",)) )

		siv = FixedTab("Bus and master inserts view", self.sibi, Pt(0, 0), Pt( 0, 0), Pt(43, 381), 0x03070e, 0x585248, ())
		iv.add( siv )
		cnd = Container("C&D", self.sibi, ("switchable",))
		siv.add( cnd )
		cnd.add( PushBtn("Compression and distortion", self.sibi, Pt(88, 304), Pt(88, 304), 0x6a6b6f, 0xfee654, ("say_enabled",)) )
		eq = Container("EQ", self.sibi, ("switchable",))
		siv.add( eq )
		eq.add( PushBtn("EQ", self.sibi, Pt(363, 304), Pt(363, 304), 0x6a6b6f, 0xfee654, ("say_enabled",)) )
		tns = Container("T&S", self.sibi, ("switchable",))
		siv.add( tns )
		tns.add( PushBtn("Tape and shape", self.sibi, Pt(670, 304), Pt(670, 304), 0x6a6b6f, 0xfee654, ("say_enabled",)) )
		noise = Container("Noise", self.sibi, ("switchable",))
		siv.add( noise )
		noise.add( PushBtn("Noise", self.sibi, Pt(946, 304), Pt(946, 304), 0x6a6b6f, 0xfee654, ("say_enabled",)) )
		cut = Container("Cut", self.sibi, ("switchable",))
		siv.add( cut )
		cut.add( PushBtn("Cut", self.sibi, Pt(1016, 304), Pt(1016, 304), 0x6a6b6f, 0xfee654, ("say_enabled",)) )
		
		fx = FixedTab("FX", self.sibi, Pt(1026, 31), Pt( 1047, 46), Pt(1026, 38), 0x202327, 0x7fa7f5, ())
		tv.add(fx)

		fx1 = Container("FX1", self.sibi, ("switchable",))
		fx.add( fx1 )
		fx1.add( PushBtn("FX1", self.sibi, Pt(44, 102), Pt(44, 102), 0x696b6e, 0xe4fefe, ("say_enabled",)) )

		fx2 = Container("FX2", self.sibi, ("switchable",))
		fx.add( fx2 )
		fx2.add( PushBtn("FX2", self.sibi, Pt(44, 275), Pt(44, 275), 0x696b6e, 0xe4fefe, ("say_enabled",)) )
		
		#####
		dlg = Dialog("Map window", self.sibi, Pt(182, 46), 0x000000, 0x5d6576, set())
		self.sibi.addDialog(dlg)
		
		btn = PopupMenuButton("Preset", self.sibi, Pt(182,109), Pt(362, 126),
							Pt(220, 126), Pt(220, 0), 0x2b303a, (0x000000, 0x363c49),
							Pt(224, 126), Pt(317, 145), Pt(224, 146),
							0x2b303a, 0x4c5361, 0x6b4933, ("dynamic_text", "skip_spacers", "slow_open", "slow_reaction") )
		btn.type = ""
		btn.defineSubmenu( Pt(340, 0), 3, 6, 4, 0, 20, 20)
		dlg.add( btn )

		dlg.add( Clickable("Set as default", self.sibi, Pt(884,580), Pt(884,580), ("silent_action",)) ) 
		dlg.add( OpenBtn("OK", self.sibi, Pt(834, 621), Pt(834, 621), 0x000000, 0x687983, ("OK",)) )
		dlg.add( OpenBtn("Cancel", self.sibi, Pt(940, 625), Pt(940, 625), 0x000000, 0x687983, ("CANCEL",)) )

		dlg.add( Label("MIDI in", self.sibi, Pt(209, 174), Pt(360, 189), ("dynamic_text",)) )
		dlg.add( Label("MIDI out", self.sibi, Pt(209, 192), Pt(360, 207), ("dynamic_text",)) )

		dlg.add( SwitchBtn("HiHat CC reverse", self.sibi, Pt(410, 496), Pt(410, 496), 0x000000, 0x321b0f, 0xfe8025, ()) )
		dlg.add( _AD2LearnBtn("HiHat CC", self.sibi, Pt(380, 520)) )
		
		dlg.add( _AD2KitpeaceSelect("Kit", self.sibi, Pt(464, 206), (0x474d5a, 0x3e4552), 0x533e34) )

		learn = _AD2LearnBtn("", self.sibi, Pt(524, 278))
		list = OptionTable("Assignments", self.sibi, Pt(565, 264), Pt(710, 293), Pt(565, 296), Pt(565, 645), learn, ())
		list.setScroll(Pt(752, 265), Pt(752, 646), 0x000000, 0x3c424e)
		dlg.add(list)
		
		###
		dlg = Dialog("Save map", self.sibi, Pt(395, 274), 0x000000, 0xabb7d0, ())
		self.sibi.addDialog(dlg)
		dlg.add( Clickable("Name", self.sibi, Pt(426,262), Pt(783,282), ("dynamic_text", "silent_action")) ) 
		dlg.add( OpenBtn("OK", self.sibi, Pt(400, 369), Pt(400, 359), 0x131720, 0x687983, ("OK",)) )
		dlg.add( OpenBtn("Cancel", self.sibi, Pt(686, 361), Pt(686, 361), 0x000000, 0x687983, ("CANCEL",)) )

		###
		dlg = Dialog("Save preset", self.sibi, Pt(661, 166), 0x000000, 0x8a94aa, ())
		self.sibi.addDialog(dlg)
		dlg.add( Clickable("Name", self.sibi, Pt(414, 177), Pt(653, 195), ("dynamic_text", "silent_action")) ) 
		dlg.add( OpenBtn("OK", self.sibi, Pt(389, 364), Pt(389, 364), 0x131720, 0x687983, ("OK",)) )
		dlg.add( OpenBtn("Cancel", self.sibi, Pt(686, 361), Pt(686, 361), 0x000000, 0x687983, ("CANCEL",)) )
		
		######
		dlg = Dialog("Preset Browser", self.sibi, Pt(185, 72), 0x000000, 0x5d6677, set())
		self.sibi.addDialog(dlg)
		dlg.add( Clickable("Filter by name", self.sibi, Pt(863,109), Pt(976,125), ("dynamic_text", "silent_action")) ) 
		dlg.add( AD2PresetSpin("Preset", self.sibi, Pt(733, 494), Pt(990, 513), Pt(431, 550), Pt(471, 550), ("slow_reaction",)) )
		dlg.add( Clickable("Set as startup", self.sibi, Pt(930,550), Pt(930,550), ("silent_action",)) ) 
		dlg.add( OpenBtn("OK", self.sibi, Pt(490, 600), Pt(490,600), 0x000000, 0x687983, ("OK",)) )
		dlg.add( OpenBtn("Cancel", self.sibi, Pt(600, 600), Pt(600,600), 0x000000, 0x687983, ("CANCEL",)) )

		
		
class SonarPluginPane(IAccessible):
	"""
	Sonar plug-in windows have common toobox with presets, VST options, etc. Operatable
	with accessibility set.
	But there is no switch to move focus to the plug-in dialog itself.
	I do not completely understand how Sonar work with the focus there, nor I understand
	what NVDA tries to do with that. So some dirty method... Note that I have tried
	obj.setFocus and that was working wierd...
	"""

	name = "Toolbar"
	
	def _findPluginWindow(self):
		""" Should be the first child after 'CakewalkVSTWndClass' or the second after 'PlugWndClass' """
		container_seen = False
		pwin = self
		while pwin.firstChild:
			pwin = pwin.firstChild
			if pwin.role == controlTypes.ROLE_PANE and pwin.windowClassName == "CakewalkVSTWndClass":
				container_seen = True
			elif container_seen and pwin.role != controlTypes.ROLE_WINDOW:
				#log.error(pwin.devInfo)
				return pwin
		return None
		
	
	def event_gainFocus(self):
		pwin = self._findPluginWindow()
		if pwin and pwin.windowHandle and isinstance(pwin, SIBINVDA) and pwin.hadFocus():
			#log.error("->.."+str(pwin.windowHandle)+"..")
			eventHandler.queueEvent("gainFocus", pwin)			
		else:
			super(SonarPluginPane,self).event_gainFocus()
			
	def script_tab(self, gesture):
		""" Give focus to plug-in """
		pwin = self._findPluginWindow()
		if pwin:
			#pwin.setFocus()
			eventHandler.queueEvent("gainFocus", pwin)
		else:
			gesture.send()


	__gestures = {
		"kb:tab": "tab"
	}

# remember windows we do not recognize
gUnknownNI = {}
	
class AppModule(appModuleHandler.AppModule):

	def chooseNVDAObjectOverlayClasses(self, obj, clsList):
		"""
		Important that we get 2 (!) objects per window, one with ROLE_WINDOW and another with ROLE_XXX
		I am not sure what is better... but lets ignore ROLE_WINDOW
		"""
		global gUnknownNI
		if obj.role == controlTypes.ROLE_WINDOW:
			return
		if obj.windowClassName == "Afx:0000000140000000:8:0000000000010003:0000000000000006:0000000000010027":
			clsList.insert(0, SonarPluginPane)
			# log.error(" "+str(obj.windowHandle)+" . "+str(obj))
		elif "NIVSTChildWindow" in obj.windowClassName:
			# NI does not make it simple, classes are dynamic and nothing
			# exposed to detect what it is
			if obj.windowHandle in gUnknownNI:
				return # we have already checked it and have not found...
			prt = obj.parent
			while prt:
				if prt.windowClassName == 'Afx:0000000140000000:8:0000000000000000:0000000000000006:0000000000000000':
					break
				prt = prt.parent
			if prt:
				if "Guitar Rig" in prt.windowText:
					clsList.insert(0, SIBGuitarRig5)
					return
				if "Absynth 5" in prt.windowText:
					clsList.insert(0, Absynth5)
					return
			#log.error("NI plug-in detected, but not recognised: " + prt.windowText)
			gUnknownNI[obj.windowHandle] = True
		elif obj.windowText == "External VST Window":
			if obj.windowHandle in gUnknownNI:
				return # we have already checked it and have not found...
			prt = obj.parent
			while prt:
				if prt.windowClassName == 'Afx:0000000140000000:8:0000000000000000:0000000000000006:0000000000000000':
					break
				prt = prt.parent
			if prt:
				if "TL-64" in prt.windowText:
					clsList.insert(0, SIBTL64)
					return
				elif "SI-" in prt.windowText:
					if "SI-Electric" in prt.windowText:
						clsList.insert(0, SIElectricPiano)
						return
					elif "SI-Drum" in prt.windowText:
						clsList.insert(0, SIDrumKit)
						return
					elif "SI-Bass" in prt.windowText:
						clsList.insert(0, SIBassGuitar)
						return
					elif "SI-String" in prt.windowText:
						clsList.insert(0, SIStringSection)
						return
				elif "Dimension Pro" in prt.windowText:
					clsList.insert(0, DimensionPro)
					return
				elif "SessionDrummer" in prt.windowText:
					clsList.insert(0, SessionDrummer)
					return
					
			log.error("Some plug-in detected, but not recognised: " + prt.windowText)
			gUnknownNI[obj.windowHandle] = True
		elif obj.windowClassName == "#32770" and obj.windowText == "Program Browser":
			clsList.insert(0, SIProgramBrowser)
		elif "JUCE_" in obj.windowClassName:
			# some JUCE based plug-in
			prt = obj.parent
			while prt:
				if prt.windowClassName == 'PlugWndClass':
					break
				prt = prt.parent
			if prt:
				if "Addictive Drums" in prt.windowText:
					clsList.insert(0, AD2)
				else:
					pass
					log.error("Some JUCE plug-in: " + prt.windowText)
		else:
			pass
			#log.error(obj.windowText)
			#log.error(obj.role)
			
			
	def isBadUIAWindow(self,hwnd):
		"""
		If accessibility is enabled in Sonar, it tries to speak too much.
		AZ Controller does that better, so I prefer to switch UIA off
		"""
		return True
