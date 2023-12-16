# Single Image Blob Interface Accessible Control
# 
# EZ Drummer 2 overlay
#
# AZ (www.azslow.com), 2018
from ctypes import *
from ctypes.wintypes import *
import time
from logHandler import log
import winUser

from . import *


class EZD_TextOutBox(TextBox):
	""" Sometimes TextOut works, sometimes not... """
	def __init__(self, sibi, lt, rb, getShift = None):
		lt = sibi.ptScale(lt)
		rb = sibi.ptScale(rb)
		super(EZD_TextOutBox,self).__init__(sibi.hwnd, lt.x, lt.y, rb.x, rb.y, getShift)
	
	def getTextOutOrText(self, ch = None):
		box = self.getBox()
		text = box.getTextOut()
		if text == "":
			text = super(EZD_TextOutBox,self).getText(ch)
			text = text.replace("(Wli", "Cymbal") # who knows why... not every time
			#log.error("OCR: (%d, %d) - (%d, %d): '%s'", box.left, box.top, box.right, box.bottom, text)
		else:
			#log.error("TextOut: '%s'" % text)
			pass
		return text

	def findInX(self, x, bg, fg):
		box = self.getBox()
		x1, x2 = FindHSegment(self.hwnd, x, box.top, bg, fg)
		if x1 is None:
			return False
		self.left = x1
		self.right = x2
		return True
		
class EZD_Label(Control):
	def __init__(self, name, sibi, lt, rb, opt = None, getShift = None):
		super(EZD_Label,self).__init__(name, sibi, opt)
		self.box = EZD_TextOutBox(sibi, lt, rb, getShift)
		
	def getTextInfo(self):
		return (self.type, self.name, self.box.getTextOutOrText())
		
class EZD_PopupBox(EZD_TextOutBox):
	def __init__(self, sibi, top_center_pt, getShift = None):
		super(EZD_PopupBox,self).__init__(sibi, Pt(0, top_center_pt.y), Pt(0, top_center_pt.y + 13), self.getBoxShift)
		self._getShift = getShift
		self.tc = sibi.MXY(top_center_pt, getShift)
	
	def getBoxShift(self):
		dx, dy = self._getShift()
		return (0, dy)
	
	def isVisible(self):
		tc = self.tc.getXY()
		return self.findInX(tc.x, 0xe0e0a0, 0xffffbb)

	def getTextOutOrText(self):
		for attempt in range(6):
			time.sleep(0.05)
			if self.isVisible():
				text = super(EZD_PopupBox,self).getTextOutOrText()
				#text = super(EZD_PopupBox,self).getText()
				if text != "":
					text = text.replace("o", "0")
					text = text.replace("O", "0")
					return text
		return "No value"
	
class EZD_KitLayout(object):

	def __init__(self, sibi):
		self.timer = TimedCaller()
		self.sibi = sibi
		self.box = sibi.Box(Pt(14, 46), Pt(885, 495)) # the whole kit area
		self.shift_pt = sibi.ptScale(Pt(621, 107)) # reference point for One Shot Pad
		self.close_pt = sibi.ptScale(Pt(261, 8)) # shift from the right top to close button in the piece conf dialogs
		self.layouts = self.__load()
		self.layout = None # last known layout
		self.learning = False
		
		self.last_piece_idx = None # saved is isPieceCfgOpen, to avoid re-detection
	
	def __load(self):
		raw_layouts = Cache("ezd_kits", []).load()
		layouts = []
		for l in raw_layouts:
			for p in l["pieces"]:
				p["btn"] = Box(self.box.hwnd, *p["btn"])
				if "play" in p:
					p["play"] = MXY(self.box.hwnd, *p["play"])
			masks = []
			for m in l["mask"]:
				masks.append(RECT(*m))
			l["mask"] = tuple(masks)
			layouts.append(l)
		return layouts
	
	def __save(self):
		raw_layouts = []
		for l in self.layouts:
			masks = []
			for m in l["mask"]:
				masks.append((m.left, m.top, m.right, m.bottom))
			pieces = []
			for p in l["pieces"]:
				btn = p["btn"]
				piece =	{ "btn": (btn.left, btn.top, btn.right, btn.bottom), "name": p["name"], "cfg": p["cfg"] }
				if "play" in p:
					piece["play"] = (p["play"].x, p["play"].y)
				pieces.append( piece )
			raw_layouts.append( { "crc": l["crc"], "mask": tuple(masks), "pieces": pieces } )
		Cache("ezd_kits").save(raw_layouts)
	
	def __learnStart(self):
		self.detect()
		if self.layout is not None:
			try:
				self.layouts.remove(self.layout)
				self.layout = None
			except:
				pass
		self.timer.Stop()
		self.learning = True
		self.lbox = self.box.getBox()
		self.lpt = MXY(self.lbox.hwnd, self.lbox.left, self.lbox.top)
		self.dx = 30 # 30
		self.dy = 30 # 15
		self.delay = 20
		self.cfg_delay = 300
		self.play_delay = 100
		self.shift = False
		self.kit_pieces = []
		self.last_piece = None
		self.lcfg_idx = None
	
		# initialize
		self.lpt.moveTo()
		self.lpt.leftClick()
		time.sleep(self.delay/1000.)
		self.lbox.pictureChange(1) # initial, save reference
		self.lbox.pictureChange()  # also prepare normal comparison

		self.anim_check = 5 # check there is no animation

		self.timer.CallAfter(0, self.__learnCheckAnim)
	
	def __learnCheckAnim(self):
		code, changes = self.lbox.pictureChange()
		if code != 0:
			self.sibi.speak("The kit animation is not over. Please check the transport is stopped and try again.")
			self.learning = False
			return
		if self.anim_check <= 0:
			self.__learnPieceScanStep()
			return
		self.anim_check -= 1
		self.timer.CallAfter(200, self.__learnCheckAnim)
	
	def __learnPieceScanStep(self):
		if not self.learning:
			return
		code, changes = self.lbox.pictureChange()
		if code < 0:
			self.sibi.speak("Error during learning")
			self.learning = False
			return
		if code == 0: # no changes
			if self.last_piece is not None:
				self.last_piece["box"].includePoint(self.lpt)
		else: #1,2,3
			if self.last_piece is None: # piece baundary cross
				found = False
				for last_piece in self.kit_pieces:
					if last_piece["btn"] == changes:
						found = True
						break
				if found:
					last_piece["box"].includePoint(self.lpt)
					self.last_piece = last_piece
				else:
					self.last_piece = { "btn": changes, "box": Box(self.lbox.hwnd, self.lpt.x, self.lpt.y) }
					self.kit_pieces.append(self.last_piece)
					#changes.moveTo(changes.left, changes.bottom)
					#time.sleep(2)
					#self.lpt.moveTo()
					#time.sleep(0.1)
			elif code == 3: # end of piece and no other piece
				self.last_piece = None
			elif code == 1: # change from one piece to other, need initial picture for comparison
				self.last_piece = None
				self.lbox.moveTo(self.lbox.left, self.lbox.top)
				time.sleep(self.delay/1000.)
				self.lbox.pictureChange()
				self.lpt.moveTo()
				self.timer.CallAfter(self.delay, self.__learnPieceScanStep) # redo with current point, may be new piece
				return
			else: # code == 2, really no changes
				self.last_piece["box"].includePoint(self.lpt)
		self.lpt.x += self.dx
		if self.lpt.x > self.lbox.right:
			self.lpt.y += self.dy
			if self.lpt.y > self.lbox.bottom:
				self.sibi.speak("Still learning, please wait...")
				self.__learnPieceCfgStep()
				return
			self.lpt.x = self.lbox.left
			if self.shift:
				self.lpt.x += self.dx//2
			self.shift = not self.shift
		self.lpt.moveTo()
		self.timer.CallAfter(self.delay, self.__learnPieceScanStep) 

	def __learnPieceCfgStep(self):
		if not self.learning:
			return
		if self.lcfg_idx is  None:
			self.lcfg_idx = 0
			self.lcfg_open = 0
		elif self.lcfg_open == 2:
			self.lcfg_idx += 1
			self.lcfg_open = 0
		if self.lcfg_idx >= len(self.kit_pieces):
			self.lplay_idx = 0
			self.lplay_step = 0
			self.lplay_direction = 0
			self.sibi.speak("Testing kit pieces...")
			self.__learnPiecePlayPoint()
			return
		piece = self.kit_pieces[self.lcfg_idx]
		btn = piece["btn"]
		if self.lcfg_open == 0:
			self.lcfg_open = 1
			btn.leftClick()
		elif self.lcfg_open == 1:
			x = (btn.left + btn.right) // 2
			y = btn.top - 100
			if y < self.lbox.top:
				y = self.lbox.top
			while True:
				left, right = FindHSegment(btn.hwnd, x, y, (0xa09d9a, 0xa09b9a), 0xa09c9a) # exact color
				if left is not None and (right - left) > 265: # should be 269
					break;
				y += 1
				if y >= self.lbox.bottom:
					break;
			if left is None:
				log.error("No config on (%d,%d - %d)" % (x, btn.top + 1, y))
				self.sibi.speak("Can not find kit piece configuration dialog")
				self.learning = False
				return
			# find height, percussion can not select EZX and so are smaller
			l, r = FindHSegment(btn.hwnd, x, y + 180, (0x645e59, 0x645e5b), 0x645e5a)
			if l is None or (r - l) < 260:
				l, r = FindHSegment(btn.hwnd, x, y + 218, (0x645e59, 0x645e5b), 0x645e5a)
				if l is None or (r - l) < 260:
					self.sibi.speak("Unknown height for kit piece configuration dialog")
					self.learning = False
					return
				height = 218
			else:
				height = 180
					
			# fine peace name
			getShift = lambda: (left - self.shift_pt.x, y - self.shift_pt.y)
			name_box = EZD_TextOutBox(self.sibi, Pt(620,109), Pt(874, 126), getShift)
			name = name_box.getTextOutOrText()
			if name == "":
				box = name_box.getBox()
				log.error("No name at (%d,%d)-(%d,%d) : %d, %d" % (box.left, box.top, box.right, box.bottom, left, y))
				self.sibi.speak("Can not find kit piece name")
				self.learning = False
				return
			name = name.replace("(Wli", "Cymbal") # who knows why... not every time
			name = name.replace("acktom", "ack tom")
			name = name.replace("loortom", "loor tom")
			piece["name"] = name
			box = name_box.getBox()
			#log.error("%s %d %d " % (name, box.top, box.bottom))
			piece["cfg"] = (left, y, height)
			piece["box"] = piece["box"].resize(-self.dx-10, -self.dy-10, self.dx+10, self.dy+10, self.lbox)
			if name.find("Hat") > 0:
				piece["box"].bottom = self.lbox.bottom # hack, hi-hat pedal can not be recognized with low resolution
				piece["box"].right += 20 # that I have not understood...
				piece["box"].left -= 20
				piece["box"].top -= 20
			close = MXY(self.lbox.hwnd, left + self.close_pt.x, y + self.close_pt.y)
			close.leftClick()
			self.lcfg_open = 2
		self.timer.CallAfter(self.cfg_delay, self.__learnPieceCfgStep)

	def __learnWaitStablePicture(self):
		code, changes = self.lbox.pictureChange()
		if code == 0:
			if self.anim_check <= 0:
				self.__learnPiecePlayPoint()
				return
			self.anim_check -= 1
		else:
			self.anim_check = 5
		self.timer.CallAfter(self.play_delay, self.__learnWaitStablePicture)

	def __learnPiecePlayPoint(self):
		if not self.learning:
			return
		if self.lplay_idx >= len(self.kit_pieces):
			self.__learnFinish()
			return
		piece = self.kit_pieces[self.lplay_idx]
		btn = piece["btn"]

		if self.lplay_step == 0:
			# move to zero and wait stable (initial) picture
			self.lbox.moveTo(self.lbox.left, self.lbox.top)
			self.anim_check = 5
			self.lplay_step = 1
			self.lbox.pictureChange()
			self.timer.CallAfter(self.play_delay, self.__learnWaitStablePicture)
		elif self.lplay_step == 1:
			# we have stable initial picture, move to potential point, wait for changes
			if self.lplay_direction == 0:
				y = piece["box"].top + 10 # original detected top
				if piece["name"].find("nare") >= 0 and btn.top > y: 
					# try middle for Snare
					self.play_pt = MXY(btn.hwnd, (btn.left + btn.right)//2, y + (btn.top - y)//2)	
				else:
					self.lplay_direction == 1
					self.play_pt = MXY(btn.hwnd, (btn.left + btn.right)//2, btn.top - 4)
			elif self.lplay_direction == 1:
				self.play_pt = MXY(btn.hwnd, (btn.left + btn.right)//2, btn.top - 4)
			elif self.lplay_direction == 2:
				self.play_pt = MXY(btn.hwnd, (btn.left + btn.right)//2, btn.bottom + 4)
			elif self.lplay_direction == 3:
				self.play_pt = MXY(btn.hwnd, btn.left - 4, (btn.top + btn.bottom)//2)
			elif self.lplay_direction == 4:
				self.play_pt = MXY(btn.hwnd, btn.right + 4, (btn.top + btn.bottom)//2)
			elif self.lplay_direction == 5:
				self.play_pt = MXY(btn.hwnd, btn.left - 4, btn.top - 4)
			elif self.lplay_direction == 6:
				self.play_pt = MXY(btn.hwnd, btn.right + 4, btn.top - 4)
			elif self.lplay_direction == 7:
				self.play_pt = MXY(btn.hwnd, btn.right + 4, btn.bottom + 4)
			elif self.lplay_direction == 8:
				self.play_pt = MXY(btn.hwnd, btn.left - 4, btn.bottom + 4)
			else:
				self.sibi.speak("Error. Can not find %s play point, please contact Sibiac developers" % piece["name"])
				self.learning = False
				return
			self.play_pt.moveTo()
			self.lplay_step = 2
			self.timer.CallAfter(self.delay, self.__learnPiecePlayPoint)
		elif self.lplay_step == 2:
			# we should get our btn as the change
			self.lplay_step = 0
			code, changes = self.lbox.pictureChange()
			if code < 1 or not changes == btn:
				log.error(code)
				log.error(changes)
				log.error(btn)
				self.lplay_direction += 1
			else: # ok, we are within our piece
				count = 30
				self.play_pt.leftClick()
				while count > 0:
					code, changes = self.lbox.pictureChange()
					if code != 0:
						break;
					time.sleep(0.005)
					count -= 1
				if count == 0: # no reaction, try next direction
					self.lplay_direction += 1
				else: # found play point
					piece["play"] = self.play_pt 
					self.lplay_idx += 1
					self.lplay_direction = 0
			self.__learnPiecePlayPoint()
	
	__piece_order = [ "ick", "nare", "Hat", "tom", "ide", "ymb", "hina", "" ]
	
	__piece_order_name = [ "Kick", "Snare", "Hi-Hat", "Tom", "Ride", "Cymbal", "China", "Extra pieces" ]
	
	def __learnOrderPieces(self, pieces):
		""" Order pieces to make them homogenious over layouts """
		o = {}
		for po in self.__piece_order:
			o[po] = []
		for p in pieces:
			for po in self.__piece_order:
				if p["name"].find(po) >= 0:
					o[po].append( p )
					break
		for i, it in o.items():
			it.sort(key = lambda x: x["name"])
		sp = []
		for po in self.__piece_order:
			sp.extend(o[po])
		return sp
	
	def __learnFinish(self):
		if not self.learning:
			return
		self.timer.Stop()
		if len(self.kit_pieces) == 0:
			self.sibi.speak("Could not detect any kit peace")
			return
		# construct layout
		p = []
		m = []
		for piece in self.kit_pieces:
			p.append( { "btn": piece["btn"], "name": piece["name"], "cfg": piece["cfg"], "play": piece["play"] } )
			m.append( piece["box"].RECT() )
		box = self.box.getBox()
		box.moveTo(box.left, box.top)
		time.sleep(self.cfg_delay/1000.)
		crc = box.pictureCRC(m)
		p = self.__learnOrderPieces(p)
		self.layout = { "pieces": p, "mask": m, "crc": crc }
		self.layouts.append( self.layout )
		self.__save()
		self.sibi.speak("Found %d pads. Learning is complete." % len(self.kit_pieces))
		self.learning = False

	def __cancelLearn(self):
		if self.learning:
			self.timer.Stop()
			self.learning = False
		
	def detect(self):
		""" Return True if the layout is changed """
		self.__cancelLearn()
		if self.last_piece_idx is not None and self.isPieceCfgOpen(self.last_piece_idx):
			return False
		self.last_piece_idx = None
		box = self.box.getBox()				
		if self.layout is not None and box.pictureCRC(self.layout["mask"]) == self.layout["crc"]:
			return False
		for i in range(5): # well, I need better tool to understand from where it comes...
			for layout in self.layouts:
				if box.pictureCRC(layout["mask"]) == layout["crc"]:
					if self.layout == layout:
						return False
					self.layout = layout
					return True
			box.leftClick(box.left, box.top)
			time.sleep(0.05)
		if self.layout is None:
			return False
		self.layout = None
		return True
	
	def getPiece(self, i):
		""" Note that detection is not automatically called """
		self.__cancelLearn()
		if self.layout is None:
			return None
		pieces = self.layout["pieces"]
		if i < 0 or i >= len(pieces):
			return None
		return pieces[i]

	def getCategoryIdx(self, i):
		self.__cancelLearn()
		try:
			i -= 1 # I count from 1 to match keyboard keys
			pieces = self.layout["pieces"]
			key = self.__piece_order[i]
		except: # No layout or unknown category 
			return None
		if key == "": # other category
			other_idx = None
			for i, p in reversed(list(enumerate(pieces))):
				for k in self.__piece_order:
					if k != "" and p["name"].find(k) >= 0:
						return other_idx
				other_idx = i
			return other_idx
		else:
			for i, p in enumerate(pieces):
				if p["name"].find(key) >= 0:
					return i
		return None
		
	def isPieceCfgOpen(self, i):
		piece = self.getPiece(i)
		if piece is None:
			return False
		cfg = piece["cfg"]
		for t in range(3): # during picture update, there are "blackout" times
			left, right = FindHRange(self.box.hwnd, cfg[0], cfg[0]+100, cfg[1], (0xa09d9a, 0xa09b9a), 0xa09c9a) # exact color
			if left is not None and (right - left) > 90:
				self.last_piece_idx = i
				return True
			time.sleep(0.01)
		return False

	def pieceCfgClose(self, i):
		if not self.isPieceCfgOpen(i):
			return
		piece = self.getPiece(i)
		cfg = piece["cfg"]
		close = MXY(self.box.hwnd, cfg[0] + self.close_pt.x, cfg[1] + self.close_pt.y)
		close.leftClick()
		
	def getPieceMask(self, i):
		self.__cancelLearn()
		if self.layout is None:
			return None
		mask = self.layout["mask"]
		if i < 0 or i >= len(mask):
			return None
		r = mask[i]
		return Box(self.box.hwnd, r.left, r.top, r.right, r.bottom) # convert from RECT
		
	def learn(self):
		self.sibi.speak("Learning has started. Please do not use your keyboard until Sibiac informs you that learning is complete. This will take a few minutes.")
		self.__learnStart()
		
	def stopLearning(self):
		self.__cancelLearn()
		
	def getNumberOfPieces(self):
		self.__cancelLearn()
		if self.layout is None:
			return 0
		return len(self.layout["pieces"])
	
	def moveLeft(self, i):
		self.__cancelLearn()
		if self.layout is None or i <= 0:
			return False
		pieces = self.layout["pieces"]
		if i >= len(pieces):
			return False
		pieces.insert(i - 1, pieces.pop(i))
		self.__save()
		return True

	def moveRight(self, i):
		self.__cancelLearn()
		if self.layout is None or i < 0:
			return False
		pieces = self.layout["pieces"]
		if i >= len(pieces) - 1:
			return False
		pieces.insert(i + 1, pieces.pop(i))
		self.__save()
		return True
	
	def getCategoryName(self, i):
		try:
			return self.__piece_order_name[i - 1]
		except:
			return "Unknown"

class EZD_PieceCfg_Control(Control):
	def __init__(self, kit, name = "", opt = None):
		super(EZD_PieceCfg_Control,self).__init__(name, kit.sibi, opt)
		self.kit = kit
		self.pshift_pt = kit.sibi.ptScale(Pt(621, 107)) # reference point for One Shot Pad
		self.cshift_pt = kit.sibi.ptScale(Pt(429, 156)) # reference point for Cymbal 3 in Modern pack
	
	def _setCfg(self):
		""" set is_percussion and coordinates for getShift """
		piece = self.kit.getPiece()
		if piece is None:
			self.x = 0
			self.y = 0
			self.is_percussion = False
			return False
		cfg = piece["cfg"]
		self.is_percussion = cfg[2] < 200
		if self.is_percussion:
			self.x = cfg[0] - self.pshift_pt.x
			self.y = cfg[1] - self.pshift_pt.y
		else:
			self.x = cfg[0] - self.cshift_pt.x
			self.y = cfg[1] - self.cshift_pt.y
		return True
		
	def getShift(self):
		return (self.x, self.y)
		
	def focusSet(self):
		self.sibi.focusChangedTo(self, False) # focusLost is normally translated with some clicks... not good when entering dialogs
		
	def onEscape(self):
		self.kit.closePieceCfg()
		self.speakFocusAfter()
		return True
		
class EZD_PieceCfg_Title(EZD_PieceCfg_Control):
	def __init__(self, kit):
		super(EZD_PieceCfg_Title,self).__init__(kit)
		sibi = kit.sibi
		self.package = EZD_TextOutBox(sibi, Pt(435, 186), Pt(613, 197), self.getShift)
	
	def getTextInfo(self):
		piece = self.kit.getPiece()
		if piece is None:
			return ("", "Unknown piece configuration dialog", "")
		text = ""
		if self._setCfg() and not self.is_percussion:
			text = ", Package " + self.package.getTextOutOrText(2)
		return ("Dialog", piece["name"], text)
	
	def onEnter(self):
		if not self._setCfg() or self.is_percussion:
			self.speakFocusAfter()
		else:
			self.kit.expectFocusReturn()
			self.package.getBox().leftClick()
		return True
	
class EZD_Piece_Selector(EZD_PieceCfg_Control):
	def __init__(self, kit):
		super(EZD_Piece_Selector,self).__init__(kit, "Piece", ("delayed_reaction",))
		sibi = kit.sibi
		self.p_piece_box = EZD_TextOutBox(sibi, Pt(667, 254), Pt(825, 267), self.getShift)
		self.c_piece_box = EZD_TextOutBox(sibi, Pt(474, 332), Pt(639, 365), self.getShift)
		
		self.p_up_btn = sibi.MXY(Pt(644, 252), self.getShift)
		self.p_down_btn = sibi.MXY(Pt(646, 271), self.getShift)
		self.c_up_btn = sibi.MXY(Pt(450, 340), self.getShift)
		self.c_down_btn = sibi.MXY(Pt(451, 359), self.getShift)

		self.p_list_item = EZD_TextOutBox(sibi, Pt(626, 0), Pt(825, 0), self.getItemShift)
		self.p_list = sibi.Box(Pt(626, 137), Pt(825, 229), self.getShift)
		self.c_list_item = EZD_TextOutBox(sibi, Pt(433, 0), Pt(633, 0), self.getItemShift)
		self.c_list = sibi.Box(Pt(433, 209), Pt(633, 316), self.getShift)
	
		self.piece_box = None
		self.up_btn = None
		self.down_btn = None
		self.list = None
		self.list_item = None
		
		self.play_volume = 50 # in percent
	
	def _setCfg(self):
		if not super(EZD_Piece_Selector,self)._setCfg():
			return False
		if self.is_percussion:
			self.piece_box = self.p_piece_box
			self.up_btn = self.p_up_btn
			self.down_btn = self.p_down_btn
			self.list = self.p_list
			box = self.p_list.getBox()
			self.list_item = self.p_list_item
		else:
			self.piece_box = self.c_piece_box
			self.up_btn = self.c_up_btn
			self.down_btn = self.c_down_btn
			self.list = self.c_list
			box = self.c_list.getBox()
			self.list_item = self.c_list_item
		box = self.list.getBox()
		y0, y1 = FindVRange(box.hwnd, box.left, box.top, box.bottom, (0x24201e, 0xa59993), 0x918782)
		if y1 - y0 > 10:
			self.list_item.top = y0
			self.list_item.bottom = y1
		else:
			self.list_item.top = 0
			self.list_item.bottom = 0
		return True
	
	def getItemShift(self):
		dx, dy = self.getShift()
		return dx, 0 # y is set from item detection
		
	def play(self, dvolume = 0):
		self.play_volume += dvolume
		if self.play_volume > 100:
			self.play_volume = 100
		elif self.play_volume < 30:
			self.play_volume = 30
		if not self._setCfg():
			return False
		box = self.piece_box.getBox()
		MXY(box.hwnd, box.left + (box.right - box.left)*self.play_volume//100, (box.top + box.bottom)//2).leftClick()
		return True
	
	def getTextInfo(self):
		if not self._setCfg():
			return (self.type, self.name, "Unknown")
		if self.list_item.top != self.list_item.bottom:
			text = self.list_item.getTextOutOrText()
		else:
			text = self.piece_box.getTextOutOrText() + ", not in selected package"
		return (self.type, self.name, text)
		
	def onUp(self):
		if self._setCfg():
			if self.is_percussion or self.list_item.top != self.list_item.bottom:
				self.up_btn.leftClick()
			else:
				list = self.list.getBox()
				list.leftClick(list.left, list.top)
		self.speakAfter()
		return True
		
	def onDown(self):
		if self._setCfg():
			if self.is_percussion or self.list_item.top != self.list_item.bottom:
				self.down_btn.leftClick()
			else:
				list = self.list.getBox()
				list.leftClick(list.left, list.top)
		self.speakAfter()
		return True

	def onLeft(self):
		self.play(-10)
		return True

	def onRight(self):
		self.play(+10)
		return True

	def onEnter(self):
		self.play()
		return True
	
	def onHelp(self, prefix = "", suffix = ""):
		self.speak(prefix + "Select preset using up and down arrows. Press Enter to play current preset, use left and right arrows to play quiter or louder hits." + suffix)
		return True
	
class EZD_Piece_Knob(EZD_PieceCfg_Control):
	def __init__(self, kit, name, p_ctrl_pt, c_ctrl_pt, selector):
		super(EZD_Piece_Knob,self).__init__(kit, name)
		self.selector = selector
		sibi = kit.sibi
		self.p_ctrl_pt = sibi.MXY(p_ctrl_pt, self.getShift)
		self.p_label = EZD_PopupBox(sibi, Pt(p_ctrl_pt.x, p_ctrl_pt.y - 65), self.getShift)
		
		self.c_ctrl_pt = sibi.MXY(c_ctrl_pt, self.getShift)
		self.c_label = EZD_PopupBox(sibi, Pt(c_ctrl_pt.x, c_ctrl_pt.y - 65), self.getShift)

		self.ctrl_pt = None
		self.label = None

	def _setCfg(self):
		if not super(EZD_Piece_Knob,self)._setCfg():
			return False
		if self.is_percussion:
			self.ctrl_pt = self.p_ctrl_pt
			self.label = self.p_label
		else:
			self.ctrl_pt = self.c_ctrl_pt
			self.label = self.c_label
		return True
		
	def getTextInfo(self):
		if not self._setCfg():
			text = "Unknown"
		else:
			self.ctrl_pt.leftClick()
			text = self.label.getTextOutOrText()
		return (self.type, self.name, text)
	
	def onUp(self):
		if not self._setCfg():
			return True
		self.ctrl_pt.moveTo()
		MouseScroll(120)
		self.speakAfter()
		return True

	def onDown(self):
		if not self._setCfg():
			return True
		self.ctrl_pt.moveTo()
		MouseScroll(-120)
		self.speakAfter()
		return True

	def onLeft(self):
		self.selector.play(-10)
		return True

	def onRight(self):
		self.selector.play(+10)
		return True

	def onEnter(self):
		self.selector.play()
		return True

	def onHelp(self, prefix = "", suffix = ""):
		self.speak(prefix + "Use up and down arrows to change the value. Do this with Shift for fine adjustment." +
				   " Press Enter to play current preset, use left and right arrows to play quiter or louder hits." + suffix)		
		return True
		
class EZD_Piece_KitsMenu(EZD_PieceCfg_Control):
	def __init__(self, kit):
		super(EZD_Piece_KitsMenu,self).__init__(kit, "Kits")
		self.type = "Menu"
		self.pt = kit.sibi.MXY(Pt(672, 335), self.getShift)

	def isFocusable(self):
		return self._setCfg() and not self.is_percussion
		
	def onEnter(self):
		if self._setCfg() and not self.is_percussion:
			self.kit.expectFocusReturn()
			self.pt.leftClick()
		return True
		
class EZD_PieceCfg(Container):
	def __init__(self, kit):
		super(EZD_PieceCfg,self).__init__("", kit.sibi, None)
		self.kit = kit
		self.add( EZD_PieceCfg_Title(kit) )
		selector = EZD_Piece_Selector(kit)
		self.add( selector )
		self.add( EZD_Piece_Knob(kit, "Volume", Pt(865, 162), Pt(673, 216), selector) )
		self.add( EZD_Piece_Knob(kit, "Pitch", Pt(865, 219), Pt(673, 280), selector) )
		self.add( EZD_Piece_KitsMenu(kit) )
	
	def play(self, dvolume=0):
		return self.selector.play(dvolume)
		
	def isFocusable(self):
		return self.kit.isPieceCfgOpen()
		
	def focusNext(self):
		if not super(EZD_PieceCfg,self).focusNext():
			return self.focusFirst()
		return True
		
	def focusPrevious(self):
		if not super(EZD_PieceCfg,self).focusPrevious():
			return self.focusLast()
		return True
	
class EZD_Kit(Control):
	def __init__(self, sibi):
		super(EZD_Kit,self).__init__("Kit", sibi, ("slow_reaction",))
		self.layout = EZD_KitLayout(sibi)
		self.idx = 0
		self.piece_cfg = EZD_PieceCfg(self)
		self.confirm = 0
	
	def getPiece(self):
		return self.layout.getPiece(self.idx)
	
	def isPieceCfgOpen(self):
		return self.layout.isPieceCfgOpen(self.idx)
	
	def _pieceCfgBtnClick(self):
		piece = self.getPiece()
		if piece is not None:
			piece["btn"].getBox().leftClick()
		self.speakFocusAfter()
	
	def openPieceCfg(self):
		if self.isPieceCfgOpen():
			return
		self.piece_cfg.ctrl_idx = 0
		self._pieceCfgBtnClick()

	def closePieceCfg(self):
		self.layout.pieceCfgClose(self.idx)
		
	def getTextInfo(self):
		piece = self.layout.getPiece(self.idx)
		if piece is None:
			return (self.type, self.name, " layout unknown. Make sure that your transport is stopped, then press enter to start learning.")
		return (self.type, self.name, piece["name"])
	
	def getFocused(self):
		if self.piece_cfg.isFocusable():
			return self.piece_cfg.getFocused()
		return self;
		
	def focusSet(self):
		self.confirm = 0
		if self.layout.detect():
			self.idx = 0
			self.piece_cfg.ctrl_idx = 0
		if self.piece_cfg.isFocusable():
			self.piece_cfg.focusSet()
		else:
			super(EZD_Kit,self).focusSet()
			
	def focusFirst(self):
		if self.piece_cfg.isFocusable():
			return self.piece_cfg.focusFirst()
		return super(EZD_Kit,self).focusFirst()

	def focusLast(self):
		if self.piece_cfg.isFocusable():
			return self.piece_cfg.focusLast()
		return super(EZD_Kit,self).focusLast()

	def focusNext(self):
		if self.piece_cfg.isFocusable():
			return self.piece_cfg.focusNext()
		return super(EZD_Kit,self).focusNext()

	def focusPrevious(self):
		if self.piece_cfg.isFocusable():
			return self.piece_cfg.focusPrevious()
		return super(EZD_Kit,self).focusPrevious()

	def focusLost(self):
		self.confirm = 0
		self.closePieceCfg()
	
	def _confirmDialog(self, confirm = None):
		if self.confirm == 0 and confirm is None:
			return False
		self.confirm = confirm
		if confirm == 1:
			self.speak("To confirm re-learning current layout press Enter. Press Escape to cancel the operation.")
		else:
			self.speak("To confirm forgetting all learned layouts press Enter. Press Escape to cancel the operation.")
		return True
	
	def onEscape(self):
		if self.confirm == 0:
			return False
		self.confirm = 0
		self.speakFocusAfter(0)
		return True
	
	def onRight(self):
		if self._confirmDialog():
			return True
		idx = self.idx + 1
		piece = self.layout.getPiece(idx)
		if piece is not None:
			self.idx = idx
		self.speakAfter(0)
		return True

	def onLeft(self):
		if self._confirmDialog():
			return True
		idx = self.idx - 1
		piece = self.layout.getPiece(idx)
		if piece is not None:
			self.idx = idx
		self.speakAfter(0)
		return True

	def onUp(self):
		if self._confirmDialog():
			return True
		self.openPieceCfg()
		#mask = self.layout.getPieceMask(self.idx)
		#if mask is not None:
		#	mask.moveTo(mask.left, mask.top)
		return True

	def onDown(self):
		if self._confirmDialog():
			return True
		self.openPieceCfg()
		#mask = self.layout.getPieceMask(self.idx)
		#if mask is not None:
		#	mask.moveTo(mask.right, mask.bottom)
		return True
		
	def onEnter(self):
		if self.confirm == 2:
			self.confirm = 0
			self.layout.layout = None
			self.layout.layouts = []
			self.layout.learn()
		elif self.confirm == 1:
			self.confirm = 0
			self.layout.learn()
		else:
			piece = self.layout.getPiece(self.idx)
			if piece is None:
				self.layout.learn()
			else:
				if "play" in piece:
					piece["play"].leftClick()
				else:
					box = piece["btn"].getBox()
					box.leftClick((box.left + box.right)//2, box.top - 2)
		return True

	def onCtrlShiftR(self):
		self._confirmDialog(2)	
		return True

	def onCtrlShiftL(self):
		self._confirmDialog(1)
		return True

	def onCtrlShiftLeft(self):
		if self._confirmDialog():
			return True
		piece = self.layout.getPiece(self.idx)
		if piece is None:
			self.speak("No kit piece")
		elif self.layout.moveLeft(self.idx):
			piece2 = self.layout.getPiece(self.idx)
			self.idx -= 1
			self.speak("%s is before %s now" % (piece["name"], piece2["name"]))
		else:
			self.speak("%s is the first pad now" % piece["name"])
		return True

	def onCtrlShiftRight(self):
		if self._confirmDialog():
			return True
		piece = self.layout.getPiece(self.idx)
		if piece is None:
			self.speak("No kit piece")
		elif self.layout.moveRight(self.idx):
			piece2 = self.layout.getPiece(self.idx)
			self.idx += 1
			self.speak("%s is after %s now" % (piece["name"], piece2["name"]))
		else:
			self.speak("%s is the last pad now" % piece["name"])
		return True

	def _onNum(self, i):
		if self._confirmDialog():
			return True
		idx = self.layout.getCategoryIdx(i)
		if idx is not None:
			self.idx = idx
			self.speakAfter(0)
		else:
			self.speak("No %s in this kit" % self.layout.getCategoryName(i))
		return True
	
	def onNum1(self):
		return self._onNum(1)

	def onNum2(self):
		return self._onNum(2)

	def onNum3(self):
		return self._onNum(3)

	def onNum4(self):
		return self._onNum(4)

	def onNum5(self):
		return self._onNum(5)

	def onNum6(self):
		return self._onNum(6)

	def onNum7(self):
		return self._onNum(7)

	def onNum8(self):
		return self._onNum(8)

	def onHelp(self, prefix = "", suffix = ""):
		self.speak(prefix + "Kit configuration page. Select kit piece using left and right arrows. You can navigate to the first piece in defined groups using keys from 1 to 8." +
			       " To change the position of current piece use Control+Shift+Left or Right arrow."
				   " Press enter to play current piece. Press down arrow to open" +
			       " piece configuration dialog. You can re-learn current kit layout pressing Control+Shift+L. You can also forget all learned layouts pressing Control+Shift+R" + suffix)
		return True
	
class EZD_Dialog_Name(Control):
	def __init__(self, name, dialog, opt = None):
		super(EZD_Dialog_Name,self).__init__(name, dialog.sibi, opt)
		self.dialog = dialog
		
	def onEscape(self):
		return self.dialog.ctrl[len(self.dialog.ctrl) - 1].onEnter()

	def onHelp(self, prefix = "", suffix = ""):
		self.speak(prefix + "Interface internal dialog name. Press escape to close the dialog." + suffix)
		return True

class EZD_About(Container):
	
	def __init__(self, sibi):
		super(EZD_About,self).__init__("About", sibi, None)
		self.add( EZD_Dialog_Name("About dialog", self) )
		self.add( EZD_Label("", sibi, Pt(264, 376), Pt(619, 392)) )
		self.add( Clickable("Copy Text", sibi, Pt(608, 359), Pt(608, 359), ("silent_action",)) )
		self.add( CloseBtn("Close", sibi, Pt(640, 206), Pt(640, 206), None) )

class EZD_Update_Dialog(Container):
	
	def __init__(self, sibi):
		super(EZD_Update_Dialog,self).__init__("Update", sibi, None)
		self.add( EZD_Dialog_Name("Update dialog", self) )
		self.add( EZD_Label("", sibi, Pt(322, 330), Pt(643, 356)) )
		self.add( CloseBtn("Close", sibi, Pt(638, 233), Pt(638, 233), None) )
		
class EZD_CheckBox(Control):
	def __init__(self, name, sibi, check_pt, check_colors):
		super(EZD_CheckBox,self).__init__(name, sibi, None)
		self.type = "Checkbox"
		self.check_pt = sibi.MXY(check_pt)
		self.check_colors = Color2Tuple(check_colors)
		
	def _isChecked(self):
		return self.check_pt.FindNearestColor(self.check_colors) == 1
		
	def getTextInfo(self):
		if self._isChecked():
			text = "Checked"
		else:
			text = "Not checked"
		return (self.type, self.name, text)
		
	def onEnter(self):
		self.check_pt.leftClick()
		self.speakAfter()
		return True
		
	def onHelp(self, prefix = "", suffix = ""):
		self.speak(prefix + "Press enter to toggle checkbox" + suffix)
		return True

class EZD_Dlg_CheckBox(EZD_CheckBox):
	def __init__(self, name, dialog, check_pt, check_colors, help = ""):
		super(EZD_Dlg_CheckBox,self).__init__(name, dialog.sibi, check_pt, check_colors)
		self.dialog = dialog
		self.help = help

	def onEscape(self):
		return self.dialog.ctrl[len(self.dialog.ctrl) - 1].onEnter()

	def onHelp(self, prefix = "", suffix = ""):
		self.speak(prefix + self.help + " Press enter to toggle" + suffix)
		return True
		
class EZD_Dlg_DropList(EZD_Label):
	def __init__(self, name, dialog, lt, rb, help):
		super(EZD_Dlg_DropList,self).__init__(name, dialog.sibi, lt, rb)
		self.dialog = dialog
		self.help = help

	def onEnter(self):
		self.expectFocusReturn()
		self.box.leftClick()
		return True
		
	def onEscape(self):
		return self.dialog.ctrl[len(self.dialog.ctrl) - 1].onEnter()

	def onHelp(self, prefix = "", suffix = ""):
		self.speak(prefix + self.help + " Press enter to open the list of possible values." + suffix)
		return True

class EZD_Settings_Pages(TabGroup):
	def __init__(self, dialog):
		super(EZD_Settings_Pages,self).__init__("page", dialog.sibi, ("slow_reaction",))
		self.type = "Settings"
		self.dialog = dialog

	def onEscape(self):
		return self.dialog.ctrl[len(self.dialog.ctrl) - 1].onEnter()

	def onHelp(self, prefix = "", suffix = ""):
		self.speak(prefix + "Use left and right arrows to select settings page" + suffix)
		return True
		
class EZD_Settings_Page(Container):
	def __init__(self, name, dialog, check_pt):
		super(EZD_Settings_Page,self).__init__(name, dialog.sibi, None)
		self.dialog = dialog
		self._check_pt = dialog.sibi.MXY(check_pt)
		self._check_colors = (0xc2bab5, 0xece7e4, 0xe4dcd8)
	
	def isActive(self):
		return self._check_pt.FindNearestColor(self._check_colors) > 1
	
	def activate(self):
		self._check_pt.leftClick()
		return True

class EZD_Dlg_RadioBtns(Control):
	def __init__(self, name, dialog, btns_pt, btns_labels, colors, help):
		sibi = dialog.sibi
		super(EZD_Dlg_RadioBtns,self).__init__(name, sibi, None)
		self._btns = []
		for i,pt in enumerate(btns_pt):
			if i >= len(btns_labels):
				break
			self._btns.append( (sibi.MXY(pt), btns_labels[i]) )
		self._btn_colors = Color2Tuple(colors)
		self.help = help
		
	def _getChecked(self):
		for i,btn in enumerate(self._btns):
			if btn[0].FindNearestColor(self._btn_colors) == 1:
				return i, btn[1]
		return None, "Unknown"
		
	def getTextInfo(self):
		i, label = self._getChecked()
		return (self.type, self.name, label)

	def onEscape(self):
		return self.dialog.ctrl[len(self.dialog.ctrl) - 1].onEnter()
		
	def onUp(self):
		i, label = self._getChecked()
		if i is not None:
			if i > 0:
				i -= 1
			else:
				i = len(self._btns) - 1
			self._btns[i][0].leftClick()
		self.speakAfter()
		return True

	def onDown(self):
		i, label = self._getChecked()
		if i is not None:
			if i < len(self._btns) - 1:
				i += 1
			else:
				i = 0
			self._btns[i][0].leftClick()
		self.speakAfter()
		return True

	def onHelp(self, prefix = "", suffix = ""):
		self.speak(prefix + self.help + " Use up and down arrows to change the setting." + suffix)
		return True

class EZD_Settings_CloseBtn(Control):
	def __init__(self, name, sibi, btn_pt):
		super(EZD_Settings_CloseBtn, self).__init__(name, sibi, ("slow_reaction",))
		self._btn_pt = sibi.MXY(btn_pt)
		
	def onEnter(self):
		self._btn_pt.leftClick()
		self.speakFocusAfter()
		return True
		
class EZD_Settings(Container):
	
	def __init__(self, sibi):
		super(EZD_Settings,self).__init__("Settings", sibi, None)
		#self.add( EZD_Dialog_Name("Settings dialog", self) )
		pages = EZD_Settings_Pages( self )
		self.add( pages )
		
		page = EZD_Settings_Page("General", self, Pt(251, 197))
		pages.add( page )
		page.add( EZD_Dlg_CheckBox("Stop All on host stops", self, Pt(223, 467), (0x393431, 0xe06512), "Stop all MIDI previews on host stop.") )

		page = EZD_Settings_Page("MIDI libraries", self, Pt(320, 197))
		pages.add( page )
		page.add( Label("Not implemented yet", sibi, Pt(0,0), Pt(0,0), None) )

		page = EZD_Settings_Page("MIDI Events", self, Pt(420, 197))
		pages.add( page )
		page.add( EZD_Dlg_DropList("MIDI In Channel", self, Pt(221, 272), Pt(350, 287), "Select which channel to receive MIDI on.") )
		page.add( EZD_Dlg_CheckBox("Enable MIDI out", self, Pt(221, 375), (0x393431, 0xe06512), "Enable MIDI routing to other software.") ) 
		page.add( EZD_Dlg_CheckBox("Allow MIDI Program change", self, Pt(221, 450), (0x393431, 0xe06512), "Enable change of Sound Library and Library Preset via MIDI.") ) 

		page = EZD_Settings_Page("E-Drums", self, Pt(501, 197))
		pages.add( page )
		page.add( EZD_Dlg_DropList("MIDI Mapping", self, Pt(219, 308), Pt(347, 320), "Set up MIDI mapping to match your e-drums. Project specific setting.") )
		page.add( EZD_Dlg_DropList("Hi-Hat Pedal Correction", self, Pt(219, 432), Pt(349, 447), "Set up response curve for your hi-hat pedal. Project specific setting.") )

		page = EZD_Settings_Page("Sound engine", self, Pt(570, 197))
		pages.add( page )
		page.add( EZD_Dlg_RadioBtns("Humanize", self, (Pt(224, 293), Pt(224, 329), Pt(224, 365)), ("Always on", "EZX Optimized (Default)", "Always off"), (0xa19a97, 0x232323),
				  "Set how the sound engine should pick samples to play for individual hits. Project specific setting.") )
		
		self.add( CloseBtn("Set All to default", sibi, Pt(267, 509), Pt(267, 509), None) )
		self.add( EZD_Settings_CloseBtn("OK", sibi, Pt(669, 506)) )
		self.add( EZD_Settings_CloseBtn("Cancel", sibi, Pt(684, 169)) )

class EZD_MIDI_Library(Container):
	
	def __init__(self, sibi):
		super(EZD_MIDI_Library,self).__init__("Question", sibi, None)
		self.add( EZD_Dialog_Name("Question", self) )
		self.add( Label("", sibi, Pt(247, 262), Pt(648, 390), ("dynamic_text",)) )
		self.add( CloseBtn("Open settings", sibi, Pt(585, 413), Pt(585, 413), None) )
		self.add( CloseBtn("Not now", sibi, Pt(481, 412), Pt(481, 412), None) )
		
class EZD_Restore_Factory(Container):
	
	def __init__(self, sibi):
		super(EZD_Restore_Factory,self).__init__("Restore", sibi, None)
		self.add( EZD_Dialog_Name("Restore Factory MIDI Database", self) )
		self.add( Label("", sibi, Pt(225, 250), Pt(656, 409), ("dynamic_text",)) )
		self.add( CloseBtn("Restore Factory Database", sibi, Pt(570, 428), Pt(570, 428), None) )
		self.add( CloseBtn("Cancel", sibi, Pt(450, 429), Pt(450, 429), None) )

class EZD_Edit(Control):
	def __init__(self, name, sibi, lt, rb):
		super(EZD_Edit,self).__init__(name, sibi, None)
		self.box = EZD_TextOutBox(sibi, lt, rb)
	
	def getTextInfo(self):
		return (self.type, self.name, self.box.getTextOutOrText())
	
	def onEnter(self):
		self.expectFocusReturn()
		self.box.getBox().leftClick()
		return True
		
	def onHelp(self, prefix = "", suffix = ""):
		self.speak(prefix + "Press enter to start editing and enter to finish editing" + suffix)
		return True

class EZD_AuthorizeWithError(Control):
	def __init__(self, name, sibi, btn_pt, err_lt, err_rb):
		super(EZD_AuthorizeWithError,self).__init__(name, sibi, None)
		self.err = EZD_TextOutBox(sibi, err_lt, err_rb)
		self.err_pt = sibi.MXY(err_lt)
		self.btn_pt = sibi.MXY(btn_pt)
		self.clicked = False
	
	def getTextInfo(self):
		if self.clicked and self.err_pt.FindNearestColor((0xdcdcdc, 0x000000, 0xd43434)) == 2:
			text = "Error: " + self.err.getTextOutOrText()
			text = text.replace(" ll", " fill")
			text = text.replace(" le", " file")
		else:
			text = self.name
		self.clicked = False
		return (self.type, "", text)
	
	def onEnter(self):
		self.btn_pt.leftClick()
		self.clicked = True
		self.speakFocusAfter(2000)
		return True
		
	def onHelp(self, prefix = "", suffix = ""):
		self.speak(prefix + "Press enter to authorize" + suffix)
		return True

class EZD_OnlineAuthorization(Container):
	
	def __init__(self, sibi):
		super(EZD_OnlineAuthorization,self).__init__("Authorize online", sibi, None)
		self.add( CloseBtn("Authorization form. Press Enter to authorize offline or tab to authorize online.", sibi, Pt(602, 201), Pt(602, 201), None) )
		self.add( EZD_Edit("Email", sibi, Pt(320, 316), Pt(635, 328)) )
		self.add( EZD_Edit("Password", sibi, Pt(320, 363), Pt(635, 375)) )
		self.add( EZD_CheckBox("Remember me", sibi, Pt(322, 411), (0xfefefe, 0x323230)) )
		self.add( EZD_AuthorizeWithError("Log In and continue", sibi, Pt(593, 458), Pt(436, 268), Pt(637, 288) ) )

class EZD_RegisteredSerial(EZD_Label):
	def __init__(self, sibi):
		super(EZD_RegisteredSerial,self).__init__("Press enter to select registered to your account serial or up arrow to paste serial from clipboard.", sibi, Pt(310, 315), Pt(548, 338))
		self.use_pt = sibi.MXY(Pt(279, 312))
		self.has_serial_pt = sibi.MXY(Pt(278, 302))
		self.paste_pt = sibi.MXY(Pt(633, 390))
	
	def getTextInfo(self):
		if not self.__hasSerial():
			return (self.type, "Please press up arrow to paste serial from clipboard.", "")
		return super(EZD_RegisteredSerial,self).getTextInfo()
	
	def __hasSerial(self):
		return self.has_serial_pt.PixelColor() == 0x409417
	
	def onEnter(self):
		if not self.__hasSerial():
			self.speak("You do not have registered into account serial")
		else:
			self.use_pt.leftClick()
			self.speak("Registered serial is selected")
		return True

	def onUp(self):
		self.paste_pt.leftClick()
		self.speak("The serial from clipboard is selected")
		return True
		
class EZD_OnlineAuthorization2(Container):
	
	def __init__(self, sibi):
		super(EZD_OnlineAuthorization2,self).__init__("Authorize online", sibi, None)
		self.add( CloseBtn("Registering serial. Press Enter to return into log in form or tab to continue.", sibi, Pt(301, 481), Pt(301, 481), None) )
		self.add( EZD_RegisteredSerial(sibi) )
		self.add( EZD_Edit("Computer description", sibi, Pt(359, 423), Pt(590, 437)) )
		self.add( EZD_AuthorizeWithError("Authorize", sibi, Pt(615, 481), Pt(371, 345), Pt(574, 364) ) )

class EZD_IDCopyBtn(Control):
	def __init__(self, sibi, click_pt):
		super(EZD_IDCopyBtn,self).__init__("Press enter to copy this computer ID into clipboard", sibi, None)
		self.click_pt = sibi.MXY(click_pt)
		
	def onEnter(self):
		self.click_pt.leftClick()
		self.speak("This computer ID is in clipboard. Visit www.toontrack.com/register-product and follow the instructions")
		return True

class EZD_FileSelect(Label):
	def __init__(self, name, sibi, lt, rb, select_pt):
		super(EZD_FileSelect,self).__init__(name, sibi, lt, rb, ("dynamic_text",))
		self.select_pt = sibi.MXY(select_pt)
		
	def onEnter(self):
		self.expectFocusReturn()
		self.select_pt.leftClick()
		return True
		
	def onHelp(self, prefix = "", suffix = ""):
		self.speak(prefix + "Press enter to select file" + suffix)
		return True
		
class EZD_OfflineAuthorization(Container):
	
	def __init__(self, sibi):
		super(EZD_OfflineAuthorization,self).__init__("Authorize offline", sibi, None)
		self.add( CloseBtn("Authorize offline. Press enter to return into Online form or tab to continue with offline authorization.", sibi, Pt(252, 428), Pt(252, 428), None) )
		self.add( EZD_IDCopyBtn(sibi, Pt(678, 296)) )
		self.add( EZD_FileSelect("Authorization file", sibi, Pt(440, 371), Pt(634, 385), Pt(669, 378)) )
		self.add( EZD_AuthorizeWithError("Authorize", sibi, Pt(659, 429), Pt(184, 455), Pt(714, 474)) )

class EZD_OK_Dialog(Container):
	def __init__(self, sibi):
		super(EZD_OK_Dialog,self).__init__("OK", sibi, None)
		self.add( EZD_Label("", sibi, Pt(260, 235), Pt(638, 376)) )
		self.add( CloseBtn("Ok", sibi, Pt(448, 402), Pt(448, 402)) )
		
class EZD_Unknown_Dialog(Container):
	
	def __init__(self, sibi):
		super(EZD_Unknown_Dialog,self).__init__("Unknown", sibi, None)
		self.add( Label("Unknown dialog. Please report to Sibiac developer", sibi, Pt(0,0), Pt(0,0), None) )
		
class EZD_Dialog(Pages):
	""" Detect presence of any dialog, then decide what it is """
	
	def __init__(self, sibi):
		super(EZD_Dialog,self).__init__("Dialog", sibi, self.__getPageIndex, None)
		self.hwnd = sibi.hwnd
		self.check_pt = sibi.ptScale(Pt(18, 7))
		self.s_pt = sibi.ptScale(Pt(204, 193)) # settings dialog detector
		self.a_pt = sibi.ptScale(Pt(245, 231)) # about dialog detector
		self.a2_pt = sibi.ptScale(Pt(616, 384))
		self.r_pt = sibi.ptScale(Pt(223, 247)) # restore factory dialog detector
		self.m_pt = sibi.ptScale(Pt(244, 262)) # midi question dialog detector
		self.m2_pt = sibi.ptScale(Pt(585, 413))
		self.l_pt = sibi.ptScale(Pt(649, 452)) # Login dialog detector
		self.o_pt = sibi.ptScale(Pt(696, 421)) # Offline dialog detector
		self.l2_pt = sibi.ptScale(Pt(651, 472)) # Second online dialog detector
		self.ok_pt = sibi.ptScale(Pt(437, 401)) # Success dialog detector
		self.check_colors = (0xf29758, 0x482d1a)
		
		self.add( EZD_About(sibi) )
		self.add( EZD_Settings(sibi) )
		self.add( EZD_Restore_Factory(sibi) )
		self.add( EZD_MIDI_Library(sibi) )
		self.add( EZD_OnlineAuthorization(sibi) )
		self.add( EZD_OfflineAuthorization(sibi) )
		self.add( EZD_OnlineAuthorization2(sibi) )
		self.add( EZD_OK_Dialog(sibi) )
		self.add( EZD_Update_Dialog(sibi) )
		self.add( EZD_Unknown_Dialog(sibi) )
	
	def __getPageIndex(self):
		if FindNearestColor(self.hwnd, self.check_pt.x, self.check_pt.y, self.check_colors) != 1:
			return None
		log.error("Dialog...")
		if PixelColor(self.hwnd, self.s_pt.x, self.s_pt.y) == 0x918a87:
			return 1
		if PixelColor(self.hwnd, self.a_pt.x, self.a_pt.y) == 0xd4ccc8:
			if PixelColor(self.hwnd, self.a2_pt.x, self.a2_pt.y) == 0xfefefe:
				return 0
		if PixelColor(self.hwnd, self.r_pt.x, self.r_pt.y) == 0xd4ccc8:
			return 2
		if PixelColor(self.hwnd, self.m_pt.x, self.m_pt.y) == 0xd4ccc8:
			if PixelColor(self.hwnd, self.m2_pt.x, self.m2_pt.y) == 0xc7bfbb:
				return 8
			return 3
		if PixelColor(self.hwnd, self.l_pt.x, self.l_pt.y) == 0x3f9317:
			return 4
		if PixelColor(self.hwnd, self.o_pt.x, self.o_pt.y) == 0x419617:
			return 5
		if PixelColor(self.hwnd, self.l2_pt.x, self.l2_pt.y) == 0x419617:
			return 6
		if PixelColor(self.hwnd, self.ok_pt.x, self.ok_pt.y) == 0xe3e3e3:
			return 7
		return 9
		
	def focusNext(self):
		if not super(EZD_Dialog,self).focusNext():
			return self.focusFirst()
		return True
		
	def focusPrevious(self):
		if not super(EZD_Dialog,self).focusPrevious():
			return self.focusLast()
		return True

class EZD_PresetCtl(Control):
	def __init__(self, name, sibi):
		super(EZD_PresetCtl,self).__init__(name, sibi, ("delayed_reaction",))
		#self.preset = EZD_TextOutBox(sibi, Pt(610, 26), Pt(794, 40))
		#self.package = EZD_TextOutBox(sibi, Pt(610, 12), Pt(794, 24))
		self.preset = EZD_TextOutBox(sibi, Pt(620, 27), Pt(778, 38))
		self.package = EZD_TextOutBox(sibi, Pt(620, 14), Pt(778, 24))
		self.up_pt = sibi.MXY(Pt(812, 18))
		self.down_pt = sibi.MXY(Pt(812, 33))
		self.load_pt = sibi.MXY(Pt(613, 35))
		self.speak_package = False
		self.loading = 0
	
	def _isLoading(self):
		if FindNearestColor(self.sibi.hwnd, self.load_pt.x, self.load_pt.y, (0xba4f10, 0xce7528)) == 1:
			return True
		# sometimes fails...
		if self.loading == 0:
			return False
		text = self.preset.getTextOutOrText(2)
		if text == "":
			return True
		self.loading = 0
		return False
	
	def getTextInfo(self):
		if self._isLoading():
			if self.loading > 0:
				self.speakAfter(1000)
				if self.loading > 1:
					return ("", "", "\t") # do not annoy the user
				self.loading = 2
				return ("", "", "Loading preset")
		if not self.speak_package:
			text = self.preset.getTextOutOrText(2)
		else:
			text = "Package " + self.package.getTextOutOrText(2)
		if text == "" or text == "Package ":
			text = text + "not detected"
		return (self.type, self.name, text)
		
	def onUp(self):
		self.speak_package = False
		self.loading = 1
		self.up_pt.leftClick()
		self.speakAfter()
		return True

	def onDown(self):
		self.speak_package = False
		self.loading = 1
		self.down_pt.leftClick()
		self.speakAfter()
		return True
	
	def onLeft(self):
		self.speak_package = not self.speak_package
		self.speakAfter(0)
		return True

	def onRight(self):
		self.speak_package = not self.speak_package
		self.speakAfter(0)
		return True

	def focusSet(self):
		self.speak_package = False
		self.loading = 0
		super(EZD_PresetCtl,self).focusSet()

	def focusLost(self):
		self.loading = 0 # end delayed reports
		super(EZD_PresetCtl,self).focusLost()
		
	def onEnter(self):
		self.loading = 1
		box = self.preset.getBox()
		self.expectFocusReturn()
		box.leftClick()
		return True

	def onHelp(self, prefix = "", suffix = ""):
		self.speak(prefix + "Preset selection control. Use left and right arrows to audition current preset name and package." +
				   " Use up and down arrows to select previous and next preset. Press Enter to open context menu." +
				   " You can return to this control pressing Control+P." + suffix)
		return True

class EZD_SavePreset_Dialog(Container):
	def __init__(self, sibi):
		super(EZD_SavePreset_Dialog,self).__init__("Save preset", sibi, None)
		self.add( EZD_Edit("Preset name", sibi, Pt(361, 327), Pt(578, 339)) )
		self.add( CloseBtn("Ok", sibi, Pt(558, 374), Pt(558, 374)) )
		self.add( CloseBtn("Cancel", sibi, Pt(489, 376), Pt(489, 376)) )
		self.pt1 = sibi.ptScale(Pt(312,263))
		self.pt2 = sibi.ptScale(Pt(589,263))

	def isFocusable(self):
		x0, x1 = FindHRange(self.sibi.hwnd, self.pt1.x, self.pt2.x, self.pt1.y, (0x544e4d, 0x544e4b), 0x544e4c)
		return x0 == self.pt1.x and x1 == self.pt2.x

	def focusNext(self):
		if not super(EZD_SavePreset_Dialog,self).focusNext():
			return self.focusFirst()
		return True
		
	def focusPrevious(self):
		if not super(EZD_SavePreset_Dialog,self).focusPrevious():
			return self.focusLast()
		return True

class EZD_Mix_Volume(Control):
	__value_correct = { }

	def __init__(self, sibi, getShift):
		super(EZD_Mix_Volume,self).__init__("Volume", sibi, None)
		self.ctl_pt = sibi.MXY(Pt(69, 208), getShift)
		self.value = EZD_TextOutBox(sibi, Pt(58, 300), Pt(84, 308), getShift)
	
	def getValue(self):
		text = self.value.getTextOutOrText(0)
		if text != "Off":
			text = text.replace("S", "5")
			text = text.replace("O", "0")
			text = text.replace(" ", "")
			text = text.replace(",", ".")
			try:
				f = float(text)
				if f > 40 and f < 50:
					f = -f + 30 # 43 -> -13
					text = "%.1f" % f
			except:
				pass
			#if text in self.__value_correct:
			#	text = self.__value_correct[text]
			text += " dB"
		return text
	
	def onUp(self):
		self.ctl_pt.mouseScroll(120)
		return True

	def onDown(self):
		self.ctl_pt.mouseScroll(-120)
		return True
		
	def onEnter(self):
		self.expectFocusReturn()
		self.ctl_pt.rightClick()
		return True

	def onCtrlEnter(self):
		self.ctl_pt.leftClick()
		return True

class EZD_Mix_Pan(Control):
	def __init__(self, sibi, lt, rb, getShift):
		super(EZD_Mix_Pan,self).__init__("Pan", sibi, None)
		self.fader = sibi.Box(lt, rb, getShift)
		self.ctl_pt = sibi.MXY(lt, getShift)
		self.fader_bg = Color2Tuple(0x000000)
		self.fader_fg = Color2Tuple(0xc8c3bf)
		
	def getValue(self):
		fader = self.fader.getBox()
		#log.error("%x" % PixelColor(fader.hwnd, fader.left, fader.top))
		pos = FindInXRight(fader.hwnd, fader.left, fader.top, self.fader_bg, self.fader_fg)
		if pos < 0:
			return "Not detected"
		value = pos * 200 // (fader.right - fader.left) - 100
		if value == 0:
			text = "Center"
		elif value == -100:
			text = "Full left"
		elif value < 0:
			text = "%d percent left" % (-value)
		elif value == 100:
			text = "Full right"
		else:
			text = "%d percent right" % (value)
		return text
	
	def onUp(self):
		self.ctl_pt.mouseScroll(120)
		return True

	def onDown(self):
		self.ctl_pt.mouseScroll(-120)
		return True

	def onEnter(self):
		self.expectFocusReturn()
		self.ctl_pt.rightClick()
		return True

	def onCtrlEnter(self):
		self.ctl_pt.leftClick()
		return True
		
class EZD_Mix_Output(Control):

	def __init__(self, sibi, getShift):
		super(EZD_Mix_Output,self).__init__("Output", sibi, None)
		self.output = EZD_TextOutBox(sibi, Pt(64, 322), Pt(78, 330), getShift)
		
	def getValue(self):
		text = self.output.getTextOutOrText(1)
		if text == "":
			text = "8" # not sure why...
		else:
			text = text.replace("S", "5")
			text = text.replace("l", "1")
		return "output " + text
	
	def onEnter(self):
		self.expectFocusReturn()
		self.output.getBox().rightClick()
		return True
		
class EZD_Mix_Switch(Control):
	def __init__(self, sibi, ctl_pt, colors, states, getShift):
		super(EZD_Mix_Switch,self).__init__("Switch", sibi, None)
		self.ctl_pt = sibi.MXY(ctl_pt, getShift)
		self.colors = colors
		self.states = states
		
	def getValue(self):
		i = self.ctl_pt.FindNearestColor(self.colors)
		if i < 0:
			return "Not detected"
		return self.states[i]
	
	def onUp(self):
		self.ctl_pt.leftClick()
		return True

	def onDown(self):
		self.ctl_pt.leftClick()
		return True
	
	def onEnter(self):
		self.expectFocusReturn()
		self.ctl_pt.rightClick()
		return True
		
class EZD_Mix_Ch(Control):
	VOL = 0
	MONO_PAN = 1
	RIGHT_PAN = 2
	SOLO = 3
	MUTE = 4
	OUTPUT = 5

	__ch_name_correct = {
		"iGck In" : "Kick In",
		"IGck In" : "Kick In",
		"chk Out" : "Kick Out",
		"SDBotbom" : "SD Bottom",
		"Hi- Ha" : "Hi-Hat",
		"OH" : "Overhead",
		"chk" : "Kick",
		"Oonnp" : "Comp",
	}

	def __init__(self, sibi):
		super(EZD_Mix_Ch,self).__init__("Channel", sibi, None)
		self.ch_name = EZD_TextOutBox(sibi, Pt(42, 347), Pt(99, 360), self._getChShift)
		self.hwnd = sibi.hwnd
		self._pb = sibi.Box(Pt(21, 380), Pt(880, 380)) # panel bottom, width detection
		self._pb_x_shift = sibi.xScale(40) # dx for "Pop" preset
		self._pb_bg = Color2Tuple(0x2c2623)
		self._pb_fg = Color2Tuple(0x020303)
		self._ch_w = sibi.xScale(62) # channel width
		self._pb_left_border = sibi.xScale(7) # dx to channel visible area from left
		self._pb_right_border = sibi.xScale(8) # dx to channel visible area from right
		
		self._scr = sibi.Box(Pt(0, 372), Pt(0, 372)) # scroll bar, left and right are set dynamically
		self._scr_left_shift = sibi.xScale(8 - 7)
		self._scr_right_shift = sibi.xScale(9 - 8) # used when width is already corrected by _pb_right_border
		self._scr_bg = Color2Tuple(0x181615)
		self._scr_fg = Color2Tuple(0x655f5b)

		self.__reset()
		
		self._stereo_pt = sibi.MXY(Pt(59, 87), self._getChShift)
		self._stereo_color = set((0xd8cec5, 0xefe9e4))
		
		self.volume = EZD_Mix_Volume(sibi, self._getChShift)
		self.mono_pan = EZD_Mix_Pan(sibi, Pt(50, 95), Pt(86, 95), self._getChShift)
		self.left_pan = EZD_Mix_Pan(sibi, Pt(51, 88), Pt(86, 88), self._getChShift)
		self.right_pan = EZD_Mix_Pan(sibi, Pt(52, 102), Pt(86, 102), self._getChShift)
		self.solo = EZD_Mix_Switch(sibi, Pt(54, 123), (0x58524f, 0xf6c993), ("not soloed", "soloed"), self._getChShift)
		self.mute = EZD_Mix_Switch(sibi, Pt(77, 123), (0x58524f, 0xa4d5ee), ("not muted", "muted"), self._getChShift)
		self.output = EZD_Mix_Output(sibi, self._getChShift)
		
		self._select_pt = sibi.MXY(Pt(72, 78), self._getChShift)
		self._select_colors = (0xd9cec6, 0xf0e9e4)
		
	def _getChShift(self):
		return (self._pb_ch_shift + self._ch_idx * self._ch_w, 0)
	
	def __reset(self):
		self._pb_left = -1 # evaluated panel position
		self._pb_right = -1
		self._pb_width = 0
		self._pb_full_width = 0
		self._ch_n = 0 # number of channels
		self._ch_idx = 0
		
		self._scr_left = -1  # scroll position
		self._scr_right = -1
		self._scr_left_dx = 0
		self._scr_right_dx = 0
		
		self._pb_ch_shift = 0
		
		self._par = self.VOL
		self._par_changed = False
	
	def __detectScroll(self):
		self._pb_ch_shift = self._pb_left - self._pb_x_shift # base position, not scrolled right
		self._scr.left = self._pb_left + self._scr_left_shift
		self._scr.right = self._pb_right - self._scr_right_shift
		self._scr_left, self._scr_right = FindHRange(self.hwnd, self._scr.left, self._scr.right, self._scr.top, self._scr_bg, self._scr_fg)
		if self._scr_left < 0: # no scroll
			self._scr_left_dx = 0
			self._scr_right_dx = 0
			self._scr_factor = 1.
			self._pb_full_width = self._pb_width
		else:
			self._scr_left_dx = self._scr_left - self._scr.left
			self._scr_right_dx = self._scr.right - self._scr_right
			sw = self._scr.right - self._scr.left + 1
			spw = self._scr_right - self._scr_left + 1
			if spw > 0:
				self._pb_full_width = sw * self._pb_width // spw
			else:
				self._pb_full_width = self._pb_width
			if self._scr_left_dx: 
				# scrolled right, need to set channel shift accurately
				# set it from the panel right
				ch_n = (self._pb_full_width + 5) // self._ch_w # try to void rounding errors...
				self._pb_ch_shift = self._pb_right - ch_n * self._ch_w + 1 - self._pb_x_shift

	def _isStereo(self):
		""" No detection """
		return self._stereo_pt.PixelColor() not in self._stereo_color
	
	def _selectedText(self):
		if self._select_pt.FindNearestColor(self._select_colors) == 1:
			return "Selected "
		return ""
	
	def _scrollToIndex(self):
		""" No detection """
		if self._ch_idx >= self._ch_n:
			return False
		x = self._ch_idx * self._ch_w - self._scr_left_dx
		if x >= 0 and (x + self._ch_w) <= (self._pb_right - self._pb_left + 1):
			return True
		# use only 2 scroll positions, full left or full right
		if self._ch_idx < self._ch_n // 2:
			box = Box(self.hwnd, self._scr_right - 2, self._scr.top, self._scr_right - 2 - self._scr_left, self._scr.top) # drag left
		else:
			box = Box(self.hwnd, self._scr_left + 2, self._scr.top, self._scr_left + 2 + self._scr_right, self._scr.top) # drag right
		box.leftDrag(0.1)
		time.sleep(0.1)
		self.__detectScroll()
		return True
	
	def _detectPanel(self):
		self._pb_left, self._pb_right = FindHRange(self.hwnd, self._pb.left, self._pb.right, self._pb.bottom, self._pb_bg, self._pb_fg)
		if self._pb_left < 0:
			self.__reset()
		else:
			self._pb_left += self._pb_left_border
			self._pb_right -= self._pb_right_border
			self._pb_width = self._pb_right - self._pb_left + 1
			self.__detectScroll()
			self._ch_n = (self._pb_full_width + 5) // self._ch_w # try to void rounding errors...
			if self._ch_idx >= self._ch_n:
				self._ch_idx = 0
			#log.error("Panel: left %d, right %d, scroll left %d + %d, scroll right %d - %d, visible panel width %d, total panel width %d, shift %d" %
			#  (self._pb_left, self._pb_right, self._scr.left, self._scr_left_dx, self._scr.right, self._scr_right_dx, self._pb_width, self._pb_full_width, self._pb_ch_shift) )
			self._scrollToIndex()
			#log.error("Panel: left %d, right %d, scroll left %d + %d, scroll right %d - %d, visible panel width %d, total panel width %d, shift %d" %
			#  (self._pb_left, self._pb_right, self._scr.left, self._scr_left_dx, self._scr.right, self._scr_right_dx, self._pb_width, self._pb_full_width, self._pb_ch_shift) )
		return self._pb_left >= 0
	
	def _getCh(self):
		if not self._detectPanel():
			return None, ""
		ch_name = self.ch_name.getTextOutOrText()
		if ch_name in self.__ch_name_correct:
			ch_name = self.__ch_name_correct[ch_name]
		return (self._ch_idx, ch_name)
	
	def _isFirstCh(self):
		""" No detection """
		return self._ch_idx < 1

	def _isLastCh(self):
		""" No detection """
		return self._ch_idx >= self._ch_n - 1
		
	def _changeCh(self, d):
		""" No detection """
		if d < 0:
			if self._isFirstCh():
				return False
		else:
			if self._isLastCh():
				return False
		self._ch_idx += d
		self._scrollToIndex()
		return True
	
	def getTextInfo(self):
		idx, text = self._getCh()
		if idx is None:
			text = "Could not detect mixer panel"
		else:
			text = self._selectedText() + text
			if self._par_changed:
				text = " "
				self._par_changed = False
			else:
				text += " "
			if self._par == self.VOL:
				text += self.volume.getValue()
			elif self._par == self.MONO_PAN:
				if self._isStereo():
					if text != " ":
						text += "left "
					text += self.left_pan.getValue()
				else:
					text += self.mono_pan.getValue()
			elif self._par == self.RIGHT_PAN:
				if self._isStereo():
					if text != " ":
						text += "right "
					text += self.right_pan.getValue()
				else:
					text += self.mono_pan.getValue()
			elif self._par == self.MUTE:
				text += self.mute.getValue()
			elif self._par == self.SOLO:
				text += self.solo.getValue()
			elif self._par == self.OUTPUT:
				text += self.output.getValue()
		return (self.type, self.name, text)
		
	def onLeft(self):
		idx, ch_name = self._getCh()
		if idx is None:
			self.speak("Could not detect mixer panel")
		elif self._isFirstCh():
			self.speak("%s, first" % ch_name)
		else:
			self._changeCh(-1)
			self.speakAfter()
		return True

	def onRight(self):
		idx, ch_name = self._getCh()
		if idx is None:
			self.speak("Could not detect mixer panel")
		elif self._isLastCh():
			self.speak("%s, last" % ch_name)
		else:
			self._changeCh(1)
			self.speakAfter()
		return True
	
	def onCtrlLeft(self):
		if self._par == self.VOL:
			self._par = self.OUTPUT
		elif self._par == self.SOLO and not self._isStereo():
			self._par = self.MONO_PAN
		else:
			self._par -= 1
		self.speakAfter(0)
		return True

	def onCtrlRight(self):
		if self._par == self.OUTPUT:
			self._par = self.VOL
		elif self._par == self.MONO_PAN and not self._isStereo():
			self._par = self.SOLO
		else:
			self._par += 1
		self.speakAfter(0)
		return True
		
	
	def onUp(self):
		idx, ch_name = self._getCh()
		if idx is None:
			return True
		if self._par == self.VOL:
			self.volume.onUp()
		elif self._par == self.MONO_PAN:
			if self._isStereo():
				self.left_pan.onUp()
			else:
				self.mono_pan.onUp()
		elif self._par == self.RIGHT_PAN:
			if self._isStereo():
				self.right_pan.onUp()
			else:
				self.mono_pan.onUp()
		elif self._par == self.MUTE:
			self.mute.onUp()
		elif self._par == self.SOLO:
			self.solo.onUp()
		self._par_changed = True
		self.speakAfter()
		return True

	def onDown(self):
		idx, ch_name = self._getCh()
		if idx is None:
			return True
		if self._par == self.VOL:
			self.volume.onDown()
		elif self._par == self.MONO_PAN:
			if self._isStereo():
				self.left_pan.onDown()
			else:
				self.mono_pan.onDown()
		elif self._par == self.RIGHT_PAN:
			if self._isStereo():
				self.right_pan.onDown()
			else:
				self.mono_pan.onDown()
		elif self._par == self.MUTE:
			self.mute.onDown()
		elif self._par == self.SOLO:
			self.solo.onDown()
		self._par_changed = True
		self.speakAfter()
		return True

	def onEnter(self):
		idx, ch_name = self._getCh()
		if idx is None:
			return True
		if self._par == self.MUTE:
			self.mute.onUp()
		elif self._par == self.SOLO:
			self.solo.onUp()
		elif self._par == self.OUTPUT:
			self.output.onEnter()
			return True
		self._par_changed = True
		self.speakAfter()
		return True

	def onShiftEnter(self):
		idx, ch_name = self._getCh()
		if idx is None:
			return True
		# detect which shift was pressed and temporarily de-press it
		if winUser.getKeyState(winUser.VK_LSHIFT) & 32768:
			vk_key = winUser.VK_LSHIFT
		elif winUser.getKeyState(winUser.VK_RSHIFT) & 32768:
			vk_key = winUser.VK_RSHIFT
		else:
			log.error("Unknown shift pressed")
			return True
		code = winUser.user32.MapVirtualKeyW(vk_key, 0) # 0 - MAPVK_VK_TO_VSC
		winUser.keybd_event(vk_key, code, winUser.KEYEVENTF_KEYUP, 0)
		if self._par == self.VOL:
			self.volume.onEnter()
		elif self._par == self.MONO_PAN:
			if self._isStereo():
				self.left_pan.onEnter()
			else:
				self.mono_pan.onEnter()
		elif self._par == self.RIGHT_PAN:
			if self._isStereo():
				self.right_pan.onEnter()
			else:
				self.mono_pan.onEnter()
		elif self._par == self.MUTE:
			self.mute.onEnter()
		elif self._par == self.SOLO:
			self.solo.onEnter()
		elif self._par == self.OUTPUT:
			self.output.onEnter()
		winUser.keybd_event(vk_key, code, 0, 0) # return to original keyboard state
		return True

	def onCtrlUp(self):
		idx, ch_name = self._getCh()
		if idx is None:
			return True
		self._select_pt.leftClick()
		self.speakInFocusAfter()
		return True

	def onCtrlDown(self):
		idx, ch_name = self._getCh()
		if idx is None:
			return True
		if winUser.getKeyState(winUser.VK_LCONTROL) & 32768:
			vk_key = winUser.VK_LCONTROL
		elif winUser.getKeyState(winUser.VK_RCONTROL) & 32768:
			vk_key = winUser.VK_RCONTROL
		else:
			log.error("Unknown shift pressed")
			return True
		code = winUser.user32.MapVirtualKeyW(vk_key, 0) # 0 - MAPVK_VK_TO_VSC
		winUser.keybd_event(vk_key, code, winUser.KEYEVENTF_KEYUP, 0)
		self._select_pt.leftClick()
		winUser.keybd_event(vk_key, code, 0, 0)
		self.speakInFocusAfter()
		return True
		
	def onCtrlEnter(self):
		idx, ch_name = self._getCh()
		if idx is None:
			return True
		if self._par == self.VOL:
			self.volume.onCtrlEnter()
		elif self._par == self.MONO_PAN:
			if self._isStereo():
				self.left_pan.onCtrlEnter()
			else:
				self.mono_pan.onCtrlEnter()
		elif self._par == self.RIGHT_PAN:
			if self._isStereo():
				self.right_pan.onCtrlEnter()
			else:
				self.mono_pan.onCtrlEnter()
		self._par_changed = True
		self.speakAfter()
		return True
	
	def onNum1(self):
		self._ch_idx = 0
		self._par = self.VOL
		self.speakFocusAfter()
		return True
	
	def onHelp(self, prefix = "", suffix = ""):
		self.speak(prefix + "Use left and right arrow to select mixer channel. Pressing 1 return to the first channel. Use control+left and right arrows to select channel parameter. Use up and down arrow to change" +
                            " the parameter value. For volume and pan, use shift+up and down arrow for fine changes and control+enter to reset the value. Use enter to toggle mute or solo, audition current" +
							" volume and pan value or select output channel. Press shift+enter to open context menu. Control+up arrow toggle channel selection without changing selection for other channels," +
							" Control+down arrow toggle channel selection and de-select all other channels. Pressing Control+down arrow twice select all channels. Parameters changes for any selected channel"
							" also change the same parameter for other selected channels" + suffix)
		return True

class EZD_FX(object):
	def __init__(self, left, right):
		self.left = left
		self.right = right
		self.name = None
		#log.error("%d %d = %d" % (left, right, right - left))
		self.n_ctrl = (right - left)//70 # real width should not be smaller but we do not want overestimate
		if self.n_ctrl == 0:
			self.n_ctrl = 1
		self.par_width =(right - left) // self.n_ctrl
		self.par_name = []
		for i in range(self.n_ctrl):
			self.par_name.append(None)

class EZD_Effects(Control):
	__par_name_correct = {
		"" : "Unknown",
		"brave" : "Drive",
		"Ambienoe" : "Ambience",
		"OH" : "Overhead",
	}

	def __init__(self, sibi):
		super(EZD_Effects,self).__init__("Effect", sibi, None)
		self.type = "Effect"
		self.hwnd = sibi.hwnd
		self._fx_name = EZD_TextOutBox(sibi, Pt(0, 398), Pt(0, 408))
		self._par_name = EZD_TextOutBox(sibi, Pt(0, 456), Pt(0, 466))
		self._par_value = EZD_TextOutBox(sibi, Pt(0, 370), Pt(0, 381))
		self._fx_top = sibi.Box(Pt(20, 392), Pt(880, 392))
		self._fx_top_bg = 0x202020
		self._fx_top_fg = 0x101010
		self._fx_border = sibi.xScale(16)
		self._fx_ctrl_y = sibi.yScale(437)
		self._fx = []
		self._fx_idx = None
		self._par_idx = None
		
	def __detectEffects(self):
		self._fx = []
		x = self._fx_top.left
		while x < self._fx_top.right:
			x1, x2 = FindHRange(self.hwnd, x, self._fx_top.right, self._fx_top.top, self._fx_top_bg, self._fx_top_fg)
			if x1 < 0 or (x2 - x1) < self._fx_border*2:
				break
			self._fx.append(EZD_FX(x1 + self._fx_border, x2 - self._fx_border))
			#log.error("%d %d" % (x1, x2))
			x = x2 + 10
		if len(self._fx) == 0:
			self._fx_idx = None
			self._par_idx = None
			return
		if self._fx_idx is None or self._fx_idx >= len(self._fx):
			self._fx_idx = 0
			self._par_idx = 0
		if self._par_idx is None or self._par_idx >= self._fx[self._fx_idx].n_ctrl:
			self._par_idx = 0

	def _getFXName(self):
		if self._fx_idx is None:
			return ""
		fx = self._fx[self._fx_idx]
		if fx.name is None:
			self._fx_name.left = fx.left
			self._fx_name.right = fx.right
			fx.name = self._fx_name.getTextOutOrText()
			if fx.name == "":
				fx.name = "Unknown"
			if fx.name == "T":
				fx.name = "Drums"
			else:
				fx.name = fx.name.replace("Revel-b", "Reverb")
				fx.name = fx.name.replace("Revert.", "Reverb")
		return fx.name
	
	def _getParName(self):
		if self._fx_idx is None:
			return ""
		fx = self._fx[self._fx_idx]
		par_name = fx.par_name[self._par_idx]
		if par_name is None:
			self._par_name.left = fx.left + fx.par_width * self._par_idx
			self._par_name.right = fx.left + fx.par_width * (self._par_idx + 1)
			par_name = self._par_name.getTextOutOrText()
			if par_name in self.__par_name_correct:
				par_name = self.__par_name_correct[par_name]
			else:
				par_name = par_name.replace("Pitd'l", "Pitch")
				par_name = par_name.replace("Pitdl", "Pitch")
			fx.par_name[self._par_idx] = par_name
		return par_name
	
	def _getParValue(self):
		ctrl = self._getParCtrl()
		if ctrl is None:
			return ""
		ctrl.leftClick()
		for attempt in range(6):
			time.sleep(0.05)
			if self._par_value.findInX(ctrl.x, (0x655f5b, 0x181615), 0xffffbb):
				text = self._par_value.getTextOutOrText()
				if text != "":
					if text != "96" and text.endswith("96"):
						text = text[:-2] + " percent"
					text.replace("o", "0")
					text.replace("O", "0")
					return text
		return "Not found"
	
	def _getParCtrl(self):
		if self._fx_idx is None:
			return None
		fx = self._fx[self._fx_idx]
		return MXY(self.hwnd, fx.left + fx.par_width * self._par_idx + fx.par_width//2, self._fx_ctrl_y) 
	
	def focusSet(self):
		super(EZD_Effects,self).focusSet()
		self.__detectEffects()
		
	def getTextInfo(self):
		fx_name = self._getFXName()
		if fx_name == "":
			return (self.type, "Current preset has no effects", "")
		par_name = self._getParName()
		return (self.type, fx_name + " " + par_name, self._getParValue())

	def onLeft(self):
		if self._fx_idx is not None:
			if self._par_idx > 0:
				self._par_idx -= 1
			else:
				if self._fx_idx > 0:
					self._fx_idx -= 1
				else:
					self._fx_idx = len(self._fx) - 1
				self._par_idx = self._fx[self._fx_idx].n_ctrl - 1
		self.speakInFocusAfter()
		return True
		
	def onRight(self):
		if self._fx_idx is not None:
			if self._par_idx < self._fx[self._fx_idx].n_ctrl - 1:
				self._par_idx += 1
			else:
				if self._fx_idx < len(self._fx) - 1:
					self._fx_idx += 1
				else:
					self._fx_idx = 0
				self._par_idx = 0
		self.speakInFocusAfter()
		return True

	def onUp(self):
		ctrl = self._getParCtrl()
		if ctrl is not None:
			ctrl.mouseScroll(120)
		self.speakAfter()
		return True

	def onDown(self):
		ctrl = self._getParCtrl()
		if ctrl is not None:
			ctrl.mouseScroll(-120)
		self.speakAfter()
		return True
	
	def onEnter(self):
		self.speakAfter(0)
		return True

	def onCtrlEnter(self):
		ctrl = self._getParCtrl()
		if ctrl is not None:
			ctrl.leftClick()
		self.speakAfter()
		return True
		
	def onShiftEnter(self):
		ctrl = self._getParCtrl()
		if ctrl is not None:
			# detect which shift was pressed and temporarily de-press it
			if winUser.getKeyState(winUser.VK_LSHIFT) & 32768:
				vk_key = winUser.VK_LSHIFT
			elif winUser.getKeyState(winUser.VK_RSHIFT) & 32768:
				vk_key = winUser.VK_RSHIFT
			else:
				log.error("Unknown shift pressed")
				return True
			code = winUser.user32.MapVirtualKeyW(vk_key, 0) # 0 - MAPVK_VK_TO_VSC
			winUser.keybd_event(vk_key, code, winUser.KEYEVENTF_KEYUP, 0)
			self.expectFocusReturn()
			ctrl.rightClick()
			winUser.keybd_event(vk_key, code, 0, 0)
		return True

	def onHelp(self, prefix = "", suffix = ""):
		self.speak(prefix + "Use left and right arrow to select effects parameter. Use up and down arrow to change" +
                            " the parameter value, with shift for fine changes. Press control+enter to reset the value to preset default." +
							" Press shift+enter to open context menu." + suffix)
		return True
		
class EZDrummer(SIBINVDA):
	sname = "EZ Drummer"
	displayText = ""
	
	def defineSIBI(self):
		self.sibi = SIBI(self.windowHandle, 900, 660, ("block_arrows","textout"))
		sibi = self.sibi
		
		#sibi.add( Label("Dummy", sibi, Pt(0,0), Pt(0,0), None) )
		sibi.add( EZD_PresetCtl("Preset", sibi) )
		pages = FixedTabControl("", sibi, ("slow_reaction",))
		pages.type = "Page"
		sibi.add( pages )
		
		drums = FixedTab("Drums", sibi, Pt(215, 17), Pt(287, 40), Pt(249, 14), 0x48413f, 0x000000, None)
		pages.add( drums )
		drums.add( EZD_Kit(sibi) )
		
		mixer = FixedTab("Mixer", sibi, Pt(503, 17), Pt(567, 40), Pt(536, 14), 0x48413f, 0x6e6662, None)
		pages.add( mixer )
		mixer.add( EZD_Mix_Ch(sibi) )
		mixer.add( EZD_Effects(sibi) )
		mixer.add( Clickable("Mixer menu", sibi, Pt(865, 59), Pt(865, 59), ("silent_action", "focus_will_move")) )

		
		sibi.add( Clickable("Settings and information menu", sibi, Pt(861, 25), Pt(861, 25), ("silent_action", "focus_will_move")) )
		
		sibi.addDialog( EZD_Dialog(sibi) )
		sibi.addDialog( EZD_SavePreset_Dialog(sibi) )

	def script_onCtrlP(self, gesture):
		if self._skipGesture(gesture):
			return
		if self.sibi.last_dlg:
			return # we are in dialog
		self.sibi.focusFirst() # preset control
		self.speakFocusAfter(0)
		
	def script_onCtrlShiftR(self, gesture):
		self._onScript("onCtrlShiftR", False, gesture)

	def script_onCtrlShiftL(self, gesture):
		self._onScript("onCtrlShiftL", False, gesture)

	def script_onCtrlShiftLeft(self, gesture):
		self._onScript("onCtrlShiftLeft", False, gesture)

	def script_onCtrlShiftRight(self, gesture):
		self._onScript("onCtrlShiftRight", False, gesture)

	def script_onCtrlUp(self, gesture):
		self._onScript("onCtrlUp", False, gesture)

	def script_onCtrlDown(self, gesture):
		self._onScript("onCtrlDown", False, gesture)

	def script_onCtrlLeft(self, gesture):
		self._onScript("onCtrlLeft", False, gesture)

	def script_onCtrlRight(self, gesture):
		self._onScript("onCtrlRight", False, gesture)

	def script_onCtrlEnter(self, gesture):
		self._onScript("onCtrlEnter", False, gesture)

	def script_onShiftEnter(self, gesture):
		self._onScript("onShiftEnter", False, gesture)

	def script_onAltEnter(self, gesture):
		self._onScript("onAltEnter", False, gesture)
		
	def script_onNum1(self, gesture):
		self._onScript("onNum1", True, gesture)
		
	def script_onNum2(self, gesture):
		self._onScript("onNum2", True, gesture)
		
	def script_onNum3(self, gesture):
		self._onScript("onNum3", True, gesture)
		
	def script_onNum4(self, gesture):
		self._onScript("onNum4", True, gesture)
		
	def script_onNum5(self, gesture):
		self._onScript("onNum5", True, gesture)
		
	def script_onNum6(self, gesture):
		self._onScript("onNum6", True, gesture)

	def script_onNum7(self, gesture):
		self._onScript("onNum7", True, gesture)

	def script_onNum8(self, gesture):
		self._onScript("onNum8", True, gesture)
		
	__gestures = {
		"kb:control+p": "onCtrlP",
		"kb:shift+upArrow": "up",
		"kb:shift+downArrow": "down",
		"kb:shift+leftArrow": "left",
		"kb:shift+rightArrow": "right",
		"kb:control+leftArrow": "onCtrlLeft",
		"kb:control+rightArrow": "onCtrlRight",
		"kb:control+shift+l": "onCtrlShiftL",
		"kb:control+shift+r": "onCtrlShiftR",
		"kb:control+shift+leftArrow": "onCtrlShiftLeft",
		"kb:control+shift+rightArrow": "onCtrlShiftRight",
		"kb:control+upArrow": "onCtrlUp",
		"kb:control+downArrow": "onCtrlDown",
		"kb:control+shift+upArrow": "onCtrlUp",
		"kb:control+shift+downArrow": "onCtrlDown",
		"kb:control+enter": "onCtrlEnter",
		"kb:shift+enter": "onShiftEnter",
		"kb:alt+enter": "onAltEnter",
		"kb:1": "onNum1",
		"kb:2": "onNum2",
		"kb:3": "onNum3",
		"kb:4": "onNum4",
		"kb:5": "onNum5",
		"kb:6": "onNum6",
		"kb:7": "onNum7",
		"kb:8": "onNum8",		
	}