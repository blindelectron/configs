# Single Image Blob Interface Accessible Control
#
# SIBIAC API
#
# AZ (www.azslow.com), 2018 - 2020
from __future__ import absolute_import
from __future__ import division
from __future__ import print_function

import re
import unicodedata
from ctypes import *
from ctypes.wintypes import *
import itertools
import os
import api
from logHandler import log
import winUser
import eventHandler
import watchdog

# NVDA compatibility
from buildVersion import version_year, version_major
if version_year < 2019 or (version_year == 2019 and version_major < 3):
	import textInfos
	textInfosRect = textInfos.Rect
else:
	import locationHelper
	textInfosRect = locationHelper.RectLTRB

# Python 2/3 compatibility

try:
    from itertools import izip as zip
except ImportError: # will be 3.x series
    pass

try:
	xrange
except NameError:
	xrange = range
	
import sys
if sys.version_info < (3,):
	def b(x):
		return x
	def u2ascii(x):
		return buf.value.decode('utf-8', 'ignore').encode('ascii', 'ignore').replace('\n', '')
else:
	import codecs
	def b(x):
		return codecs.latin_1_encode(x)[0]
	def u2ascii(x):
		u = unicodedata.normalize('NFKC', x)
		a = re.sub(r'[^\x00-\x7f]',r' ',u)
		return a.replace('\n', '') # extract ASCII symbols only, without new space

	
# Sometimes it takes too long for NVDA
watchdog.asleep()

user32 = windll.user32
kernel32 = windll.kernel32 

WS_CHILD = 0x40000000
		
def isWinNotChild(hwnd):
	dwStyle = user32.GetWindowLongW(hwnd, winUser.GWL_STYLE)
	return (dwStyle & WS_CHILD) == 0  

def findTopWindowFor(hwnd):
	"""
	return Top window for specified window
	"""
	while hwnd:
		if isWinNotChild(hwnd):
			break
		hwnd = user32.GetParent(hwnd)
	return hwnd

SPI_GETWORKAREA = 0x0030
SW_MAXIMIZE = 3

def maximizeWindow(hwnd):
	"""
	Maximize top window for hwnd
	"""
	#r = RECT()
	#user32.SystemParametersInfoA(SPI_GETWORKAREA, 0, byref(r), 0)
	top = findTopWindowFor(hwnd)
	#user32.MoveWindow(top, r.left, r.top, r.right - r.left, r.bottom - r.top, BOOL(True))
	user32.ShowWindow(top, SW_MAXIMIZE)

def fitWindowIntoDesktop(hwnd):
	"""
	Try to move top window so it is inside Desktop. Maximze if not possible
	"""
	top = findTopWindowFor(hwnd)
	r = RECT()
	dr = RECT()
	user32.GetWindowRect(top, byref(r))
	user32.SystemParametersInfoA(SPI_GETWORKAREA, 0, byref(dr), 0)
	w = r.right - r.left
	h = r.bottom - r.top
	if w < dr.right and h < dr.bottom:
		# should fit
		if r.left < 0:
			r.left = 0
		elif r.right > dr.right:
			r.left = dr.right - w
		if r.top < 0:
			r.top = 0
		elif r.bottom > dr.bottom:
			r.top = dr.bottom - h
		user32.MoveWindow(top, r.left, r.top, w, h, BOOL(True))
	else:
		# no other options...
		user32.ShowWindow(top, SW_MAXIMIZE)

def findTopmostOver(hwnd):
	"""
	Find if there are topmost windows which can hide our window
	"""
	top = findTopWindowFor(hwnd)
	r = RECT()
	user32.GetWindowRect(hwnd, byref(r))
	tr = RECT()
	desk = user32.GetDesktopWindow()
	pwnd = None
	twnd = user32.FindWindowExW(desk, pwnd, None, None)
	while twnd:
		if user32.GetWindowLongW(twnd, winUser.GWL_EXSTYLE) & winUser.WS_EX_TOPMOST and winUser.isWindowVisible(twnd):
			user32.GetWindowRect(twnd, byref(tr))
			cls = winUser.getClassName(twnd)
			#if isWinNotChild(twnd) and tr.left < r.right and tr.right > r.left and tr.top < r.bottom and tr.bottom > r.top:
			if not cls.startswith("Windows") and tr.left < r.right and tr.right > r.left and tr.top < r.bottom and tr.bottom > r.top:
				# check it is really visible over us
				x = (max(tr.left, r.left) + min(tr.right, r.right)) // 2
				y = (max(tr.top, r.top) + min(tr.bottom, r.bottom)) // 2
				p = POINT(x, y)
				ttwnd = findTopWindowFor(_windowFromPoint(p))
				if ttwnd == twnd and ttwnd != top:
					title = winUser.getWindowText(twnd)
					log.error("Wnd: %x (Top: %x), Pt(%d,%d), OnTop: %x, Vis:%x '%s' '%s' (Top:%x)" % (hwnd, top, x, y, twnd, _windowFromPoint(p), cls, title, findTopWindowFor(_windowFromPoint(p))))
					if title == "":
						return "Unknown"
					return title
		pwnd = twnd
		twnd = user32.FindWindowExW(desk, pwnd, None, None)
	return ""
		
def GetClientRect(hwnd):
	r = RECT()
	user32.GetClientRect(hwnd, byref(r))
	return Box(hwnd, r.left, r.top, r.right, r.bottom)
	
def ClientRectToScreen(hwnd, r):
	left_top = POINT(r.left, r.top)
	right_bottom = POINT(r.right, r.bottom)
	user32.ClientToScreen(hwnd, byref(left_top))
	user32.ClientToScreen(hwnd, byref(right_bottom))
	return RECT(left_top.x, left_top.y, right_bottom.x, right_bottom.y)	

def MoveFocusTo(hWnd):
	"""
	Move system focus to specified application window
	"""
	if eventHandler.isPendingEvents("gainFocus"):
		return # do not mess too much
	curThreadID = kernel32.GetCurrentThreadId()
	targetThreadID = user32.GetWindowThreadProcessId(hWnd, 0);
	if curThreadID != targetThreadID:
		user32.AttachThreadInput(curThreadID, targetThreadID,True)
	user32.SetFocus(hWnd)
	if curThreadID != targetThreadID:
		user32.AttachThreadInput(curThreadID, targetThreadID, False)


def getTopWindowsFor(pid, exclude_cls):
	"""
	Return sorted list of top windows for process.
	If exclude_cls is not None, that should be list of class names to exclude.
	Could be done with EnumWindows...
	"""
	wlist = []
	desk = user32.GetDesktopWindow()
	nwnd = user32.FindWindowExW(desk, None, None, None)
	while nwnd:
		if winUser.isWindowVisible(nwnd) and winUser.getWindowThreadProcessID(nwnd)[0] == pid:
			if not exclude_cls or not winUser.getClassName(nwnd) in exclude_cls:
				wlist.append( nwnd )
				# log.info("++ %d '%s' %s " % (nwnd, winUser.getWindowText(nwnd), winUser.getClassName(nwnd)))
		nwnd = user32.FindWindowExW(desk, nwnd, None, None)
	wlist.sort()
	return wlist

def focusNextTopWindowFor(pid, exclude_cls):
	"""
	if current focus is in pid, find next top window and focus it
	
	Implementation notes: I could use EnumWindows...
	   Since Z Order change after focus is changed, simply selecting the "next" loops inside dialogs.
	   So I order by hwnd (number)
	"""
	cwnd = winUser.getForegroundWindow()
	if cwnd:
		top = findTopWindowFor(cwnd)
		wlist = getTopWindowsFor(pid, exclude_cls)
		if top and top in wlist and len(wlist) > 1:
			idx = wlist.index( top ) + 1
			if idx >= len(wlist):
				idx = 0
			nwnd = wlist[idx]
			MoveFocusTo(wlist[idx])
			return True
	return False

sibiac_dir = os.path.dirname(__file__)
dll = os.path.join(sibiac_dir, "libsibiac.dll")
sibiac = cdll.LoadLibrary(dll)

def getScreenPos(hWnd, x, y):
	p = POINT(x, y)
	user32.ClientToScreen(hWnd, byref(p))
	return (p.x, p.y)

def OneIfZero( n ):
	"""
	Return n if n != 0, 1 otherwise
	"""
	if n == 0:
		return 1
	return n

def reversed_enumerate( L ):
	"""
	enumberate in reversed order
	"""
	return zip( reversed(xrange(len(L))), reversed(L) )

def enumerate_from( L , start):
	"""
	enumberate slice
	"""
	stop = len(L)
	return enumerate( itertools.islice(L, start, stop), start )

def reversed_enumerate_from( L , start):
	"""
	enumberate slice in reversed order
	"""
	return zip( reversed(xrange(start+1)), reversed(L[:start+1]) ) 

def Color2Array( color ):
	"""
	Return c_int array from color, color can be a number of tuple of numbers
	"""
	if isinstance(color, tuple):
		return (c_int*len(color))(*color)
	return (c_int*1)(color)

def Colors2Array( color1, color2 ):
	"""
	Return c_int array from color1 and color2, both can be a number of tuple of numbers
	"""
	if not isinstance(color1, tuple):
		color1 = (color1,)
	if not isinstance(color2, tuple):
		color2 = (color2,)
	color = color1+color2
	return (c_int*len(color))(*color)

def Arg2Tuple(arg, default = None):
	if arg is None:
		if default is None:
			return None
		else:
			arg = default
	if isinstance(arg, tuple):
		return tuple(arg)
	return tuple((arg,))

def Color2Tuple(color, default = None):
	return Arg2Tuple(color, default)

def FindNearestColor(hwnd, x, y, colors ):
	global sibiac
	n = len(colors)
	ccolors = (c_int*n)(*colors)
	if not sibiac:
		return -1
	return sibiac.NearestColor(hwnd, c_int(x), c_int(y), c_int(n), byref(ccolors))

def FindInXRight(hwnd, x, y, bg_colors, fg_colors):
	global sibiac
	if not sibiac:
		return -1
	return sibiac.FindColorInX(hwnd, c_int(x), c_int(y), c_int(len(bg_colors)), (c_int*len(bg_colors))(*bg_colors), c_int(len(fg_colors)), (c_int*len(fg_colors))(*fg_colors))

def FindInYDown(hwnd, x, y, bg_colors, fg_colors):
	global sibiac
	if not sibiac:
		return -1
	return sibiac.FindColorInY(hwnd, c_int(x), c_int(y), c_int(len(bg_colors)), (c_int*len(bg_colors))(*bg_colors), c_int(len(fg_colors)), (c_int*len(fg_colors))(*fg_colors))

def FindInYRange(hwnd, x, y, bg_colors, fg_colors):
	"""
	Search specified fg_colors up and down, return upper position and lower position, or a pair of None
	"""
	global sibiac
	if not sibiac:
		return (None, None)
	bg_len = len(bg_colors)
	bg = (c_int*bg_len)(*bg_colors)
	fg_len = len(fg_colors)
	fg = (c_int*fg_len)(*fg_colors)
	down = sibiac.FindColorInY(hwnd, c_int(x), c_int(y), c_int(fg_len), fg, c_int(bg_len), bg) # it search the length of first colors!
	up = sibiac.FindColorInYUp(hwnd, c_int(x), c_int(y), c_int(fg_len), fg, c_int(bg_len), bg) # it search the length of first colors!
	if up < 1 or down < 1:
		return (None, None)
	return (y - up + 1, y + down - 1)

def FindRow(hwnd, rect, bg_colors, fg_colors):
	"""
	Return the length and the first y in the row.
	The length is None in case of errors, 0 means there is no rows in the rectangle.
	"""
	global sibiac
	if not sibiac:
		return (None, None)
	r = RECT(*rect)
	row = c_int(0)
	height = sibiac.FindRow(hwnd, byref(r),  c_int(len(bg_colors)), (c_int*len(bg_colors))(*bg_colors), c_int(len(fg_colors)), (c_int*len(fg_colors))(*fg_colors), byref(row))
	if height < 0:
		return None
	return (height, row.value)

def FindVRange(hwnd, x, y0, y1, bg_colors, fg_colors):
	global sibiac
	if not sibiac:
		return (-1, -1)
	r = RECT(x, y0, x, y1)
	bg_colors = Color2Tuple(bg_colors)
	fg_colors = Color2Tuple(fg_colors)
	allc = bg_colors + fg_colors
	n_colors = len(allc)
	fg_idx = len(bg_colors)
	c = (c_int*n_colors)(*allc)
	range = sibiac.FindVRange(hwnd, byref(r), c_int(n_colors), c_int(fg_idx), byref(c))
	if range < 1:
		return (-1, -1)
	return (r.top, r.bottom)

def FindHRange(hwnd, x0, x1, y, bg_colors, fg_colors):
	global sibiac
	if not sibiac:
		return (-1, -1)
	r = RECT(x0, y, x1, y)
	bg_colors = Color2Tuple(bg_colors)
	fg_colors = Color2Tuple(fg_colors)
	allc = bg_colors + fg_colors
	n_colors = len(allc)
	fg_idx = len(bg_colors)
	c = (c_int*n_colors)(*allc)
	range = sibiac.FindHRange(hwnd, byref(r), c_int(n_colors), c_int(fg_idx), byref(c))
	if range < 1:
		return (-1, -1)
	#log.error("HRange: (%d - %d) = %d" % (r.left, r.right, range))
	return (r.left, r.right)

def FindHSegment(hwnd, x, y, bg_colors, fg_colors):
	""" return (x1, x2) coordinates of hirozontal segment in fg which includes specified point, or (None, None) """
	global sibiac
	if not sibiac:
		return (None, None)
	bg_colors = Color2Array(bg_colors)
	fg_colors = Color2Array(fg_colors)
	left = c_int(0)
	right = c_int(1)
	if sibiac.FindSegmentInX(hwnd, c_int(x), c_int(y), c_int(len(bg_colors)), byref(bg_colors), c_int(len(fg_colors)), byref(fg_colors), byref(left), byref(right)) == 0:
		return (None, None)
	return (left.value, right.value)
	
def PixelColor(hwnd, x, y):
	global sibiac
	if not sibiac:
		return 0
	return sibiac.PixelColor(hwnd, c_int(x), c_int(y))

#
# From what I could find, lengthy operations are not foreseen in NVDA 
# So I organize my own base class
#
import wx

class TimedCaller(wx.Timer):
	def CallAfter(self, delay, callable, *args, **kwargs):
		self.callable = callable
		self.args = args
		self.kwargs = kwargs
		return self.Start(delay, wx.TIMER_ONE_SHOT)
		
	# Stop can be used directly

	def Notify(self):
		try:
			self.callable(*self.args, **self.kwargs)
		except Exception as e:
			log.error(str(e))

try:
    # Windows >= Vista
    _windowFromPoint = user32.WindowFromPhysicalPoint
except AttributeError:
    _windowFromPoint = user32.WindowFromPoint

def WindowFromPoint(hwnd, x, y):
	p = POINT(x, y)
	user32.ClientToScreen(hwnd, byref(p))
	return _windowFromPoint(p)
	
# How to scale Y of configuration points
FIXED = 0
MOVE = 1
PROPORTIONAL = 2

def MouseSlowLeftClick():
	"""
	Generate "slow" left mouse button Down and Up (without moving it)
	"""
	global sibiac
	if not sibiac:
		return False;
	sibiac.MouseLeft(c_int(1))
	time.sleep(0.05)
	return sibiac.MouseLeft(c_int(0))

def MouseScroll(delta):
	"""
	Scroll at current mouse position
	"""
	global sibiac
	if not sibiac:
		return False
	return sibiac.MouseScroll(c_int(delta))
	
class Pt(object):
	"""
	Represent configuration point, with attributes
	how this point should react on size difference of SIBUI
	"""
	def __init__(self, x, y, scale_mode = FIXED, scale_factor = 1):
		self.x = x
		self.y = y
		self.scale_mode = scale_mode
		self.scale_factor = scale_factor
		
	def copy(self):
		return Pt(self.x, self.y, self.scale_mode, self.scale_factor)

class XY(object):
	"""
	Simple coordinates
	"""
	def __init__(self, x, y):
		self.x = x
		self.y = y

class MXY(XY):
	"""
	A point for mouse operations
	"""
	def __init__(self, hwnd, x, y, getShift = None):
		super(MXY,self).__init__(x, y)
		self.hwnd = hwnd
		self.getShift = getShift
	
	def __getXY(self):
		if self.getShift is None:
			return (self.x, self.y)
		dx, dy = self.getShift()
		return self.x + dx, self.y + dy
		
	def leftClick(self):
		global sibiac
		if not sibiac:
			return False;
		x, y = self.__getXY()
		#log.error("MXY: (%d,%d)\n", x,y)
		return sibiac.MouseLeftClick(self.hwnd, c_int(x), c_int(y))
	
	def leftDblClick(self):
		self.leftClick()
		time.sleep(0.01)
		self.leftClick()
	
	def rightClick(self):
		global sibiac
		if not sibiac:
			return False;
		x, y = self.__getXY()
		return sibiac.MouseRightClick(self.hwnd, c_int(x), c_int(y))

	def moveTo(self):
		global sibiac
		if not sibiac:
			return False;
		x, y = self.__getXY()
		sibiac.MouseMove(self.hwnd, c_int(x), c_int(y))

	def leftDrag(self, dx, dy, timeout = 0.3, mouse_up = True):
		""" Dangerous! When mouse_up is false, should be paired with leftUp !!!"""
		global sibiac
		if not sibiac:
			return 
		x, y = self.__getXY()
		sibiac.MouseMove(self.hwnd, c_int(x), c_int(y))
		winUser.mouse_event(winUser.MOUSEEVENTF_LEFTDOWN,0,0,None,None)
		time.sleep(timeout)
		sibiac.MouseMove(self.hwnd, c_int(x+dx), c_int(y+dy))
		if mouse_up:
			winUser.mouse_event(winUser.MOUSEEVENTF_LEFTUP,0,0,None,None)

	def mouseScroll(self, delta):
		self.moveTo()
		MouseScroll(delta)
		
	def leftDown(self):
		""" Dangerous! Should be paired with leftUp !!!"""
		self.moveTo()
		winUser.mouse_event(winUser.MOUSEEVENTF_LEFTDOWN,0,0,None,None)

	def leftUp(self):
		""" Dangerous! Should be paired with leftDown !!!"""
		winUser.mouse_event(winUser.MOUSEEVENTF_LEFTUP,0,0,None,None)
		
	def FindNearestColor(self, colors):
		x, y = self.__getXY()
		return FindNearestColor(self.hwnd, x, y, colors)

	def PixelColor(self):
		x, y = self.__getXY()
		return PixelColor(self.hwnd, x, y)

	def FindHSegment(self, bg, fg):
		x, y = self.__getXY()
		return FindHSegment(self.hwnd, x, y, bg, fg)
	
	def getXY(self):
		x, y = self.__getXY()
		return XY(x, y)
		
class Box(object):
	"""
	Basic Sibiac object, represet a box on screen and provides
	Sibiac (binary library) methods for it
	
	The box can be dynamic, getBox method always return static, possibly recalculated, box
	"""
	def __init__(self, hwnd, left, top, right = None, bottom = None, getShift = None):
		"""
		@param hwnd: Window handler for target window
		@type hwnd: HWND
		@param left, top, right, buttom: box coordinates in hwnd, inclusive
		@type left, top, right, buttom: int
		"""
		super(Box,self).__init__()
		self.hwnd = hwnd
		self.left = left
		self.top = top
		if right is None or bottom is None:
			self.right = left
			self.bottom = top
		else:
			self.right = right
			self.bottom = bottom
		self.vshift = 0
		self.getShift = getShift
	
	def copy(self, getShift = None):
		if getShift is None:
			getShift = self.getShift
		return Box(self.hwnd, self.left, self.top, self.right, self.bottom, getShift)
		
	def getBox(self):
		"""
		Return self or static box with applied shift
		"""
		if self.getShift is None:
			return self
		dx, dy = self.getShift()
		return Box(self.hwnd, self.left + dx, self.top + dy, self.right + dx, self.bottom + dy)

	def vShift(self, vshift = 0):
		"""
		Shift box vertical position
		"""
		self.vshift = vshift

	def getText(self, ch = None):
		"""
		Get text from the box (OCR, slow)
		"""
		global sibiac
		if not sibiac:
			return "";
		buf = create_string_buffer(1024)
		if ch is None:
			sibiac.Screen2Text(self.hwnd, c_int(self.left), c_int(self.top + self.vshift), c_int(self.right), c_int(self.bottom + self.vshift), buf, len(buf))
		else:
			sibiac.Screen2TextCh(self.hwnd, c_int(ch), c_int(self.left), c_int(self.top + self.vshift), c_int(self.right), c_int(self.bottom + self.vshift), buf, len(buf))
		return u2ascii(buf.value.decode('utf-8', 'ignore'))

	def getCRC(self):
		"""
		Get CRC32 sum of the box
		"""
		global sibiac
		if not sibiac:
			return "";
		return sibiac.Screen2CRC(self.hwnd, c_int(self.left), c_int(self.top + self.vshift), c_int(self.right), c_int(self.bottom + self.vshift))

	def getTextOut(self):
		"""
		NVDA collects text outputs. If GUI is build this way, we can use it instead of OCR
		"""
		import displayModel
		obj = api.getForegroundObject()
		if not obj:
			return Box.getText(self) # can produce endless recursion with TextOutBox
		r = ClientRectToScreen(self.hwnd, RECT(self.left, self.top + self.vshift, self.right, self.bottom + self.vshift))
		return displayModel.DisplayModelTextInfo(obj, textInfosRect(r.left, r.top, r.right, r.bottom)).text

	def getPositionOrCenter(self, x = None, y = None):
		"""
		Return (x,y) or coordinates of the center of the box
		@param x, y: point coordinates or None
		@param x, y: int
		"""
		if x is None:
			x = self.left + (self.right - self.left)//2
		if y is None:
			y = self.top + self.vshift + (self.bottom - self.top)//2
		return (x, y)
		
	def leftClick(self, x = None, y = None):
		"""
		Click at specified coordinates or in the center of the box
		@param x, y: point coordinates
		@param x, y: int
		"""
		global sibiac
		if not sibiac:
			return False;
		x, y = self.getPositionOrCenter(x, y)
		return sibiac.MouseLeftClick(self.hwnd, c_int(x), c_int(y))

	def leftDrag(self, timeout = 0.3):
		global sibiac
		if not sibiac:
			return 
		sibiac.MouseMove(self.hwnd, c_int(self.left), c_int(self.top))
		winUser.mouse_event(winUser.MOUSEEVENTF_LEFTDOWN,0,0,None,None)
		time.sleep(timeout)
		sibiac.MouseMove(self.hwnd, c_int(self.right), c_int(self.bottom))
		winUser.mouse_event(winUser.MOUSEEVENTF_LEFTUP,0,0,None,None)
		
	def leftClickX(self, x= None, y = None):
		self.moveTo(x, y)
		winUser.mouse_event(winUser.MOUSEEVENTF_LEFTDOWN,0,0,None,None)
		winUser.mouse_event(winUser.MOUSEEVENTF_LEFTUP,0,0,None,None)
		
	def rightClick(self, x = None, y = None):
		"""
		Right click at specified coordinates or in the center of the box
		@param x, y: point coordinates
		@param x, y: int
		"""
		global sibiac
		if not sibiac:
			return False;
		x, y = self.getPositionOrCenter(x, y)
		return sibiac.MouseRightClick(self.hwnd, c_int(x), c_int(y))

	def moveTo(self, x = None, y = None):
		"""
		Move mouse to specified coordinates or in the center of the box
		@param x, y: point coordinates
		@param x, y: int
		"""
		global sibiac
		if not sibiac:
			return False;
		x, y = self.getPositionOrCenter(x, y)
		sibiac.MouseMove(self.hwnd, c_int(x), c_int(y))
		# For some reason, that does not always work...
		#x, y = getScreenPos(self.hwnd, x, y)
		#log.error("Cursor target "+str(x)+","+str(y))
		#winUser.setCursorPos(x, y)
	
	def pictureChange(self, save_ref = 0):
		"""
		Return -1 (error), 0 (no changes), 3 (changes, but current equal to reference), 2 (the boundary of changes with ref are equal to changes with last), 1 (all other change cases)
		and the box of changes in the picture since the last call  ( for positive return code )
		"""
		global sibiac
		if not sibiac:
			return (-1, None);
		box = self.getBox()
		r = RECT(box.left, box.top, box.right, box.bottom)
		ret = sibiac.DetectPictureChange(self.hwnd, byref(r), c_int(save_ref))
		if ret < 1:
			return (ret, None)
		return (ret, Box(box.hwnd, r.left, r.top, r.right, r.bottom))
	
	def RECT(self):
		box = self.getBox()
		return RECT(box.left, box.top, box.right, box.bottom)
	
	def pictureCRC(self, masks = None):
		"""
		Capture box, apply masks and calculate CRC. Masks should be an iterable with RECTs
		"""
		global sibiac
		if not sibiac:
			return None;
		nmasks = 0
		rmasks = RECT()
		if masks is not None:
			if isinstance(masks, RECT):
				masks = (masks,)
			nmasks = len(masks)
			rmasks = (RECT*nmasks)(*masks)
		box = self.getBox()
		r = self.RECT()
		return sibiac.MaskedPictureCRC(box.hwnd, byref(r), nmasks, byref(rmasks))
	
	def includePoint(self, xy):
		"""
		extend the box boundary to include specified point
		"""
		if self.left > xy.x:
			self.left = xy.x
		if self.right < xy.x:
			self.right = xy.x
		if self.top > xy.y:
			self.top = xy.y
		if self.bottom < xy.y:
			self.bottom = xy.y
	
	def resize(self, dl, dt, dr, db, limit = None):
		"""
		return resizes box by specified number of pixels. If <limit> is specified, resulting box will be within it.
		"""
		box = self.getBox().copy()
		box.left += dl
		box.top += dt
		box.right += dr
		box.bottom += db
		if isinstance(limit, Box):
			if box.left < limit.left:
				box.left = limit.left
			if box.top < limit.top:
				box.top = limit.top
			if box.right > limit.right:
				box.right = limit.right
			if box.bottom > limit.bottom:
				box.bottom = limit.bottom
		return box
	
	def __eq__(self, other):
		if not isinstance(other, Box):
			return False
		if self.left == other.left and self.right == other.right and self.top == other.top and self.bottom == other.bottom and self.vshift == other.vshift and self.getShift == other.getShift:
		   return True
		return False
	
	def __neq__(self, other):
		return not self.__eq__(other)

	def __str__(self):
		box = self.getBox()
		return "(%d,%d) - (%d,%d)" % (box.left, box.top, box.right, box.bottom)

class TextBox(Box):
	"""
	A Box which saves hashes for already returned getText()
	getText() method workes for shifted boxes.
	"""
	def __init__(self, hwnd, left, top, right = None, bottom = None, getShift = None, old = None):
		super(TextBox,self).__init__(hwnd, left, top, right, bottom, getShift)
		self.text_hash = {}
		if self == old:
			self.text_hash = old.text_hash

	def invalidate(self):
		pass # we do not save the text, just CRC hashes

	def getText(self, ch = None):
		box = self.getBox()
		crc32 = box.getCRC()
		if crc32 in self.text_hash:
			#log.error("Hashed %x : %s" % (crc32,self.text_hash[crc32]) )
			return self.text_hash[crc32]
		if box == self:
			text = super(TextBox,self).getText(ch)
		else:
			text = box.getText(ch)
		self.text_hash[crc32] = text
		#log.error("Not hashed %x : %s" % (crc32,self.text_hash[crc32]) )
		return text
		
class TextOutBox(TextBox):
	"""
	TextBox with tries to get the text using getTextOut method.
	NVDA in some apps "forget" the text, so it is cached. 
	It needs invalidate to drop the cache when the content could be changed.
	"""
	def __init__(self, hwnd, left, top, right = None, bottom = None, getShift = None, old = None):
		super(TextOutBox,self).__init__(hwnd, left, top, right, bottom, getShift, old)
		if self == old:
			self.last_text = old.last_text
		else:
			self.last_text = None

	def invalidate(self):
		super(TextOutBox,self).invalidate()
		self.last_text = None

	def getText(self, ch = None):
		text = self.getTextOut()
		if text != "":
			self.last_text = text
			return text
		if self.last_text is None:
			self.last_text = super(TextOutBox, self).getText(ch)
		return self.last_text
		
	def __eq__(self, other):
		if not isinstance(other, TextOutBox):
			return False
		return super(TextOutBox,self).__eq__(other)
	
class PopupBox(Box):
	"""
	Colored popup box with fixed top and bottom but dynamic width. 
	Note that isVisible should be called prior any other functions by used control, to avoid stocked isVisible.
	"""
	def __init__(self, hwnd, x, top, bottom, bg, fg):
		super(PopupBox,self).__init__(hwnd, x, top, x, bottom)
		self.x = x
		self.bg = Color2Array(bg)
		self.fg = Color2Array(fg)
	
	def isVisible(self):
		global sibiac
		left = c_int(0)
		right = c_int(0)
		if not sibiac or sibiac.FindSegmentInX(self.hwnd, c_int(self.x), c_int(self.top + self.vshift),
									c_int(len(self.bg)), byref(self.bg), c_int(len(self.fg)),byref(self.fg),
									byref(left), byref(right)) == 0:
			return False
		self.left = left.value
		self.right = right.value
		return True
		
class Scroll(Box):
	"""
	Vertical or horizontal scroll element (width or height should be 1). 
	"""
	def __init__(self, hwnd, left, top, right, bottom, bg, fg, getShift = None, delta = 120):
		super(Scroll,self).__init__(hwnd, left, top, right, bottom, getShift)
		self.delta = delta
		bg  = Color2Tuple(bg)
		colors = bg + Color2Tuple(fg)
		self.n_colors = c_int(len(colors))
		self.fg_idx = c_int(len(bg))
		self.colors = (c_int*len(colors))(*colors)
				
class ScrollV(Scroll):
	""" Vertical scroll box """
	
	def __scroll(self, delta):
		global sibiac
		if not sibiac:
			return False
		box = self.getBox()
		r = RECT(box.left, box.top, box.right, box.bottom)
		range = sibiac.FindVRange(self.hwnd, byref(r), self.n_colors, self.fg_idx, byref(self.colors))
		if range < 1 or (delta > 0 and r.top == box.top) or (delta < 0 and r.bottom == box.bottom) :
			return False
		box.moveTo()
		MouseScroll(delta)
		return True
	
	def scrollUp(self):
		return self.__scroll(self.delta)

	def scrollDown(self):
		return self.__scroll(-self.delta)	

class VScroll(Box):
	"""
	Vertial scroll control, with position mover. Detects it exists and can be scrolled up or down.
	"""
	def __init__(self, hWnd, x, top, bottom, bg_color, pos_color):
		super(VScroll,self).__init__(hWnd, x, top, x, bottom)
		self.bg_color = Color2Array( bg_color )
		self.pos_color = Color2Array( pos_color )
		
	def canScrollUp(self):
		global sibiac
		if not sibiac:
			return False
		# we are measuring bg_color on top. If the whole line is bg_color, there is no pos... 
		i = sibiac.FindColorInY(self.hwnd, c_int(self.left), c_int(self.top),
							  c_int(len(self.bg_color)),byref(self.bg_color),
							  c_int(len(self.pos_color)),byref(self.pos_color))
		if i <= 0:
			return False # Failed or pos color at the top
		if i < self.bottom - self.top:
			return True # Position box was found
		return False # Position box was not found

	def canScrollDown(self):
		global sibiac
		if not sibiac:
			return False
		# we are measuring bg_color on bottom. If the whole line is bg_color, there is no pos... 
		i = sibiac.FindColorInYUp(self.hwnd, c_int(self.right), c_int(self.bottom),
							  c_int(len(self.bg_color)),byref(self.bg_color),
							  c_int(len(self.pos_color)),byref(self.pos_color))
		if i <= 0:
			return False # Failed or pos color at the bottom
		if i < self.bottom - self.top:
			return True # Position box was found
		return False # Position box was not found

class VListBox(Box):
	"""
	A box with selectable text list 
	"""
	def _calculateNItems(self):
		"""
		Recalculate n_items from last_bottom
		"""
		h = self.last_bottom - self.top + 1
		if self.item_h == 0 or h < self.box_h:
			self.n_items = 0
		else:
			self.n_items = int(1 + (h - self.box_h) / self.item_h)
	
	def __init__(self, hwnd, first_left, first_top, first_right, first_bottom, second_top, last_bottom, inactive_color, active_color):	
		super(VListBox,self).__init__(hwnd, first_left, first_top, first_right, first_bottom)
		
		self.second_top = second_top
		self.last_bottom = last_bottom
		self.inactive_color = Color2Array(inactive_color)
		self.active_color = Color2Array(active_color)
		self.item_hash = {}

		self.box_h = self.bottom - self.top + 1
		self.item_h = self.second_top - self.top
		self._calculateNItems()

	def setHeight(self, height):
		"""
		Set new last_bottom if required
		"""
		last_bottom = self.top + height - 1
		if self.last_bottom == last_bottom:
			return
		self.last_bottom = last_bottom
		self._calculateNItems()			
			
	def selection(self):
		global sibiac
		if not sibiac:
			return -1
		shift = c_int(0)
		i = sibiac.VListSelection(self.hwnd, c_int(self.left), c_int(self.top + self.vshift), c_int(self.bottom + self.vshift), 
								  c_int(self.second_top + self.vshift), c_int(self.last_bottom + self.vshift), 
								  c_int(len(self.inactive_color)),byref(self.inactive_color),
								  c_int(len(self.active_color)),byref(self.active_color),
								  byref(shift))
		return (int(i), shift.value)

	def getText(self, n = -1, shift = 0):
		global sibiac
		if n < 0 or n >= self.n_items:
			n, shift = self.selection()
		if n < 0:
			return ""
		top = self.top + self.vshift + self.item_h*n + shift
		box = Box(self.hwnd, self.left, top, self.right, top + self.box_h)
		crc32 = box.getCRC()
		if crc32 in self.item_hash:
			return self.item_hash[crc32]
		text = box.getText()
		self.item_hash[crc32] = text
		return text

	def leftClick(self, n = -1, shift = 0):
		if n < 0 or n >= self.n_items:
			n, shift  = self.selection()
		if n < 0:
			return ""
		top = self.top + self.vshift + self.item_h*n + shift
		# log.error("%d %d %d %d" % (n, shift, self.n_items, self.box_h))
		box = Box(self.hwnd, self.left, top, self.right, top + self.box_h)
		return box.leftClick()

	def moveTo(self, n = -1, shift = 0):
		if n < 0 or n >= self.n_items:
			n, shift  = self.selection()
		if n < 0:
			return ""
		top = self.top + self.vshift + self.item_h*n + shift
		box = Box(self.hwnd, self.left, top, self.right, top + self.box_h)
		return box.moveTo()

class PopupMenu(Box):
	"""
	A box with selectable text list 
	"""
	def _calculateNItems(self):
		"""
		Recalculate n_items from last_bottom
		"""
		h = self.last_bottom - self.top + 1
		if self.item_h == 0 or h < self.box_h:
			self.n_items = 0
		else:
			self.n_items = int(1 + (h - self.box_h) / self.item_h)
	
	def __init__(self, hwnd, nw_x, nw_y, se_x, se_y, left_border, right_border, top_border, bottom_border, box_h, item_h, inactive_color, active_color, selected_color):
		"""
		nw and se - top left and right bottom (inclusive) corners coordinates, can be changed later with updateGeometry methods
		borders are relative to the corners position to selectable boxes
		"""
		self.type = "Menu"
		self.left_border = left_border
		self.right_border = right_border
		self.top_border = top_border
		self.bottom_border = bottom_border
		self.box_h = box_h
		self.item_h = item_h
		if selected_color:
			self.not_current_color = Colors2Array(inactive_color, selected_color)
			self.current_color = Color2Array(active_color)
			self.not_selected_color = Colors2Array(inactive_color, active_color)
			self.selected_color = Color2Array(selected_color)
		else:
			self.not_current_color = Color2Array(inactive_color)
			self.current_color = Color2Array(active_color)
			self.not_selected_color = None
			self.selected_color = None
		self.item_hash = {}

		super(PopupMenu,self).__init__(hwnd, nw_x + left_border, nw_y + top_border, se_x - right_border, nw_y + top_border + box_h - 1)
		self.last_bottom = se_y - bottom_border
		self._calculateNItems()

	def updateGeometry(self, nw_x, nw_y, se_x, se_y):
		self.left = nw_x + self.left_border
		self.top = nw_y + self.top_border
		self.right = se_x - self.right_border
		self.bottom = nw_y + self.top_border + self.box_h - 1
		self.last_bottom = se_y - self.bottom_border
		self._calculateNItems()
		
	def setHeight(self, height):
		"""
		Set new last_bottom if required, the height is from the first item top
		"""
		last_bottom = self.top + height - 1
		if self.last_bottom == last_bottom:
			return
		self.last_bottom = last_bottom
		self._calculateNItems()			
			
	def current(self):
		global sibiac
		if not sibiac:
			return -1
		shift = c_int(0)
		i = sibiac.VListSelection(self.hwnd, c_int(self.left), c_int(self.top + self.vshift), c_int(self.bottom + self.vshift), 
								  c_int(self.top + self.item_h + self.vshift), c_int(self.last_bottom + self.vshift), 
								  c_int(len(self.not_current_color)),byref(self.not_current_color),
								  c_int(len(self.current_color)),byref(self.current_color),
								  byref(shift))
		if self.selected_color:
			j = sibiac.VListSelection(self.hwnd, c_int(self.left), c_int(self.top + self.vshift), c_int(self.bottom + self.vshift), 
									c_int(self.top + self.item_h + self.vshift), c_int(self.last_bottom + self.vshift), 
									c_int(len(self.not_selected_color)),byref(self.not_selected_color),
									c_int(len(self.selected_color)),byref(self.selected_color),
									byref(shift))
		else:
			j = -1
		i = int(i)
		j = int(j)
		if i < 0:
			if j >= 0:
				return (j, shift.value, True)
			return (i, shift.value, False)
		return (i, shift.value, i == j)

	def getText(self, n = -1, shift = 0):
		global sibiac
		if n < 0 or n >= self.n_items:
			n, shift, is_selected = self.current()
		if n < 0:
			return ""
		top = self.top + self.vshift + self.item_h*n + shift
		box = Box(self.hwnd, self.left, top, self.right, top + self.box_h)
		crc32 = box.getCRC()
		if crc32 in self.item_hash:
			return self.item_hash[crc32]
		text = box.getText()
		self.item_hash[crc32] = text
		return text

	def leftClick(self, n = -1, shift = 0):
		if n < 0 or n >= self.n_items:
			n, shift  = self.current()
		if n < 0:
			return ""
		top = self.top + self.vshift + self.item_h*n + shift
		box = Box(self.hwnd, self.left, top, self.right, top + self.box_h)
		return box.leftClick()

	def moveTo(self, n = -1, shift = 0):
		if n < 0 or n >= self.n_items:
			n, shift  = self.current()
		if n < 0:
			return ""
		top = self.top + self.vshift + self.item_h*n + shift
		box = Box(self.hwnd, self.left, top, self.right, top + self.box_h)
		return box.moveTo()

class ColorMatcherObj(object):
	def __init__(self, colors, name = None):
		self.match_colors = colors
		if name is not None:
			self.name = name
		
class ColorMatcher(object):
	"""
	Match specified list of coordinates against the list of objects.
	Objects should have match_colors property, with colors for specified coordinates.
	Can return None when none of the objects have matched
	
	Points should be absolute
	"""
	
	def __init__(self, hwnd, points, objs, getShift = None):
		super(ColorMatcher,self).__init__()
		self.hwnd = hwnd
		self.__getShift = getShift
		self.nobjs = len(objs)
		MatchColorArray = c_int * self.nobjs

		self.objs = objs
		self.points = []
		for pt in points:
			xy = XY( pt.x, pt.y )
			xy.colors = MatchColorArray()
			self.points.append( xy )

		for i, obj in enumerate(objs):
			for j, color in enumerate(obj.match_colors):
				self.points[j].colors[i] = color

	def getShift(self):
		if self.__getShift is not None:
			return self.__getShift()
		else:
			return (0, 0)
	
	def match(self):
		global sibiac
		colors = []
		n = None
		for j, pt in enumerate(self.points):
			dx, dy = self.getShift()
			i = sibiac.NearestColor(self.hwnd, c_int(pt.x + dx), c_int(pt.y + dy), c_int(self.nobjs), byref(pt.colors))
			if i < 0:
				return None
			if n is None:
				n = i
			elif n == i:
				pass # the same
			elif self.objs[n].match_colors[j] == self.objs[i].match_colors[j]:
				pass # existing one is also ok
			else:
				search_other = False
				# if new one is ok, choose it.
				for k in range(0, j):
					if self.objs[n].match_colors[k] != self.objs[i].match_colors[k]:
						search_other = True
						break
				if not search_other:
					n = i
				else: 
					# n is not good and i is also not good...
					# find the first object which has first j colors from n and j-th color from i
					objn = self.objs[n]
					obji = self.objs[i]
					n = None
					for m, obj in enumerate(self.objs):
						bad = False
						for k in range(0, j):
							if obj.match_colors[k] != objn.match_colors[k]:
								bad = True
								break
						if not bad and obj.match_colors[j] == obji.match_colors[j]:
							n = m
							break
					if n is None:
						#color = sibiac.PixelColor(self.hwnd, c_int(pt.x), c_int(pt.y) )
						#log.error("%d %d %d %d 0x%x 0x%x 0x%x" % (j, i, n, k, self.objs[n].match_colors[k], self.objs[i].match_colors[k], color))
						return None
		if n is None:
			return None
		return self.objs[n]
		
	def log(self):
		global sibiac
		# print actual colors into the log		
		colors = []
		dx, dy = self.getShift()
		for pt in self.points:
			colors.append("0x%x" % ( sibiac.PixelColor(self.hwnd, c_int(pt.x + dx), c_int(pt.y + dy) ) ))
		#log.error(', '.join(colors))		
		
class Control(object):
	"""
	Any control in SIBI
	"""
	
	def __init__(self, name, sibi, opt = None):
		"""
		@param name: the name of the element
		@type name: string
		@param sibi: SIBI
		@type sibi: SIBI
		@param opt: set of options (control type specific)
		@type opt: set
		"""
		super(Control,self).__init__();
		self.name = name
		self.sibi = sibi # should be weakref, but I do not care at the moment
		if opt is None:
			self.opt = set()
		else:
			self.opt = set(opt)
		self.type = '' # to be overloaded

	# Control Interface
	def expectFocusReturn(self):
		""" focus will leave Sibiac, but that is a part of the interface. So cancel speack on return. """
		self.sibi.expectFocusReturn()
	
	def speak(self, text, cancel = True):
		if cancel:
			speech.cancelSpeech()
		queueHandler.queueFunction(queueHandler.eventQueue, speech.speakMessage, text)
	
	def speakAfter(self, delay = None):
		if delay is None:
			delay = self.reactionTime()
		self.sibi.speakAfter(delay)
		return True

	def speakFocusAfter(self, delay = None):
		if delay is None:
			delay = self.reactionTime()
		self.sibi.speakFocusAfter(delay)
		return True

	def speakInFocusAfter(self, delay = None):
		if delay is None:
			delay = self.reactionTime()
		self.sibi.speakInFocusAfter(delay)
		return True
		
	def getTextInfo(self):
		"""
		Should return (CtrlType, CtrlName, CtrlText) tuple
		"""
		return (self.type, self.name, '')

	def getNameInGroup(self):
		"""
		Should return Name on (Name, Text) tuple.
		It used when the control is in a group, to annouce when the group select it
		"""
		return self.name
		
	def isFocusable(self):
		"""
		Should return True in case the control can get focus.
		That is not always the case, f.e. if under some conditions the control does not exist...
		"""
		return True

	def isEnabled(self):
		"""
		Should return True in case the control has "on" state. Used with "switchable" containers.
		"""
		return True
		
	def focusSet(self):
		"""
		Inform the control it got the focus
		"""
		self.sibi.focusChangedTo(self)

	def getFocused(self):
		"""
		Return focused containing control, by default control itself
		"""
		return self
		
	def focusLost(self):
		"""
		Inform the control it has lost the focus, note that it does not triggered externally
		So general control should not expect it is notified on focus lose 
		"""
		pass

	def focusFirst(self):
		"""
		Focus the first containing control, by default control itself
		"""
		self.focusSet()
		return True

	def focusLast(self):
		"""
		Focus the last containing control, by default control itself
		"""
		self.focusSet()
		return True

	def focusNext(self):
		"""
		Focus next containing control
		"""
		return False

	def focusPrevious(self):
		"""
		Focus previous containing control
		"""
		return False
		
	def reactionTime(self):
		"""
		Defines how fast control is "reacting" on operations
		"""
		if "slow_reaction" in self.opt:
			return 300
		if "delayed_reaction" in self.opt:
			return 500
		if "load_reaction" in self.opt:
			return 1000
		return 100
				
	# All keyboard processors should return False in case the gesture should be sent "as is"
		
#	def onUp(self):
#		return False

#	def onDown(self):
#		return False

#	def onLeft(self):
#		return False

#	def onRight(self):
#		return False

#	def onEnter(self):
#		return False
	
#	def onEscape(self):
#		return False
	
	def onPageUp(self):
		if "page_keys" in self.opt:
			self.sibi.speakAfter(self.reactionTime())
		return False

	def onPageDown(self):
		if "page_keys" in self.opt:
			self.sibi.speakAfter(self.reactionTime())
		return False

	def onHome(self):
		if "homeend_keys" in self.opt:
			self.sibi.speakAfter(self.reactionTime())
		return False

	def onEnd(self):
		if "homeend_keys" in self.opt:
			self.sibi.speakAfter(self.reactionTime())
		return False
		
class Container(Control):
	"""
	General control which can have other controls, also can have a list of pop-up dialogs
	"""
	def __init__(self, name, sibi, opt):
		"""
		@param name: the name of the element
		@type name: string
		@param sibi: SIBI
		@type sibi: SIBI
		@param opt: set of options ("switchable")
		@type opt: set
		"""
		super(Container,self).__init__(name, sibi, opt);
		self.ctrl = [] # flat list of controls
		self.ctrl_idx = 0 # current control index, the first is default
		self.dlg = [] # a list of pop-up dialogs
		self.last_dlg = None # was some dialog open last time we have checked?

	def add(self, ctrl):
		"""
		Add new control to the Container.
		@param ctrl: Control object to add
		@type ctrl: Control derived, but may be container specific
		"""
		if not isinstance(ctrl, Control):
			log.error( "Bug... not control in container "+str(ctrl) )
			return
		self.ctrl.append( ctrl )
		
	def addDialog(self, dlg):
		"""
		Add new dialog
		@param dlg: dialog like control, with isFocusable and cyclic focusing
		@type dlg: Control derived, b
		"""
		self.dlg.append( dlg )

	def _getDialog(self):
		if not self.dlg:
			return None
		if self.last_dlg:
			if self.last_dlg.isFocusable():
				return self.last_dlg
		for dlg in self.dlg:
			if dlg.isFocusable():
				self.last_dlg = dlg
				dlg.focusFirst()
				return dlg
		self.last_dlg = None
		return None

	def _checkEnabled(self):
		if "switchable" in self.opt and self.ctrl and not self.ctrl[0].isEnabled():
			self.ctrl_idx = 0
			return False
		return True
	
	def getFocused(self):
		dlg = self._getDialog()
		if dlg:
			return dlg.getFocused()
		if not self.ctrl:
			return None
		self._checkEnabled()
		return self.ctrl[self.ctrl_idx].getFocused()

	def getTextInfo(self):
		"""
		Should return (CtrlType, CtrlName, CtrlText) tuple
		"""
		dlg = self._getDialog()
		if dlg:
			return dlg.getTextInfo()
		if not self.ctrl:
			return ('', self.name, 'has no controls')
		self._checkEnabled()
		return self.ctrl[self.ctrl_idx].getTextInfo()

	def isFocusable(self):
		"""
		Should return True in case the control can get focus.
		That is not always the case, f.e. if under some conditions the control does not exist...
		
		General container get no focus by itself, but other containers can
		"""
		dlg = self._getDialog()
		if dlg:
			return True
		for ctrl in self.ctrl:
			if ctrl.isFocusable():
				return True
		return False
		
	def focusSet(self):
		"""
		Inform the control it got the focus
		"""
		dlg = self._getDialog()
		if dlg:
			dlg.focusSet()
			return
		if not self.ctrl:
			return
		self._checkEnabled()
		if self.ctrl[self.ctrl_idx].isFocusable():
			self.ctrl[self.ctrl_idx].focusSet()
		else:
			self.focusFirst()

	def focusFirst(self):
		"""
		Set focus to the first element
		"""
		dlg = self._getDialog()
		if dlg:
			dlg.focusFirst()
			return True
		for i, ctrl in enumerate(self.ctrl):
			if ctrl.isFocusable():
				self.ctrl_idx = i
				if not self._checkEnabled() and i > 0:
					return False
				ctrl.focusFirst()
				return True
		return False

	def focusLast(self):
		"""
		Set focus to the last element
		"""
		dlg = self._getDialog()
		if dlg:
			dlg.focusLast()
			return True
		if not self._checkEnabled():
			return self.focusFirst()
		for i, ctrl in reversed_enumerate(self.ctrl):
			if ctrl.isFocusable():
				self.ctrl_idx = i
				ctrl.focusLast()
				return True
		return False
				
	def focusNext(self):
		"""
		Containers process Tab as long as they can focus corresponsing control.
		"""
		dlg = self._getDialog()
		if dlg:
			dlg.focusNext()
			return True
		if not self.ctrl:
			return False
		if not self._checkEnabled():
			return False
		if self.ctrl[self.ctrl_idx].focusNext():
			return True
		for i, ctrl in enumerate_from( self.ctrl, self.ctrl_idx + 1):
			if ctrl.isFocusable():
				self.ctrl_idx = i
				ctrl.focusFirst()
				return True
		return False
		
	def focusPrevious(self):
		"""
		Containers process Tab as long as they can focus corresponsing control.
		"""
		dlg = self._getDialog()
		if dlg:
			dlg.focusPrevious()
			return True
		if not self.ctrl:
			return False
		if not self._checkEnabled():
			return False
		if self.ctrl[self.ctrl_idx].focusPrevious():
			return True
		for i, ctrl in reversed_enumerate_from( self.ctrl, self.ctrl_idx - 1):
			if ctrl.isFocusable():
				self.ctrl_idx = i
				ctrl.focusLast()
				return True
		return False

class Pages(Control):
	"""
	Sometimes an interface has several pages, like tabs but without usual tab control
	"""
	def __init__(self, name, sibi, getPageIndex, opt = None):
		super(Pages,self).__init__(name, sibi, opt);
		self.ctrl = [] # flat list of controls
		self.getPageIdx = getPageIndex

	def _getCurrentPage(self):
		idx = self.getPageIdx()
		if (idx is None) or (not self.ctrl) or (idx >= len(self.ctrl)):
			return None
		return self.ctrl[idx]
		
	def add(self, ctrl):
		if not isinstance(ctrl, Control):
			log.error( "Bug... not control in container "+str(ctrl) )
			return
		self.ctrl.append( ctrl )
	
	def getFocused(self):
		ctrl = self._getCurrentPage()
		if ctrl:
			return ctrl.getFocused()
		return None

	def getTextInfo(self):
		ctrl = self._getCurrentPage()
		if ctrl:
			return ctrl.getTextInfo()
		return (self.type, self.name, "Unknown page")

	def isFocusable(self):
		ctrl = self._getCurrentPage()
		if ctrl:
			return ctrl.isFocusable()
		return False
		
	def focusSet(self):
		ctrl = self._getCurrentPage()
		if ctrl:
			ctrl.focusSet()

	def focusFirst(self):
		ctrl = self._getCurrentPage()
		if ctrl:
			return ctrl.focusFirst()
		return False

	def focusLast(self):
		ctrl = self._getCurrentPage()
		if ctrl:
			return ctrl.focusLast()
		return False
				
	def focusNext(self):
		ctrl = self._getCurrentPage()
		if ctrl:
			return ctrl.focusNext()
		return False
		
	def focusPrevious(self):
		ctrl = self._getCurrentPage()
		if ctrl:
			return ctrl.focusPrevious()
		return False

class GroupBase(Control):
	"""
	The base for Group or Tab like containers
	"""
	def __init__(self, name, sibi, opt = None):
		super(GroupBase,self).__init__(name, sibi, opt);
		self._no_ctrl_text = "no selection"
		self._element_type = "control"
		self._ctrl = [] # flat list of controls
		self._chooser = True # indicate we are in focus		

	def _getCurrent(self):
		""" Should return current control, excluding _chooser. Should be implemented """
		return None
		
	def add(self, ctrl):
		if not isinstance(ctrl, Control):
			log.error( "Bug... not control "+str(ctrl) )
			return
		self._ctrl.append( ctrl )
	
	def getFocused(self):
		if self._chooser:
			return self
		ctrl = self._getCurrent()
		if ctrl:
			focused = ctrl.getFocused()
			if focused:
				return focused
		self._chooser = True
		return self

	def getTextInfo(self):
		ctrl = self._getCurrent()
		if ctrl is None:
			self._chooser = True
			return (self.type, self.name, self._no_ctrl_text)
		if self._chooser:
			text = ctrl.getNameInGroup()
			if isinstance(text, tuple):
				return (self.name, text[0], text[1])
			return (self.type, self.name, text)
		return ctrl.getTextInfo()
		
	def focusSet(self):
		focused = self.getFocused()
		if focused == self:
			super(GroupBase,self).focusSet()
			return
		focused.focusSet()

	def focusFirst(self):
		self._chooser = True
		self.focusSet()
		return True

	def focusLast(self):
		ctrl = self._getCurrent()
		if ctrl is not None and ctrl.isFocusable() and ctrl.focusLast():
			self._chooser = False
			return True
		return self.focusFirst()
				
	def focusNext(self):
		ctrl = self._getCurrent()
		if ctrl:
			if self._chooser:
				if ctrl.isFocusable():
					self._chooser = False
					return ctrl.focusFirst()
				return False
			return ctrl.focusNext()
		return False
		
	def focusPrevious(self):
		if self._chooser:
			return False
		ctrl = self._getCurrent()
		if ctrl is not None:
			if ctrl.focusPrevious():
				return True
		return self.focusFirst()

	def onLeft(self):
		""" should be implemented """
		self.sibi.speakInFocusAfter(self.reactionTime())
		return True

	def onRight(self):
		""" should be implemented """
		self.sibi.speakInFocusAfter(self.reactionTime())
		return True
	
	def onUp(self):
		ctrl = self._getCurrent()
		if ctrl is not None and hasattr(ctrl, "onUp"):
			return ctrl.onUp()
		return True

	def onDown(self):
		ctrl = self._getCurrent()
		if ctrl is not None and hasattr(ctrl, "onDown"):
			return ctrl.onDown()
		return True

	def onEnter(self):
		ctrl = self._getCurrent()
		if ctrl is not None and hasattr(ctrl, "onEnter"):
			return ctrl.onEnter()
		return True

	def onShiftUp(self):
		ctrl = self._getCurrent()
		if ctrl is not None and hasattr(ctrl, "onShiftUp"):
			return ctrl.onShiftUp()
		return True

	def onShiftDown(self):
		ctrl = self._getCurrent()
		if ctrl is not None and hasattr(ctrl, "onShiftDown"):
			return ctrl.onShiftDown()
		return True

	def onHelp(self, prefix = "", suffix = ""):
		text = prefix + "Use left and right arrows to select %s. Then use %s type dependent keys." % (self._element_type, self._element_type)
		ctrl = self._getCurrent()
		if ctrl is not None and hasattr(ctrl, "onHelp"):
			text = text + " For %s: " % ctrl.name + " "
			ctrl.onHelp(text, suffix)
		else:
			self.speak(text + suffix)
		return True
	
class Group(GroupBase):
	"""
	Virtual switch between contained controls.
	It call current control for Up/Down/Enter, in case that control defines the reaction
	"""
	def __init__(self, name, sibi, opt = None):
		super(Group,self).__init__(name, sibi, opt);
		self._idx = None # current control index
		
	def _getCurrent(self):
		if self._idx is None:
			self._chooser = True
			return None
		return self._ctrl[self._idx]

	def add(self, ctrl):
		super(Group,self).add(ctrl)
		if len(self._ctrl):
			self._idx = 0
		
	def reactionTime(self):
		return 0

	def onLeft(self):
		idx = self._idx
		if idx is not None:
			if idx > 0:
				self._idx = idx - 1
			else:
				self._idx = len(self._ctrl) - 1
		self.sibi.speakInFocusAfter(self.reactionTime())
		return True

	def onRight(self):
		idx = self._idx
		if idx is not None:
			if idx < len(self._ctrl) - 1:
				self._idx = idx + 1
			else:
				self._idx = 0
		self.sibi.speakInFocusAfter(self.reactionTime())
		return True

class TabGroup(GroupBase):
	"""
	Group based tab control, so allow operations with elements inside _chooser
	Childred shoud have "idActive" and "Activate" methods
	"""
	def __init__(self, name, sibi, opt = None):
		super(TabGroup,self).__init__(name, sibi, opt);
		self._element_type = "page"
		
	def _getCurrent(self):
		for ctrl in self._ctrl:
			if hasattr(ctrl, "isActive") and getattr(ctrl, "isActive")():
				return ctrl
		self._chooser = True
		return None

	def onLeft(self):
		count = len(self._ctrl)
		if count == 0:
			return True
		for i, ctrl in enumerate(self._ctrl):
			if ctrl.isActive():
				if i > 0:
					self._ctrl[i - 1].activate()
				else:
					self._ctrl[count-1].activate()
				self.speakInFocusAfter()
				return True
		self._ctrl[0].activate()
		self.speakInFocusAfter()
		return True

	def onRight(self):
		count = len(self._ctrl)
		if count == 0:
			return True
		for i, ctrl in enumerate(self._ctrl):
			if ctrl.isActive():
				if i < count - 1:
					self._ctrl[i + 1].activate()
				else:
					self._ctrl[0].activate()
				self.speakInFocusAfter()
				return True
		self._ctrl[0].activate()
		self.speakInFocusAfter()
		return True
		
class _DialogControl(Control):

	def __init__(self, name, sibi, dialog):
		super(_DialogControl,self).__init__( name, sibi, set() )
		self.dialog = dialog
		self.type = 'Dialog'

	def focusSet(self):
		self.sibi.focusChangedTo(self, False) # focusLost is normally translated with some clicks... not good when entering dialogs
		
	def onEnter(self):
		for ctrl in self.dialog.ctrl:
			if 'OK' in ctrl.opt:
				return ctrl.onEnter()
		return True

	def onEscape(self):
		for ctrl in self.dialog.ctrl:
			if 'CANCEL' in ctrl.opt:
				return ctrl.onEnter()
		return True

class Dialog(Container):
	"""
	Dialog inside the interface,
	isFocusable() return if it is really opened, focus navigation is cyclic
	"""
	def __init__(self, name, sibi, sense_pt, inactive_color, active_color, opt):
		"""
		@param name: the name of the dialog
		@type name: string
		@param sibi: SIBI
		@type sibi: SIBI
		@param sense_pt: the coordinates of the sence point, which determines the dialog activation
		@type sence_pt: Pt
		@param inactive_color, active_color: background and dialog colors, if inactive_color is 0, active_color identifz EXACT color expected
		@type inactive_color, active_color: BGR 
		@param opt: a set of options
		@type opt: set
		"""
		super(Dialog,self).__init__( name, sibi, opt)
		self.type = "Dialog"
		self.add( _DialogControl(name, sibi, self) )
		self.box = Box(sibi.hwnd, sibi.xScale(sense_pt), sibi.yScale(sense_pt))
		if inactive_color:
			self.colors = (c_int*2)(inactive_color, active_color)
			self.color = None
		else:
			self.colors = None
			self.color = active_color
		
	def isFocusable(self):
		global sibiac
		if not sibiac:
			return False
		if self.colors:
			i = sibiac.NearestColor(self.box.hwnd, c_int(self.box.left), c_int(self.box.top + self.box.vshift), c_int(2), byref(self.colors))
			return  i == 1
		color = sibiac.PixelColor(self.box.hwnd, c_int(self.box.left), c_int(self.box.top + self.box.vshift))
		return color == self.color
				
	def focusNext(self):
		if not super(Dialog,self).focusNext():
			return self.focusFirst()
		return True
		
	def focusPrevious(self):
		if not super(Dialog,self).focusPrevious():
			return self.focusLast()
		return True

	
def isWinWithinScreen(hwnd):
	"""
	Check specified window is completely inside its parents, including screen
	"""
	r = RECT()
	user32.GetClientRect(hwnd, byref(r))
	r = ClientRectToScreen(hwnd, r)
	while hwnd:
		istop = isWinNotChild(hwnd)
		if istop:
			hprt = user32.GetDesktopWindow()
		else:
			hprt = user32.GetParent(hwnd)
		if not hprt:
			break
		pr = RECT()
		user32.GetClientRect(hprt, byref(pr))
		pr = ClientRectToScreen(hprt, pr)
		if r.left < pr.left or r.right > pr.right or r.top < pr.top or r.bottom > pr.bottom:
			# log.error("Parent: (%d %d)-(%d %d), Window: (%d %d)-(%d %d)" % (pr.left, pr.top, pr.right, pr.bottom, r.left, r.top, r.right, r.bottom))
			return False
		if istop:
			break
		r = pr
		hwnd = hprt
	return True		
		
class SIBI(Container):
	"""
	Whole interface container to calculate geometry and define speak callback
	"""

	def __init__(self, hwnd, width, height, opt = None):
		super(SIBI,self).__init__("SIBI", None, () )
		self.sibi = self # not really useful
		self.hwnd = hwnd
		# precalculate the geometry (assume never change)
		self.width = OneIfZero(width)
		self.height = OneIfZero(height)
		r = RECT()
		user32.GetClientRect(self.hwnd, byref(r))
		self.win_width = OneIfZero(r.right)
		self.win_height = OneIfZero(r.bottom)
		self.mag_factor = float(self.win_width) / float(self.width)
		mag_height = self.mag_factor * self.height
		self.y_shift = self.win_height - mag_height
		self.y_scale = float(self.win_height) / float(self.height)
		#
		self.info_queued = 0
		self.focus_queued = 0
		self.infocus_queued = 0
		self.focus_ctrl = None
		self.nvda_class = None
		self.panel_defined = False # If panel is defined, dialog should not react on focus
		self.had_focus = False # Set to true the first time we get focus
	
		if opt is None:
			opt = set()
		self.block_arrows = ("block_arrows" in opt)
		self.textout = ("textout" in opt) # use NVDA collected text globally
		
		self.expected_focus_return = False # When opening in-plugin dialogs
		
		if not isWinWithinScreen(self.hwnd):
			fitWindowIntoDesktop(self.hwnd)
			time.sleep(0.5)
			if not isWinWithinScreen(self.hwnd):
				maximizeWindow(self.hwnd) # that can help in case REAPER options not set correctly
				time.sleep(0.5)
				if not isWinWithinScreen(self.hwnd):
					self.add( Label("Warning: the interface is partially hidden. Sibiac can not work as desired.", self, Pt(0, 0), Pt(1, 1), None) )
	
	def MXY(self, pt, getShift = None):
		xy = self.ptScale(pt)
		return MXY(self.hwnd, xy.x, xy.y, getShift)
		
	def Box(self, lt, rb = None, getShift = None):
		lt = self.ptScale(lt)
		if rb is None:
			rb = lt
		else:
			rb = self.ptScale(rb)
		return Box(self.hwnd, lt.x, lt.y, rb.x, rb.y, getShift)

	def TextBox(self, lt, rb, getShift = None):
		lt = self.ptScale(lt)
		rb = self.ptScale(rb)
		return TextBox(self.hwnd, lt.x, lt.y, rb.x, rb.y, getShift)

	def TextOutBox(self, lt, rb, getShift = None):
		lt = self.ptScale(lt)
		rb = self.ptScale(rb)
		return TextOutBox(self.hwnd, lt.x, lt.y, rb.x, rb.y, getShift)
	
	def expectFocusReturn(self):
		""" inform Sibiac the focus will temporary leave, cancel speach on return """
		self.expected_focus_return = True
		
	def focusReturn(self):
		""" Return True in case focus return was expected """
		ret = self.expected_focus_return
		self.expected_focus_return = False
		return ret

	def focusChangedTo(self, ctrl, inform_last = True):
		"""
		Called by controls when they got focus
		"""
		if inform_last and self.focus_ctrl and self.focus_ctrl != ctrl:
			self.focus_ctrl.focusLost()
		self.focus_ctrl = ctrl
	
	def speakAfter(self, delay):
		"""
		Shoud inform user about status, possible with delay (if > 0)
		
		General code has no NVDA references (for clarity), so define it later
		"""
		pass

	def speakFocusAfter(self, delay):
		"""
		Ask NVDA to report current focus (with control type and name), possible with delay (if > 0)
		
		General code has no NVDA references (for clarity), so define it later
		"""
		pass

	def speakInFocusAfter(self, delay):
		"""
		Ask NVDA to report partial focus (with name), possible with delay (if > 0)
		
		General code has no NVDA references (for clarity), so define it later
		"""
		pass

	def getHWnd(self):
		return self.hwnd

	def xScale(self, pt_or_x):
		if isinstance(pt_or_x, Pt):
			return int(pt_or_x.x*self.mag_factor)
		return int(pt_or_x*self.mag_factor)
	
	def yScale(self, pt_or_y):
		if isinstance(pt_or_y, Pt):
			y = pt_or_y.y
			if pt_or_y.scale_mode == MOVE:
				return int(y*self.mag_factor + self.y_shift/pt_or_y.scale_factor)
			if pt_or_y.scale_mode == PROPORTIONAL:
				return int(y*self.y_scale) 
		else:
			y = pt_or_y
		return int(y*self.mag_factor)

	def ptScale(self, pt, shift = None):
		xy = XY(self.xScale(pt), self.yScale(pt))
		if shift is not None:
			xy.x += shift[0]
			xy.y += shift[1]
		return xy
		
	def focusSet(self):
		"""
		Inform the control it got the focus
		"""
		self.had_focus = True
		super(SIBI,self).focusSet()
		
	def focusNext(self):
		"""
		Orginize global cyclic focus
		"""
		if not super(SIBI, self).focusNext():
			return self.focusFirst()
		return True

	def focusPrevious(self):
		"""
		Orginize global cyclic focus
		"""
		if not super(SIBI, self).focusPrevious():
			return self.focusLast()
		return True
		
class VList(Control):
	"""
	Vertical list box
	"""
	def __init__(self, name, sibi, first_left_top, first_right_bottom, second_left_top, last_left_bottom,
				 inactive_color, active_color, opt):
		"""
		@param name: the name of the list
		@type name: string
		@param sibi: SIBI
		@type sibi: SIBI
		@param first_left_top, first_right_bottom, second_left_top, last_left_bottom: coordinates
		@type first_left_top, first_right_bottom, second_left_top, last_left_bottom: Pt
		@param inactive_color, active_color: background colors
		@type inactive_color, active_color: BGR or tuple of BGRs
		@param opt: set of options, "click_to_focus", "click_on_enter", "pass_enter", "multi_one", "dblclick_on_enter"
		@type opt: set
		"""
		super(VList,self).__init__( name, sibi, opt)
		self.type = "List"
		self.vlist = VListBox( sibi.hwnd, sibi.xScale(first_left_top), sibi.yScale(first_left_top),
						sibi.xScale(first_right_bottom), sibi.yScale(first_right_bottom),
						sibi.yScale(second_left_top), sibi.yScale(last_left_bottom),
						inactive_color, active_color )


	def vShift(self, vshift = 0):
		self.vlist.vShift(vshift)

						
	# Control Interface
	
	def getTextInfo(self):
		n, shift = self.vlist.selection()
		if n < 0:
			return (self.type, self.name, "No selection")
		return (self.type, self.name, self.vlist.getText(n, shift))

	def focusSet(self):
		super(VList,self).focusSet()
		n, shift = self.vlist.selection()
		if n < 0:
			return
		if "click_to_focus" in self.opt or "page_keys" in self.opt or "homeend_keys" in self.opt:
			self.vlist.leftClick(n, shift)
			time.sleep(0.05)
		
	def onUp(self):
		n, shift = self.vlist.selection()
		if n < 0:
			n = 0
			self.vlist.leftClick(n, 0)
			self.sibi.speakAfter(self.reactionTime())
			return True
		if "click_to_focus" in self.opt:
			self.sibi.speakAfter(self.reactionTime())
			return False
		if n > 0:
			if "multi_one" in self.opt:
				self.vlist.leftClick(n, shift) # to deselect
			n -= 1
			self.vlist.leftClick(n, shift)
		time.sleep(0.05)
		self.sibi.speakAfter(self.reactionTime())
		return True

	def onDown(self):
		n, shift = self.vlist.selection()
		if n < 0:
			n = 0
			self.vlist.leftClick(n, 0)
			self.sibi.speakAfter(self.reactionTime())
			return True
		if "click_to_focus" in self.opt:
			self.sibi.speakAfter(self.reactionTime())
			return False
		if n + 1 < self.vlist.n_items:
			if "multi_one" in self.opt:
				self.vlist.leftClick(n, shift) # to deselect
			n += 1
			self.vlist.leftClick(n, shift)
		time.sleep(0.05)
		self.sibi.speakAfter(self.reactionTime())
		return True

	def onEnter(self):
		if not "click_on_enter" and not "dblclick_on_enter" in self.opt:
			self.sibi.speakAfter(self.reactionTime())		
			return not "pass_enter" in self.opt
		n, shift = self.vlist.selection()
		if n >= 0:
			self.vlist.leftClick(n, shift)
			if "dblclick_on_enter" in self.opt:
				time.sleep(0.05)
				self.vlist.leftClick(n, shift)
		self.sibi.speakAfter(self.reactionTime())
		return True

class YRange(Control):
	"""
	A box with a vertical range control in the middle, the range is set by mouse scroll at the left or right border.
	RangeToText is responsibe for value convertation
	"""
	def __init__(self, name, sibi, left_top, right_bottom, bg, fg, opt):
		super(YRange,self).__init__( name, sibi, opt)
		nw_x = sibi.xScale(left_top)
		nw_y = sibi.yScale(left_top)
		se_x = sibi.xScale(right_bottom)
		se_y = sibi.yScale(right_bottom)
		self.lbox = Box( sibi.hwnd, nw_x, nw_y + (se_y - nw_y)//2 )
		self.rbox = Box( sibi.hwnd, se_x, nw_y + (se_y - nw_y)//2 )
		self.x = nw_x + (se_x - nw_x)//2
		self.y0 = nw_y
		self.y1 = se_y
		self.h  = se_y - nw_y
		self.bg = Color2Array(bg)
		self.fg = Color2Array(fg)
		self.type = "Range"
	
	def vShift(self, vshift = 0):
		self.lbox.vShift(vshift)
		self.rbox.vShift(vshift)
		
	def getCurrentRange(self):
		global sibiac
		if not sibiac:
			return (None, None)
		top = c_int(0)
		bottom = c_int(0)
		if not sibiac or sibiac.FindRangeInY(self.lbox.hwnd, c_int(self.x), c_int(self.y0 + self.lbox.vshift), c_int(self.y1 + self.lbox.vshift),
										c_int(len(self.bg)), byref(self.bg), c_int(len(self.fg)),byref(self.fg),
										byref(top), byref(bottom)) == 0:
			return (100, 100)
		return ( (top.value - self.y0 - self.lbox.vshift)*100//self.h, (bottom.value - self.y0 - self.lbox.vshift)*100//self.h )

	def RangeToText(self, vfrom, vto):
		return str(100-vto)+"% - "+str(100-vfrom)+"%"
	
	def getTextInfo(self):
		vfrom, vto = self.getCurrentRange()
		if vfrom is None:
			text = "Undefined"
		else:
			text = self.RangeToText(vfrom, vto)
		return ( self.type, self.name, text )
	
	def onUp(self):
		self.rbox.moveTo()
		MouseScroll(120)
		self.sibi.speakAfter(self.reactionTime())
		return True

	def onDown(self):
		self.rbox.moveTo()
		MouseScroll(-120)
		self.sibi.speakAfter(self.reactionTime())
		return True

	def onLeft(self):
		self.lbox.moveTo()
		MouseScroll(-120)
		self.sibi.speakAfter(self.reactionTime())
		return True

	def onRight(self):
		self.lbox.moveTo()
		MouseScroll(120)
		self.sibi.speakAfter(self.reactionTime())
		return True

class YBar(YRange):
	"""
	A box with a vertical value control in the middle, the range is set by mouse scroll.
	RangeToText is responsibe for value convertation
	"""
	def __init__(self, name, sibi, left_top, right_bottom, bg, fg, opt):
		super(YBar,self).__init__( name, sibi, left_top, right_bottom, bg, fg, opt)
		self.type = "Bar"

	def RangeToText(self, vfrom, vto):
		return str(100-vfrom)+"%"
		
class Label(Control):
	"""
	Label, support "dynamic_text", "silent_action", "textout", "notextout"
	"""
	def __init__(self, name, sibi, left_top, right_bottom, opt = None, getShift = None):
		"""
		@param name: the name of the element
		@type name: string
		@param sibi: SIBI
		@type sibi: SIBI
		@param left_top, right_bottom: coordinates
		@type left_top, right_bottom: Pt
		@param opt: set of options, "dynamic_text"
		@type opt: set
		"""
		super(Label,self).__init__( name, sibi, opt)
		if right_bottom is None:
			right_bottom = left_top
		self.box = TextBox( sibi.hwnd, sibi.xScale(left_top), sibi.yScale(left_top), sibi.xScale(right_bottom), sibi.yScale(right_bottom), getShift )

	def vShift(self, vshift = 0):
		self.box.vShift(vshift)

	def getStateText(self):
		return ""

	def getTextInfo(self):
		state = self.getStateText()
		if state != "":
			state = ", " + state
		if not "dynamic_text" in self.opt:
			return ( self.type, '', self.name + state )
		if ("textout" in self.opt) or (self.sibi.textout and "notextout" not in self.opt):
			#text = "Textout :" + self.box.getTextOut()
			text = self.box.getTextOut()
		else:
			text = self.box.getText()
		if text == "":
			text = "Empty"
		return ( self.type, self.name, text + state )

	def getNameInGroup(self):
		"""
		Should return Name on (Name, Text) tuple.
		It used when the control is in a group, to annouce when the group select it
		"""
		type, name, text = self.getTextInfo()
		if name == "":
			return text
		return (name, text)

	# Control Interface
	
	def onUp(self):
		return True

	def onDown(self):
		return True

	def onEnter(self):
		if "speak_on_enter" in self.opt:
			self.sibi.speakAfter(0)
		return True

class ScrollLabel(Label):
	"""
	Label, which change the value when mouse is scrolled
	"""
	def __init__(self, name, sibi, left_top, right_bottom, opt):
		super(ScrollLabel,self).__init__( name, sibi, left_top, right_bottom, opt)
		self.opt.add("dynamic_text")
	
	def onUp(self):
		self.box.moveTo()
		MouseScroll(120)
		self.sibi.speakAfter(self.reactionTime())
		return True

	def onDown(self):
		self.box.moveTo()
		MouseScroll(-120)
		self.sibi.speakAfter(self.reactionTime())
		return True

	def onLeft(self):
		self.box.moveTo()
		MouseScroll(-30)
		self.sibi.speakAfter(self.reactionTime())
		return True

	def onRight(self):
		self.box.moveTo()
		MouseScroll(30)
		self.sibi.speakAfter(self.reactionTime())
		return True
	
class PopupLabel(Control):
	"""
	Popup label with dynamic width
	"""
	def __init__(self, name, sibi, center_top, center_bottom, bg, fg, opt):
		super(PopupLabel,self).__init__( name, sibi, opt)
		self.box = PopupBox( sibi.hwnd, sibi.xScale(center_top), sibi.yScale(center_top), sibi.yScale(center_bottom), bg, fg)
		self.text_hash = {}
	
	def vShift(self, vshift = 0):
		self.box.vShift(vshift)
		
	def isVisible(self):
		return self.box.isVisible()

	def getTextInfo(self):
		if not self.isVisible():
			return ( self.type, self.name, "Not visible" )
		crc32 = self.box.getCRC()
		if crc32 in self.text_hash:
			text = self.text_hash[crc32]
		else:
			text = self.box.getText()
			self.text_hash[crc32] = text
		if text == "":
			text = "Empty"
		return ( self.type, self.name, text )

class SpinLabel(Label):
	"""
	Dynamic label + next/previous points, support "click_on_enter"
	"""
	def __init__(self, name, sibi, left_top, right_bottom, prev_pt, next_pt, opt = None, getShift = None):
		super(SpinLabel,self).__init__( name, sibi, left_top, right_bottom, opt, getShift)
		self.type = 'Choose'
		if "dynamic_text" not in self.opt:
			self.opt.add('dynamic_text')
		self.prev_btn = sibi.MXY(prev_pt, getShift)
		self.next_btn = sibi.MXY(next_pt, getShift)
							   
	def onDown(self):
		self.next_btn.leftClick()
		return self.speakAfter()

	def onUp(self):
		self.prev_btn.leftClick()
		return self.speakAfter()
	
	def onEnter(self):
		if "click_on_enter" not in self.opt:
			return True
		self.box.leftClick()
		return self.speakFocusAfter()
		
class Combo(Label):
	"""
	Combo box with drop down list
	"""

	def _getListTextInfo(self):
		n, shift = self.vlist.selection()
		if n < 0:
			return ("Combo list", self.name, "No selection")
		return ("Combo list", self.name, 'Select ' + self.vlist.getText(n, shift))

	def _isOpen(self):
		global sibiac
		if not sibiac:
			return False
		i = sibiac.FindColorInY(self.box.hwnd, c_int(self.list_left_x), c_int(self.list_left_y + self.box.vshift),
								  c_int(len(self.list_bg_color)),byref(self.list_bg_color),
								  c_int(len(self.list_not_bg_color)),byref(self.list_not_bg_color))
		i = int(i)
		if i < 1:
			return False
		if self.isDynamic:
			self.vlist.setHeight(i)
		return True
	
	def __init__(self, name, sibi, left_top, right_bottom, 
					list_left_top, list_left_bottom, list_bg_color, list_not_bg_color,
					first_left_top, first_right_bottom, second_left_top,
					inactive_color, active_color, opt):
		"""
		@param name: the name of the element
		@type name: string
		@param sibi: SIBI
		@type sibi: SIBI
		@param left_top, right_bottom: coordinates for closed combo text
		@type left_top, right_bottom: Pt
		@param list_left_top, list_left_bottom: coordinates to detect the list, if list_left_bottom y is 0, it is dynamic
		@type list_left_top, list_left_bottom: Pt
		@parm list_bg_color, list_not_bg_color: colors on list_left to detect it is opened
		@type list_bg_color, list_not_bg_color: BGR or tuple of BGRs
		@param first_left_top, first_right_bottom, second_left_top, last_left_bottom: coordinates
		@type first_left_top, first_right_bottom, second_left_top, last_left_bottom: Pt
		@param inactive_color, active_color: background colors
		@type inactive_color, active_color: BGR or tuple of BGRs
		@param opt: set of options
		@type opt: set
		"""
		super(Combo,self).__init__( name, sibi, left_top, right_bottom, opt)
		self.type = 'Combo'
		self.opt.add('dynamic_text')

		self.list_left_x = sibi.xScale(list_left_top)
		self.list_left_y = sibi.yScale(list_left_top)
		if list_left_bottom.y == 0:
			self.isDynamic = True
			last_bottom = 0
		else:
			self.isDynamic = False
			last_bottom = sibi.yScale(list_left_bottom)
		self.list_bg_color = Color2Array(list_bg_color)
		self.list_not_bg_color = Color2Array(list_not_bg_color)

		self.vlist = VListBox( sibi.hwnd, sibi.xScale(first_left_top), sibi.yScale(first_left_top),
							   sibi.xScale(first_right_bottom), sibi.yScale(first_right_bottom),
							   sibi.yScale(second_left_top), last_bottom,
							   inactive_color, active_color )
							   
	def vShift(self, vshift = 0):
		super(Combo,self).vShift(vshift)
		self.vlist.vShift(vshift)
		
	def getTextInfo(self):
		if self._isOpen():
			return self._getListTextInfo()
		return super(Combo,self).getTextInfo()

	def onDown(self):
		if not self._isOpen():
			self.box.leftClick()
			time.sleep(0.05)
		if not self._isOpen():
			return True

		n, shift = self.vlist.selection()
		if n < 0:
			n = 0
		elif n + 1 < self.vlist.n_items:
			n += 1
		self.vlist.moveTo(n, 0)
		self.sibi.speakAfter(self.reactionTime())
		return True			

	def onUp(self):
		if not self._isOpen():
			return True

		n, shift = self.vlist.selection()
		if n < 0:
			n = 0
		elif n > 0:
			n -= 1
		self.vlist.moveTo(n, 0)
		self.sibi.speakAfter(self.reactionTime())
		return True			
		
	def onEnter(self):
		if not self._isOpen():
			return True
		n, shift = self.vlist.selection()
		if n >= 0:
			self.vlist.leftClick(n, 0)
		else:
			self.box.leftClick()
		self.sibi.speakAfter(self.reactionTime())
		return True
		
	def onEscape(self):
		if self._isOpen():
			self.box.leftClick()
			time.sleep(0.05) # give time to close
			self.sibi.speakAfter(self.reactionTime())
			return True
		return False
		
	def focusLost(self):
		if self._isOpen():
			self.box.leftClick()
			time.sleep(0.05) # give time to close

class PopupMenuButton(Label):
	"""
	A button which calls popup (not acesssible) menu. The button can have dynamic text.
	Menu should hilight current element and can have "selection" (selected_color not None).
	"""

	def _getMenuTextInfo(self):
		if self.in_submenu:
			menu = self.submenu
		else:
			menu = self.menu
		n, shift, is_selected = menu.current()
		if n < 0:
			return (menu.type, menu.name, "No selection")
		selected = "";
		if is_selected:
			selected = ", selected"
		if self.has_submenu and not self.in_submenu:
			selected += ", submenu"
		return (menu.type, menu.name, menu.getText(n, shift)+selected)

	def _isSubmenuOpen(self):
		global sibiac
		if self.submenu_x and sibiac:
			# submenu configured, check it is open, use Y center of the current item for detection
			n, shift, is_selected = self.menu.current()
			if n >= 0:
				y = self.menu.top + self.menu.item_h*n + self.menu.box_h//2
				i = sibiac.FindColorInY(self.box.hwnd, c_int(self.submenu_x), c_int(y),
								  c_int(len(self.menu_bg_color)),byref(self.menu_bg_color),
								  c_int(len(self.menu_not_bg_color)),byref(self.menu_not_bg_color))
				i = int(i)
				if i > 0:
					j = sibiac.FindColorInYUp(self.box.hwnd, c_int(self.submenu_x), c_int(y),
								  c_int(len(self.menu_bg_color)),byref(self.menu_bg_color),
								  c_int(len(self.menu_not_bg_color)),byref(self.menu_not_bg_color))
					j = int(j)
					if j > 0:
						k = sibiac.FindColorInX(self.box.hwnd, c_int(self.submenu_x), c_int(y + 2 - j),
								  c_int(len(self.menu_bg_color)),byref(self.menu_bg_color),
								  c_int(len(self.menu_not_bg_color)),byref(self.menu_not_bg_color))
						k = int(k)
						if k > 0 and k > self.submenu.left_border+self.submenu.right_border+30 and i+j-1 >= self.submenu.top_border+self.submenu.bottom_border+self.submenu.box_h:
							# detected
							self.submenu.updateGeometry(self.submenu_x, y + 1 - j, self.submenu_x + k - 1, y + i - 1)
							self.has_submenu = True
							return True
		self.in_submenu = False
		self.has_submenu = False
		return False
		
	def _isOpen(self):
		global sibiac
		if not sibiac:
			return False
		i = sibiac.FindColorInY(self.box.hwnd, c_int(self.menu_left_x), c_int(self.menu_left_y + self.box.vshift),
								  c_int(len(self.menu_bg_color)),byref(self.menu_bg_color),
								  c_int(len(self.menu_not_bg_color)),byref(self.menu_not_bg_color))
		i = int(i)
		if i < 1:
			self.in_submenu = False
			self.has_submenu = False
			return False
		if self.isDynamic:
			self.menu.setHeight(i)
		self._isSubmenuOpen()
		return True
	
	def __init__(self, name, sibi, left_top, right_bottom, 
					menu_left_top, menu_left_bottom, menu_bg_color, menu_not_bg_color,
					first_left_top, first_right_bottom, second_left_top,
					inactive_color, active_color, selected_color, opt):
		"""
		@param name: the name of the element
		@type name: string
		@param sibi: SIBI
		@type sibi: SIBI
		@param left_top, right_bottom: coordinates for the button
		@type left_top, right_bottom: Pt
		@param menu_left_top, menu_left_bottom: coordinates to detect the list, if list_left_bottom y is 0, it is dynamic
		@type menu_left_top, menu_left_bottom: Pt
		@parm menu_bg_color, menu_not_bg_color: colors on menu_left to detect it is opened
		@type menu_bg_color, menu_not_bg_color: BGR or tuple of BGRs
		@param first_left_top, first_right_bottom, second_left_top, last_left_bottom: coordinates
		@type first_left_top, first_right_bottom, second_left_top, last_left_bottom: Pt
		@param inactive_color, active_color, selected_color: background colors. selected_color can be None, in this case menu has no permanently selected items 
		@type inactive_color, active_color, selecter_color: BGR or tuple of BGRs
		@param opt: set of options , "slow_reaction", "dynamic_text", "skip_spacers", "slow_open", "strip_prefix"
		@type opt: set
		"""
		super(PopupMenuButton,self).__init__( name, sibi, left_top, right_bottom, opt)
		self.type = 'Open menu'

		self.menu_left_x = sibi.xScale(menu_left_top)
		self.menu_left_y = sibi.yScale(menu_left_top)
		if menu_left_bottom.y == 0:
			self.isDynamic = True
			last_bottom = 0
		else:
			self.isDynamic = False
			last_bottom = sibi.yScale(menu_left_bottom)
		self.menu_bg_color = Color2Array(menu_bg_color)
		self.menu_not_bg_color = Color2Array(menu_not_bg_color)

		nw_x = sibi.xScale(first_left_top)
		nw_y = sibi.yScale(first_left_top)
		se_x = sibi.xScale(first_right_bottom)
		se_y = last_bottom	
		self.menu = PopupMenu( sibi.hwnd, nw_x, nw_y, se_x, se_y, 0, 0, 0, 0, sibi.yScale(first_right_bottom) - nw_y + 1, sibi.yScale(second_left_top) - nw_y,
							   inactive_color, active_color, selected_color )
		self.menu.name = self.name

		self.submenu = PopupMenu( sibi.hwnd, nw_x, nw_y, se_x, se_y, 0, 0, 0, 0, sibi.yScale(first_right_bottom) - nw_y + 1, sibi.yScale(second_left_top) - nw_y,
							   inactive_color, active_color, selected_color )
		self.submenu.type = ""
		self.submenu.name = self.name
		self.submenu_x = 0 # means not defined
		self.in_submenu = False
		self.has_submenu = False # detected in _isOpen

	def defineSubmenu(self, pt_x, left_border, right_border, top_border, bottom_border, box_h, item_h):
		""" That is just configuration, geometry will be updated when submenu is detected """
		self.submenu_x = self.sibi.xScale(pt_x)
		self.submenu.left_border = left_border
		self.submenu.right_border = right_border
		self.submenu.top_border = top_border
		self.submenu.bottom_border = bottom_border
		self.submenu.box_h = box_h
		self.submenu.item_h = item_h

	def vShift(self, vshift = 0):
		super(PopupMenuButton,self).vShift(vshift)
		self.menu.vShift(vshift)
		
	def getTextInfo(self):
		if self._isOpen():
			type, name, txt = self._getMenuTextInfo()
			if txt and "strip_prefix" in self.opt:
				while len(txt) > 0 and not txt[0].isalnum():
					txt = txt[1:]
			return (type, name, txt)
		return super(PopupMenuButton,self).getTextInfo()

	def onDown(self):
		if not self._isOpen():
			return True

		if self.in_submenu:
			menu = self.submenu
		else:
			menu = self.menu
		n, shift, is_selected = menu.current()
		current_n = n
		if n < 0:
			n = 0
			current_n = 0
		elif n + 1 < menu.n_items:
			n += 1
		if "skip_spacers" in self.opt:
			while n < menu.n_items:
				if menu.getText(n, shift):
					break
				n += 1
			if n == menu.n_items:
				n = current_n
			
		menu.moveTo(n, 0)
		self.sibi.speakAfter(self.reactionTime())
		return True			

	def onUp(self):
		if not self._isOpen():
			return True
		if self.in_submenu:
			menu = self.submenu
		else:
			menu = self.menu
		n, shift, is_selected = menu.current()
		current_n = n
		if n < 0:
			n = 0
			current_n = 0
		elif n > 0:
			n -= 1
		if "skip_spacers" in self.opt:
			while n >= 0:
				if menu.getText(n, shift):
					break
				n -= 1
			if n < 0:
				n = current_n
		menu.moveTo(n, 0)
		self.sibi.speakAfter(self.reactionTime())
		return True			
	
	def onLeft(self):
		if not self._isOpen() or not self.in_submenu:
			return True
		self.in_submenu = False
		self.sibi.speakAfter(self.reactionTime())
		return True			

	def onRight(self):
		if not self._isOpen() or self.in_submenu:
			return True
		self.in_submenu = True
		self.sibi.speakAfter(self.reactionTime())
		return True
		
	def onEnter(self):
		if not self._isOpen():
			if "slow_open" in self.opt:
				self.box.moveTo()
				time.sleep(0.05)
				MouseSlowLeftClick()
			else:
				self.box.leftClick()
			time.sleep(0.05)
			if self.submenu_x and self._isOpen():
				n, shift, is_selected = self.menu.current()
				if n >= 0:
					self.menu.moveTo(n, 0)
			self.sibi.speakFocusAfter(self.reactionTime())
			return True
		if self.in_submenu:
			menu = self.submenu
		else:
			menu = self.menu
		n, shift, is_selected = menu.current()
		if n >= 0:
			menu.moveTo(n, 0)
			time.sleep(0.05)
			menu.leftClick(n, 0)
			self.sibi.speakFocusAfter(self.reactionTime())
		return True
		
	def onEscape(self):
		if self._isOpen():
			box = Box(self.box.hwnd, 5, self.box.top)
			box.leftClick()
			time.sleep(0.05) # give time to close
			self.sibi.speakFocusAfter(self.reactionTime())
			return True
		return False
		
	def focusLost(self):
		if self._isOpen():
			box = Box(self.box.hwnd, 5, self.box.top)
			box.leftClick()
			time.sleep(0.05) # give time to close

class Clickable(Label):
	"""
	Clickable Label, support "dynamic_text", "silent_action", "move_away", "slow_reaction", "focus_will_move"
	"""
	# Control Interface
	
	def onEnter(self):
		box = self.box.getBox()
		if "focus_will_move" in self.opt:
			self.expectFocusReturn()
		box.leftClick()
		if not "silent_action" in self.opt:
			self.sibi.speakAfter(self.reactionTime())
		if "move_away" in self.opt:
			MXY(self.sibi.hwnd, box.left - 10, box.top - 10).moveTo()
		return True

	def onHelp(self, prefix = "", suffix = ""):
		self.speak(prefix + "Press enter to click" + suffix)
		return True
	
class PushBtn(Clickable):
	"""
	Simple push button, support disabled state and dynamic text
	"""
	def __init__(self, name, sibi, left_top, right_bottom, disabled_colors, enabled_colors, opt = None, getShift = None):
		"""
		@param name: the name of the list
		@type name: string
		@param sibi: SIBI
		@type sibi: SIBI
		@param left_top, right_bottom: coordinates
		@type left_top, right_bottom: Pt
		@param disabled_color, enabled_color: when (both!) not None, detect disabled state
		@type disabled_color, enabled_color: BGR
		@param opt: set of options, "dynamic_text", "say_enabled"
		@type opt: set
		"""
		super(PushBtn,self).__init__( name, sibi, left_top, right_bottom, opt, getShift )
		self.disabled_colors = Color2Tuple(disabled_colors)
		self.enabled_colors = Color2Tuple(enabled_colors)

	def isEnabled(self):
		if self.disabled_colors is None or self.enabled_colors is None:
			return True
		box = self.box.getBox()
		i = FindNearestColor(box.hwnd, box.left, box.top + box.vshift, self.disabled_colors + self.enabled_colors)
		return i >= len(self.disabled_colors)

	def getStateText(self):
		if not self.isEnabled():
			return ", disabled"
		if "say_enabled" in self.opt:
			return ", enabled"
		return ""

class OpenBtn(PushBtn):
	"""
	Push button which open dialog on clicking, support disabled state and dynamic text
	"""
	def onEnter(self):
		self.box.getBox().leftClick()
		self.sibi.speakFocusAfter(self.reactionTime())
		return True

class CloseBtn(Clickable):
	"""
	Push button which close dialog on clicking
	"""
	def onEnter(self):
		self.box.getBox().leftClick()
		self.sibi.speakFocusAfter(self.reactionTime())
		return True

class SwitchBtn(Clickable):
	"""
	Push On/Off button, support disabled state and dynamic text
	"""
	def __init__(self, name, sibi, left_top, right_bottom, disabled_color, off_color, on_color, opt, getShift = None):
		"""
		@param name: the name of the list
		@type name: string
		@param sibi: SIBI
		@type sibi: SIBI
		@param left_top, right_bottom: coordinates
		@type left_top, right_bottom: Pt
		@param disabled_color, enabled_color: when (both!) not None, detect disabled state
		@type disabled_color, enabled_color: BGR
		@param opt: set of options, "dynamic_text"
		@type opt: set
		"""
		super(SwitchBtn,self).__init__( name, sibi, left_top, right_bottom, opt, getShift )
		self.disabled_color = disabled_color
		self.off_color = off_color
		self.on_color = on_color
		if self.off_color is None:
			self.off_color = 0x000000
		if self.on_color is None:
			self.on_color = 0xffffff
		self.text_hash = {}
		self.off_text = "Off"
		self.on_text = "On"

	def getStateText(self):
		global sibiac
		if not sibiac:
			return "Disabled"
		if self.disabled_color is None:
			colors = (c_int*2)(self.off_color, self.on_color)
			n = 2
		else:
			colors = (c_int*3)(self.off_color, self.on_color, self.disabled_color)
			n = 3
		i = sibiac.NearestColor(self.box.hwnd, c_int(self.box.left), c_int(self.box.top + self.box.vshift), c_int(n), byref(colors))
		if i == 1:
			return self.on_text
		if i == 0:
			return self.off_text
		return "Disabled"

class CheckBtn(Label):
	"""
	Check or radio button, support disabled state and dynamic text
	"""
	def __init__(self, name, sibi, left_top, right_bottom, check_pt, disabled_color, off_color, on_color, opt, getShift = None):
		super(CheckBtn,self).__init__( name, sibi, left_top, right_bottom, opt, getShift )
		if check_pt is None:
			check_pt = left_top
		self.check_box = Box(sibi.hwnd, sibi.xScale(check_pt), sibi.yScale(check_pt), None, None, getShift)
		self.disabled_color = Color2Tuple(disabled_color)
		self.off_color = Color2Tuple(off_color, 0x000000)
		self.on_color = Color2Tuple(on_color, 0xffffff)
		self.off_text = "Off"
		self.on_text = "On"

	def getStateText(self):
		disabled = self.disabled_color
		if disabled is None:
			disabled = ()
		colors = disabled + self.off_color + self.on_color
		box = self.check_box.getBox()
		i = FindNearestColor(box.hwnd, box.left, box.top, colors)
		if i < len(disabled):
			return "Disabled"
		if i < len(disabled) + len(self.off_color):
			return self.off_text
		return self.on_text

	def onEnter(self):
		self.check_box.getBox().leftClick()
		if not "silent_action" in self.opt:
			self.sibi.speakAfter(self.reactionTime())
		return True

class FixedTab(Container):
	
	def __init__(self, name, sibi, left_top, right_bottom, active_point, off_color, on_color, opt):
		super(FixedTab,self).__init__( name, sibi, opt )
		self.box = Box( sibi.hwnd, sibi.xScale(left_top), sibi.yScale(left_top), sibi.xScale(right_bottom), sibi.yScale(right_bottom) )
		self.active_x = c_int(sibi.xScale(active_point))
		self.active_y = c_int(sibi.yScale(active_point))
		self.off_colors = Color2Array(off_color)
		self.on_colors = Color2Array(on_color)
		self.colors = Colors2Array(off_color, on_color)

	def isActive(self):
		global sibiac
		if not sibiac:
			return False
		i = sibiac.NearestColor(self.box.hwnd, self.active_x, self.active_y, c_int(len(self.colors)), byref(self.colors))
		return i >= len(self.off_colors)
	
	def isFocusable(self):
		return self.isActive()

	def getTabName(self):
		return self.name

	def activate(self):
		self.box.leftClick()

class _TabControl(Control):

	def __init__(self, name, sibi, container):
		super(_TabControl,self).__init__( name, sibi, set() )
		self.container = container
	
	def getTextInfo(self):
		ctab = self.container._getCurrentTab()
		if ctab is None:
			return (self.container.type, self.name, 'active tab not found')
		return (self.container.type, self.name, self.container.ctrl[ctab].getTabName())
	
	def isActive(self):
		return False # it is not a tab

	def isFocusable(self):
		return not "skip_control" in self.container.opt
		
	def onLeft(self):
		return self.container.previousTab()
	
	def onUp(self):
		return self.onLeft()
		
	def onRight(self):
		return self.container.nextTab()
	
	def onDown(self):
		return self.onRight() 

	def onHelp(self, prefix = "", suffix = ""):
		self.speak(prefix + "Switch between %ss using left and right arrows" % self.container.type + suffix)
		return True

class FixedTabControl(Container):
	"""
	Tab control with predefined set of tabs
	"""

	def __init__(self, name, sibi, opt):
		""" opt - "skip_control" """
		super(FixedTabControl,self).__init__(name, sibi, opt )
		self.type = "Tab"
		self.last_active = 0  # cache last found active tab
		self.add( _TabControl(name, sibi, self) )

	def _getCurrentTab(self):
		""" Detect which tab is currently selected """
		if self.ctrl[self.last_active].isActive():
			return self.last_active
		for ctab, ctrl in enumerate(self.ctrl):
			if ctrl.isActive():
				self.last_active = ctab
				return ctab
		return None

	def _checkCurrent(self):
		""" Check current tab is still valid, switch to TabControl otherwise """
		if self.ctrl_idx > 0 or "skip_control" in self.opt:
			ctab = self._getCurrentTab()
			if ctab is None or ctab != self.ctrl_idx:
				self.ctrl_idx = 0		
		
	def getFocused(self):
		self._checkCurrent()
		return super(FixedTabControl,self).getFocused()
		
	def focusSet(self):
		self._checkCurrent()
		return super(FixedTabControl,self).focusSet()

	def focusLast(self):
		ctab = self._getCurrentTab()
		if ctab is None or not self.ctrl[ctab].focusLast():
			return self.focusFirst()
		self.ctrl_idx = ctab
		return True
				
	def focusNext(self):
		ctab = self._getCurrentTab()
		if ctab is None:
			return False
		if self.ctrl_idx == ctab:
			return self.ctrl[self.ctrl_idx].focusNext()
		self.ctrl_idx = ctab
		return self.ctrl[ctab].focusFirst()
		
	def focusPrevious(self):
		self._checkCurrent()
		if self.ctrl_idx == 0:
			return False
		if self.ctrl[self.ctrl_idx].focusPrevious():
			return True
		if "skip_control" in self.opt:
			return False
		return self.focusFirst()

	def previousTab(self):
		ctab = self._getCurrentTab()
		if ctab is None:
			if not self.ctrl:
				return True
			ctab = 1
		ctab -= 1
		if ctab < 1:
			ctab = len(self.ctrl) - 1
		self.ctrl[ctab].activate()
		if ctab and "skip_control" in self.opt:
			self.ctrl_idx = ctab
			time.sleep(0.05)
			self.focusFirst()
		if not "silent_action" in self.opt:
			self.sibi.speakAfter(self.reactionTime())
		return True
		
	def nextTab(self):
		ctab = self._getCurrentTab()
		if ctab is None:
			if not self.ctrl:
				return True
			ctab = 0
		ctab += 1
		if ctab >= len(self.ctrl):
			ctab = 1
		self.ctrl[ctab].activate()
		if ctab and "skip_control" in self.opt:
			self.ctrl_idx = ctab
			time.sleep(0.05)
			self.focusFirst()
		if not "silent_action" in self.opt:
			self.sibi.speakAfter(self.reactionTime())
		return True

class OptionTable(Container):
	"""
	A table with 2 colums, the first column is the list of options, the second column it
	configuring element for the options. Each row is represented as a separate configuring control with the option name
	"""

	def _setRow(self, row_index = 0):
		if row_index >= self.vlist.n_items:
			return False
		text = self.vlist.getText(row_index)
		if text == "":
			return False
		self.row_ctrl.name = text + ":"
		if self.row_index == row_index:
			self.row_ctrl.focusSet()
			return True
		# To trigger focusLost with old row_ctrl position (when desired), temporarily set sibi focus to us
		Control.focusSet(self)
		self.row_index = row_index
		self.row_ctrl.vShift(self.vlist.item_h*row_index)
		self.row_ctrl.focusSet()
		if row_index > self.last_known_row:
			self.last_known_row = row_index
		return True

	def __init__(self, name, sibi, left_top, right_bottom, second_left_top, last_left_bottom, row_ctrl, opt):
		super(OptionTable,self).__init__( name, sibi, opt)
		self.type = "Table"
		self.vlist = VListBox( sibi.hwnd, sibi.xScale(left_top), sibi.yScale(left_top),
						sibi.xScale(right_bottom), sibi.yScale(right_bottom),
						sibi.yScale(second_left_top), sibi.yScale(last_left_bottom), 0xffffff, 0x000000 )
		self.row_ctrl = row_ctrl
		self.row_index = -1
		self.last_known_row = 0
		self.screll = None
		
	def setScroll(self, nw_pt, se_pt, bg_color, pos_color):
		self.scroll = VScroll(self.vlist.hwnd, self.sibi.xScale(nw_pt), self.sibi.yScale(nw_pt), self.sibi.yScale(se_pt), bg_color, pos_color)

	def getFocused(self):
		if self.row_index < 0:
			return None
		return self.row_ctrl

	def getTextInfo(self):
		"""
		Should return (CtrlType, CtrlName, CtrlText) tuple
		"""
		if self.row_index < 0:
			return (self.type, self.name, 'has no focus')
		return self.row_ctrl.getTextInfo()

	def isFocusable(self):
		text = self.vlist.getText(0)
		if text == "":
			return False
		return True
		
	def focusSet(self):
		"""
		Inform the control it got the focus
		"""
		if self.row_index < 0:
			self.focusNext()
		else:
			self.row_ctrl.focusSet()

	def focusFirst(self):
		"""
		Set focus to the first element
		"""
		return self._setRow(0)

	def focusLast(self):
		"""
		Tricky... we do not know the last row for sure
		"""
		if self.vlist.getText(self.last_known_row) != "":
			# look forward
			while self.last_known_row + 1 < self.vlist.n_items and self.vlist.getText(self.last_known_row + 1):
				self.last_known_row += 1
		else:
			# look backward
			while self.last_known_row > 0 and self.vlist.getText(self.last_known_row - 1):
				self.last_known_row -= 1
		return self._setRow(self.last_known_row)
				
	def _tryScroll(self, up):
		global sibiac
		current_text = self.vlist.getText(self.row_index)
		if not current_text:
			return False
		self.vlist.moveTo(self.row_index) # good position to scroll
		if up:
			delta = 600
		else:
			delta = -600
		sibiac.MouseScroll(c_int(delta))
		time.sleep(0.05)
		if up:
			log.error("Current: "+current_text)
			for idx in xrange(5):
				text = self.vlist.getText(idx)
				if not text:
					break
				log.error(text)
				if text == current_text:
					self.row_index = idx # _setTow sould be called later
					return True
			self.row_index = 1 # assumed scrolled by 1
		else:
			for idx in xrange(self.vlist.n_items-1, self.vlist.n_items-5, -1):
				text = self.vlist.getText(idx)
				if not text:
					break
				if text == current_text:
					self.row_index = idx # _setTow sould be called later
					return True
			self.row_index = self.vlist.n_items-2 # assumed scrolled by 1			
		return True
			
	
	def focusNext(self):
		"""
		Containers process Tab as long as they can focus corresponsing control.
		"""
		if self.row_index < 0:
			return self.focusFirst()
		if self.row_index == self.vlist.n_items-1 and self.scroll and self.scroll.canScrollDown():
			self._tryScroll(False)
		return self._setRow(self.row_index + 1)
		
	def focusPrevious(self):
		"""
		Containers process Tab as long as they can focus corresponsing control.
		"""
		if self.row_index <= 0:
			if not self.scroll or self.row_index < 0 or not self.scroll.canScrollUp(): 
				return False
			if not self._tryScroll(True):
				return False
		return self._setRow(self.row_index - 1)
		

def cleanWindowCache(cache):
	for hwnd in list(cache):
		if not user32.IsWindow(hwnd):
			del cache[hwnd]


from NVDAObjects.window import Window
import speech
import time
import queueHandler
import core
import oleacc
import controlTypes


# NVDA creates all objects all the time... Confirmed that is by design.
# Not a good solution, but working
gSIBI = {}

def findKnownOverlay(obj):
	"""
	Some dialogs have several "windows", but in many cases they are not accessible. In this case
	we can assign (some) parent window SIBI to control it.
	Oberlay SIBI indicates it is ready to accept child elements by setting nvda_class
	"""
	global gSIBI
	if obj.role != controlTypes.ROLE_PANE and obj.role != controlTypes.ROLE_UNKNOWN:
		return None
	try:
		p = obj.parent
		while p:
			if p.windowHandle in gSIBI and gSIBI[p.windowHandle].nvda_class is not None:
				return gSIBI[p.windowHandle]
			if p.role == controlTypes.ROLE_DIALOG:
				break # That was "top level" and there was no SIBI for it
			p = p.parent
	except:
		pass
	return None

def chooseKnownOverlay(obj, clsList):
	"""
	Report overlay class for not accessible windows, hasKnownSIBI will find it later
	"""
	sibi = findKnownOverlay(obj)
	if not sibi:
		return False
	clsList.insert(0, sibi.nvda_class)
	return True
			
class SIBINVDA(Window):
	"""
	Basic class for an overlays using SIBIAC
	
	at least defineSIBI should be overloaded defined to be useful, should set self.sibi
	better set name or _get_name as well
	"""

	def _get_name(self):
		"""
		Cooperate with OSARA, which use accName for audition
		"""
		name = super(SIBINVDA,self)._get_name()
		if name is None or name == "" or name == self.windowText:
			name = getattr(self, 'sname', name)
		return name

	def hasKnownSIBI(self):
		"""
		To be used in defineSIBI, in addition to setting sibi.nvda_class and chooseKnownOverlay
		To point dialogs and covering panes to the same SIBI
		"""
		sibi = findKnownOverlay(self)
		if not sibi:
			return False
		self.sibi = sibi
		self.sibi.panel_defined = True
		return True

	def defineSIBI(self):
		"""
		SHOULD be overloaded with meaningful SIBI declaration
		"""
		self.sibi = SIBI(self.windowHandle, 1, 1, set())

	def hadFocus(self):
		"""
		Return True if it was ever focused
		"""
		return self.sibi.had_focus
		
	def _speakInfo(self):
		if self.sibi.info_queued > 1:
			self.sibi.info_queued -= 1
			return
		self.sibi.info_queued = 0
		ctrl = self.sibi.getFocused()
		if ctrl:
			type, name, text = ctrl.getTextInfo()
			speech.cancelSpeech()
			queueHandler.queueFunction(queueHandler.eventQueue, speech.speakMessage, text)
			#speech.speakMessage(text)

	def speakOnFocus(self, cancel = True):
		"""
		Callback 
		"""
		ctrl = self.sibi.getFocused()
		if ctrl:
			if cancel:
				speech.cancelSpeech()
			type, name, text = ctrl.getTextInfo()
			queueHandler.queueFunction(queueHandler.eventQueue, speech.speakMessage, type+' '+name+' '+text)

	def _speakFocus(self):
		"""
		When dialog and panel share the same SIBI gainFocus is sometimes called for both.
		I do not want "cancel", since I want dialog speach is not interrupted
		But I do not want duplicated focus message
		"""
		if self.sibi.focus_queued > 1:
			self.sibi.focus_queued -= 1
			return
		self.sibi.focus_queued = 0
		self.speakOnFocus(False)

	def speakInFocus(self, cancel = True):
		"""
		Callback 
		"""
		ctrl = self.sibi.getFocused()
		if ctrl:
			type, name, text = ctrl.getTextInfo()
			if cancel:
				speech.cancelSpeech()
			queueHandler.queueFunction(queueHandler.eventQueue, speech.speakMessage, name+' '+text)

	def _speakInFocus(self):
		"""
		When dialog and panel share the same SIBI gainFocus is sometimes called for both.
		I do not want "cancel", since I want dialog speach is not interrupted
		But I do not want duplicated focus message
		"""
		if self.sibi.infocus_queued > 1:
			self.sibi.infocus_queued -= 1
			return
		self.sibi.infocus_queued = 0
		self.speakInFocus()

	def speakAfter(self, delay):
		"""
		Callback 
		"""
		if delay == 0 and self.sibi.info_queued == 0:
			ctrl = self.sibi.getFocused()
			if ctrl:
				type, name, text = ctrl.getTextInfo()
				speech.cancelSpeech()
				speech.speakMessage(text)
		else:
			self.sibi.info_queued += 1
			core.callLater(delay, self._speakInfo)

	def speakFocusAfter(self, delay):
		"""
		Callback 
		"""
		if delay == 0 and self.sibi.focus_queued == 0:
			self.speakOnFocus()
		else:
			self.sibi.focus_queued += 1
			core.callLater(delay, self._speakFocus)

	def speakInFocusAfter(self, delay):
		"""
		Callback 
		"""
		if delay == 0 and self.sibi.infocus_queued == 0:
			self.speakInFocus()
		else:
			self.sibi.infocus_queued += 1
			core.callLater(delay, self._speakInFocus)

	def initOverlayClass(self):
		"""
		__init__ is NOT called for overlay classes, so should do initialization here
		"""
		global gSIBI
		if not self.windowHandle in gSIBI:
			cleanWindowCache(gSIBI)
			self.defineSIBI()
			self.sibi.speakAfter = self.speakAfter
			self.sibi.speakFocusAfter = self.speakFocusAfter
			self.sibi.speakInFocusAfter = self.speakInFocusAfter
			gSIBI[self.windowHandle] = self.sibi
		else:
			self.sibi = gSIBI[self.windowHandle]
	
	def event_gainFocus(self):
		super(SIBINVDA,self).event_gainFocus()
		over = findTopmostOver(self.sibi.hwnd)
		if over != "":
			speech.cancelSpeech()
			speech.speakMessage("Warning: %s window with topmost flag hide a part of the interface. Sibiac will not work correctly." % over)
		self.sibi.focusSet()
		if self.sibi.panel_defined:
			self.sibi.focus_queued += 1
			core.callLater(0.05, self._speakFocus)
		else:
			self.speakOnFocus(self.sibi.focusReturn());
		
	def _skipGesture(self, gesture):
		"""
		If control is recognized, it can process gestures. For example menu.
		"""
		if self.role == controlTypes.ROLE_UNKNOWN or self.role == controlTypes.ROLE_DIALOG  or self.role == controlTypes.ROLE_PANE:
			return False
		gesture.send()
		return True
		
	def script_focusNext(self, gesture):
		if self._skipGesture(gesture):
			return
		self.sibi.focusNext();
		self.speakOnFocus();

	def script_focusPrev(self, gesture):
		if self._skipGesture(gesture):
			return
		self.sibi.focusPrevious();
		self.speakOnFocus();
	
	def _onScript(self, method, send_by_default, gesture):
		if self._skipGesture(gesture):
			return
		ctrl = self.sibi.getFocused()
		if ctrl and hasattr(ctrl, method):
			if not getattr(ctrl, method)():
				gesture.send()
			return
		if method == "onHelp":
			speech.cancelSpeech()
			queueHandler.queueFunction(queueHandler.eventQueue, speech.speakMessage, "No help for current element")
		elif send_by_default:
			gesture.send()
	
	def script_up(self, gesture):
		self._onScript("onUp", False, gesture)

	def script_down(self, gesture):
		self._onScript("onDown", False, gesture)

	def script_left(self, gesture):
		self._onScript("onLeft", False, gesture)
		
	def script_right(self, gesture):
		self._onScript("onRight", False, gesture)
			
	def script_enter(self, gesture):
		self._onScript("onEnter", False, gesture)

	def script_escape(self, gesture):
		self._onScript("onEscape", True, gesture)

	def script_pageUp(self, gesture):
		self._onScript("onPageUp", True, gesture)

	def script_pageDown(self, gesture):
		self._onScript("onPageDown", True, gesture)

	def script_home(self, gesture):
		self._onScript("onHome", True, gesture)

	def script_end(self, gesture):
		self._onScript("onEnd", True, gesture)
		
	def script_maximize(self, gesture):
		maximizeWindow(self.sibi.hwnd)
	
	def script_delete(self, gesture):
		if self._skipGesture(gesture):
			return
		ctrl = self.sibi.getFocused()
		if ctrl and hasattr(ctrl, "onDelete") and ctrl.onDelete():
			return
		gesture.send()
	
	def script_help(self, gesture):
		self._onScript("onHelp", False, gesture)
		
	__gestures = {
		"kb:tab": "focusNext",
		"kb:shift+tab": "focusPrev",
		"kb:upArrow": "up",
		"kb:downArrow": "down",
		"kb:leftArrow": "left",
		"kb:rightArrow": "right",
		"kb:enter": "enter",
		"kb:escape": "escape",
		"kb:pageUp": "pageUp",
		"kb:pageDown": "pageDown",
		"kb:home": "home",
		"kb:end": "end",
		"kb:windows+upArrow": "maximize",
		"kb:delete": "delete",
		"kb:f1": "help",
	}

# Sibiac cache
import globalVars
import pickle
import sys

class Cache(object):
	"""
	Permanently save data into file. Root folder is <NVDA>\\sibcache
	"""
	def __init__(self, name, default = None):
		self.name = name
		self.default = default
	
	def _open(self, mode = "r"):
		sibcache_dir = os.path.abspath(os.path.join(globalVars.appArgs.configPath, "sibcache"))
		if not os.path.isdir(sibcache_dir):
			os.mkdir(sibcache_dir)
		file_name = os.path.join(sibcache_dir, self.name + ".pickle")
		return open(file_name, mode)
		
	def save(self, obj):
		try:
			f = self._open("wb")
			pickle.dump(obj, f)
			return True
		except:
			log.error(sys.exc_info()[1])
			return False

	def load(self):
		try:
			f = self._open("rb")
			return pickle.load(f)
		except:
			log.error(sys.exc_info()[1])
			return self.default

# restore watchdog
watchdog.alive()
