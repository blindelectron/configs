# -*- coding: utf-8 -*-
# Single Image Blob Interface Accessible Control
#
# SIBIAC API, v2
#
# Alexey Zhelezov (www.azslow.com), 2019
from __future__ import absolute_import
from __future__ import division
from __future__ import print_function

from ctypes import *
from ctypes.wintypes import *
import winUser
user32 = winUser.user32
import time
import watchdog
import os
from logHandler import log
import core

from .location import *
from .location import _LocationRef

watchdog.asleep() # can take a while...
sibiac_dir = os.path.dirname(__file__)
dll = os.path.join(sibiac_dir, "..", "sibiac", "libsibiac.dll")
libsibiac = cdll.LoadLibrary(dll)
watchdog.alive()

def Colors2Tuple(arg, default = None):
	""" Convert an argument to tuple
	Args:
		arg (tuple of ints or single int, optional): the argument to convert
	"""
	if arg is None:
		if default is None:
			return None
		else:
			arg = default
	if isinstance(arg, tuple):
		return arg
	if isinstance(arg, int):
		return (arg,)		
	return tuple(arg)

def GetClientSize(windowHandle):
	""" Return the size of the window (NVDA had a bug in getClientRect)
	Returns:
		tuple: width and height of the window
	"""
	r = RECT()
	user32.GetClientRect(windowHandle, byref(r))
	return (r.right, r.bottom)

class InterfaceWindow(_LocationRef):
	""" Location for a real window. It is not derived from Location class since near everything is special.
		But it can be used as Location.
	"""
	
	def __init__(self, windowHandle, original_width, original_height, defRelation = LABEGIN):
		"""Initialize location.
		Args:
			windowHandle (int): handle for the window
			original_width, original_height (int): "original" size of the window, for absolute relations calculation
			defRelation (int, optional): default relation for this (and so hierarchy) locations
		"""	
		super(InterfaceWindow, self).__init__()
		self.windowHandle = windowHandle
		self._defRelation = defRelation
		self._originalHSegment = (0, original_width - 1)
		self._originalVSegment = (0, original_height - 1)
		self.x = 0 # we always work in client coordinates
		self.y = 0
		self.w, self.h = GetClientSize(windowHandle)
		self.r = self.w - 1
		self.b = self.h - 1

	def __call__(self, x, y, r = None, b = None, defRelation = None):
		""" create new WindowLocation with this one as reference. Just simplify coding
		Args:
			x,y,r,b (tuple): left, top, right, bottom positions specification, see _LocationDynamicPosition.
						when r and/or b are not specified, corrsponding x/y is used
			defRelation (int, optional): can be set here, refLocation defRelation or LBEGIN used otherwise
		"""
		return WindowLocation(self, x, y, r, b, defRelation)
		
	def originalHSegment(self):
		"""Ruturns own original horisontal position.
		That is not always required, so not pre-calculated at initialization.
		But once calculated, it is constant by definition.

		Returns:
			tuple: segment
		"""
		return self._originalHSegment

	def originalVSegment(self):
		"""Ruturns own original vertical position.
		That is not always required, so not pre-calculated at initialization.
		But once calculated, it is constant by definition.

		Returns:
			tuple: segment
		"""
		return self._originalVSegment

	def checkSize(self):
		"""Checks the size of the window. If changed, inform all related locations
		"""
		w, h = GetClientSize(self.windowHandle)
		if w != self.w or h != self.h:
			self.w = w
			self.h = h
			self.r = w - 1
			self.b = h - 1
			self._changed()

	def __str__(self):
		return "IWindow(x=%d,y=%d,r=%d,b=%d)" % (self.x, self.y, self.r, self.b)

	__repr__ = __str__
	
class WindowLocation(Location):
	""" Location with window handle, several color based operations and mouse operations
	"""
	def __init__(self, refLocation, x, y, r = None, b = None, defRelation = None):
		"""
		Args:
			refLocation (WindowLocation or InterfaceWindow): default reference location to use for coordinates
			x,y,r,b (tuple): left, top, right, bottom positions specification, see _LocationDynamicPosition.
							 when r and/or b are not specified, corrsponding x/y is used
			defRelation (int, optional): can be set here, refLocation defRelation or LBEGIN used otherwise
		"""
		super(WindowLocation, self).__init__(refLocation, x, y, r, b, defRelation)
		self.windowHandle = refLocation.windowHandle

	def __call__(self, x, y, r = None, b = None, defRelation = None):
		""" create new WindowLocation with this one as reference. Just simplify coding
		Args:
			x,y,r,b (tuple): left, top, right, bottom positions specification, see _LocationDynamicPosition.
						when r and/or b are not specified, corrsponding x/y is used
			defRelation (int, optional): can be set here, refLocation defRelation or LBEGIN used otherwise
		"""
		return WindowLocation(self, x, y, r, b, defRelation)
		
	def leftClick(self):
		""" Left click on x,y position
		"""
		return libsibiac.MouseLeftClick(self.windowHandle, c_int(self.x), c_int(self.y))
	
	def leftDblClick(self):
		""" Double left click on x,y position
		"""
		self.leftClick()
		time.sleep(0.01)
		self.leftClick()
	
	def rightClick(self):
		""" Right click on x,y position
		"""
		return libsibiac.MouseRightClick(self.windowHandle, c_int(self.x), c_int(self.y))

	def moveTo(self):
		""" Move mouse to x,y position
		"""
		libsibiac.MouseMove(self.windowHandle, c_int(self.x), c_int(self.y))

	def leftDrag(self, timeout = 0.3, mouse_up = True):
		""" Drag with left button from x,y to r,b coordinates.
		Dangerous! When mouse_up is false, should be paired with leftUp !!!
		Args:
			timeout (float): time (in seconds) to wait before moving from initial to final position
			mouse_up (bool): generate mouse_up event after dragging (!!! Dangerous !!!) 
		"""
		self.moveTo()
		winUser.mouse_event(winUser.MOUSEEVENTF_LEFTDOWN,0,0,None,None)
		watchdog.asleep() # do not let watchdog interrupt us
		time.sleep(timeout)
		watchdog.alive()
		libsibiac.MouseMove(self.windowHandle, c_int(self.r), c_int(self.b))
		if mouse_up:
			winUser.mouse_event(winUser.MOUSEEVENTF_LEFTUP,0,0,None,None)

	def mouseScroll(self, delta = 120):
		""" Move to x,y and generate mouse scroll event with specified delta
		Args:
			delta (int): scroll parameter, normally 120 for up and -120 for down
		"""
		self.moveTo()
		libsibiac.MouseScroll(c_int(delta))
		
	def leftDown(self):
		""" Move to x, y and press left mouse button. Dangerous! Should be paired with leftUp !!!"""
		self.moveTo()
		winUser.mouse_event(winUser.MOUSEEVENTF_LEFTDOWN,0,0,None,None)

	def leftUp(self):
		""" Release left mouse button. Dangerous! Should be paired with leftDown !!!"""
		winUser.mouse_event(winUser.MOUSEEVENTF_LEFTUP,0,0,None,None)
		
	def FindNearestColor(self, colors):
		""" Compare the color at x,y position with given colors.
		Args:
			colors (iterable): the list of (RGB, int) colors to compare with
		Returns:
			int: the index of nearest color in the arg on -1 on error
		"""
		colors = Colors2Tuple(colors)
		n = len(colors)
		if n < 1:
			return -1
		ccolors = (c_int*n)(*colors)
		return libsibiac.NearestColor(self.windowHandle, c_int(self.x), c_int(self.y), c_int(n), byref(ccolors))

	def PixelColor(self):
		""" Return pixel color from x,y coordinates
		Returns:
			int: RGB color value
		"""
		return libsibiac.PixelColor(hwnd, c_int(self.x), c_int(self.y))

	def FindRow(self, bg, fg, lastRow = None):
		""" Find a row in location based on colors.
			There must be bg only colored lines before the row.
			Continues lines with at least one fg colored pixel assumed to be the row.
		Args:
			bg (tuple or int): bg color(s)
			fg (tuple or int): fg color(s)
			last_row (WindowLocation, optional): previously found row (when given, this method starts searching from that row)
		Returns:
			WindowLocation or None: found row location or None
		"""
		bg = Colors2Tuple(bg)
		fg = Colors2Tuple(fg)
		y = self.y if lastRow is None else lastRow.b
		r = RECT(self.x, y, self.r, self.b)
		# log.error("(%d,%d)-(%d,%d), %d, %d" % (self.x, y, self.r, self.b, len(bg), len(fg)))
		row_y = c_int(0)
		row_height = libsibiac.FindRow(self.windowHandle, byref(r),  c_int(len(bg)), (c_int*len(bg))(*bg), c_int(len(fg)), (c_int*len(fg))(*fg), byref(row_y))
		if row_height < 1:
			return None
		row_y = row_y.value
		return self(0, row_y - self.y - 1, (0, LEND), row_y + row_height - self.y, defRelation = LBEGIN) # we add 1 bg line before and after to the location

	def getTextOCR(self, tag = None, ch = None):
		"""
		Get text from the box (OCR, slow)
		"""
		if tag is not None:
			crc = libsibiac.Screen2CRC(self.windowHandle, c_int(self.x), c_int(self.y), c_int(self.r), c_int(self.b))
			text = ocrCache.find(tag, crc)
			if text is not None:
				return text
		buf = create_string_buffer(1024)
		if ch is None:
			libsibiac.Screen2Text(self.windowHandle, c_int(self.x), c_int(self.y), c_int(self.r), c_int(self.b), buf, len(buf))
		else:
			libsibiac.Screen2TextCh(self.windowHandle, c_int(ch), c_int(self.x), c_int(self.y), c_int(self.r), c_int(self.b), buf, len(buf))
		text = buf.value.decode('utf-8', 'ignore').encode('ascii', 'ignore').replace('\n', '')
		if tag is not None:
			ocrCache.add(tag, crc, text)
		return text
		
class VScrollElement(WindowLocation):
	""" A scroll graphical element.
	
	Scroll location should be upper and lower possible position element, x should be within position element
	"""
	def __init__(self, refLocation, x, t, b, bg, fg, defRelation = None):
		"""
		Args:
			refLocation (WindowLocation or InterfaceWindow): default reference location to use for coordinates
			x (tuple): some coordinate within position element
			t, b (tuple): topmost and bottommost pixels which can be occupied by position element
			bg (tuple or int): bg colors for position element
			fg (tuple or int): fg colors for position element
			defRelation (int, optional): can be set here, refLocation defRelation or LBEGIN used otherwise
		"""
		super(VScrollElement, self).__init__(refLocation, x, t, None, b, defRelation)
		self.bg = Colors2Tuple(bg)
		self.fg = Colors2Tuple(fg)
	
	def _position(self):
		""" search for position element and returns its coordinates
		Returns:
			tuple: (top, bottom) of the position element (inclusive), (-1,-1) when not found
		"""
		r = RECT(self.x, self.y, self.x, self.b)
		allcolors = self.bg + self.fg
		n = len(allcolors)
		allcolors = (c_int*n)(*allcolors)
		height = libsibiac.FindVRange(self.windowHandle, byref(r), c_int(n), c_int(len(self.bg)), byref(allcolors))
		if height < 1:
			return (-1, -1)
		return (r.top, r.bottom)
	
	def scrollUp(self, delta = 120):
		""" Scroll up, if possible.
		Args:
			delta (int): delta value for mouse event
		Returns:
			bool: True in case we made an attempt to scroll, False otherwise (position not found or already at the top)
		"""
		t,b = self._position()
		if t < 0 or t <= self.y:
			log.error("No position found...")
			return False
		
		self.mouseScroll(delta)
		return True

	def scrollDown(self, delta = -120):
		""" Scroll down, if possible.
		Args:
			delta (int): delta value for mouse event
		Returns:
			bool: True in case we made an attempt to scroll, False otherwise (position not found or already at the bottom)
		"""
		t,b = self._position()
		if t < 0 or b >= self.b:
			return False
		self.mouseScroll(delta)
		return True

	def pageUp(self):
		""" Click on top pixel, if possible.
		Returns:
			bool: True in case we made an attempt to scroll, False otherwise (position not found or already at the top)
		"""
		t,b = self._position()
		if t < 0 or t <= self.y:
			return False
		self.leftClick()
		return True

	def pageDown(self):
		""" Click on bottom pixel, if possible.
		Returns:
			bool: True in case we made an attempt to scroll, False otherwise (position not found or already at the bottom)
		"""
		t,b = self._position()
		if t < 0 or b >= self.b:
			return False
		self(0, 0, defRelation=LEND).leftClick()
		return True
		
	def isVisible(self):
		"""
		Returns:
			bool: True if we can find position element 
		"""
		t,b = self._position()
		return t >= 0

	def toBegin(self, timeout = 0.3):
		""" Drag position element to the top, if possible.
		Args:
			timeout (float): time (in seconds) to wait before moving from initial to final position
		Returns:
			bool: True in case we made an attempt to scroll, False otherwise (position not found or already at the top)
		"""
		t,b = self._position()
		if t < 0 or t <= self.y:
			return False
		y = (b + t)/2 - self.y # we drage the middle of position
		self(0, y, None, y - (t - self.y), defRelation = LBEGIN).leftDrag(timeout)
		return True

	def toEnd(self, timeout = 0.3):
		""" Drag position element to the bottom, if possible.
		Args:
			timeout (float): time (in seconds) to wait before moving from initial to final position
		Returns:
			bool: True in case we made an attempt to scroll, False otherwise (position not found or already at the bottom)
		"""
		t,b = self._position()
		if t < 0 or b >= self.b:
			return False
		y = (b + t)/2 - self.y # we drage the middle of position
		self(0, y, None, y + (self.b - b), defRelation = LBEGIN).leftDrag(timeout)
		return True
		
# Sibiac cache
import globalVars
import pickle

class Cache(object):
	""" Permanently save data into file. Root folder is <NVDA>\\sibcache
	"""
	def __init__(self, name, default = None):
		"""
		Args:
			name (str): the file name
			default (obj): default value to return on load, when there is no file or other errors
		"""
		self.name = name
		self.default = default
	
	def _open(self, mode = "r"):
		sibcache_dir = os.path.abspath(os.path.join(globalVars.appArgs.configPath, "sibcache"))
		if not os.path.isdir(sibcache_dir):
			os.mkdir(sibcache_dir)
		file_name = os.path.join(sibcache_dir, self.name + ".pickle")
		return open(file_name, mode)
		
	def save(self, obj):
		""" Save object into cache file
		Args:
			obj (object): object to save
		Returns:
			bool: if saving was successful
		"""
		try:
			f = self._open("w")
			pickle.dump(obj, f)
			f.close()
			return True
		except:
			return False

	def load(self):
		""" Load object from the cache file
		Returns:
			object: loaded object or sepcified at creation default
		"""
		try:
			f = self._open("r")
			o = pickle.load(f)
			f.close()
			return 0
		except:
			return self.default

class OCRCache(object):
	""" Save CRC:text map for returned by OCR text
		There are separate files for each module and separate maps for each element.
		So each map is tagged with '<module>:<element>'.
	"""
	def __init__(self):
		self._cache = {} # the cache
		self._saveTimer = None # timer to save the cache
		
	def find(self, tag, crc):
		""" Find crc in tag cache.
		Args:
			tag (str): the tag for cache in format 'module:element'
			crc (int): image CRC
		Returns:
			str or None: text if found, None otherwise
		"""
		try:
			module, element = tag.split(":")
			self._loadModule(module)
			return self._cache[module][element][crc]
		except:
			return None
	
	def add(self, tag, crc, text):
		""" Add crc+text under tag into cache.
		Args:
			tag (str): the tag for cache in format 'module:element'
			crc (int): image CRC
			text (str): text to associate with CRC
		"""
		try:
			module, element = tag.split(":")
			if module in self._cache:
				if element in self._cache[module]:
					if crc in self._cache[module][element]:
						if self._cache[module][element][crc] == text:
							return # already there and match
					self._cache[module][element][crc] = text
				else:
					self._cache[module][element] = {}
			else:
				self._loadModule(module)
				self._cache[module][element] = {}
			self._cache[module][element][crc] = text
			self._cache[module]["changed"] = True
			if self._saveTimer is None:
				self._saveTimer = core.callLater(5000, self._saveAll)
			else:
				self._saveTimer.Start(5000)
		except:
			pass
			
	def _loadModule(self, module):
		if module in self._cache:
			return
		c = Cache("%s" % module).load()
		self._cache[module] = c if c is not None else {}
		
	def _saveAll(self):
		""" we save all files on exit/unload, when they was changed
		"""
		for module, m in self._cache.iteritems():
			if "changed" not in m:
				continue
			del m["changed"]
			Cache("%s" % module).save(m)

# we need only one
ocrCache = OCRCache()
