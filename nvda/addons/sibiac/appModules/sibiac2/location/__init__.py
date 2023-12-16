# -*- coding: utf-8 -*-
#
# Alexey Zhelezov, 2019
"""
Location defines a rectangle by its boundaries.

In case primary window size is changed, the interface layout is normally scaled using some rules.
Location allows defining position of particular controls and other elements using relations to ancestors.
That way correct position can be automatically calculated after primary window is re-sized.

Since Location use relative coordinates to arbitrary rectangle (not only real window), that can be used
to transparently define blocks (f.e. stomb boxes) which can apppear on different places.

To simplify required coordinates capturing without calculating relations manually, Location support
"absolute" coordinates. In case absolute position of some element in hierarchy is known for particular
top element size, and all dynamic size rules are correctly identified, relative positions in hierarchy
can be automatically calculated from that absolute position.

"""
import weakref

# Relation type specify how position is calculate from the reference location position and parameter.
# The relation is to the reference position in the same axis as currently avaluated position,
# so left and rigt are used for x position, top and buttom for y position. 
LBEGIN = 1   # parameter is a shift from reference begin position
LEND   = 2   # parameter is a shift from reference end position
LCENTER = 3  # parameter is a shift from current center of reference 
LPROP = 4    # parameter is proportional (in %) to the reference size

# Absolute position. Specified number is a position in reference (hierarchy) with predefine size.
# Used in description only, required parameter is calculated at initialization and defined before types are used afterwards.
LABEGIN = 11
LAEND = 12
LACENTER = 13
LAPROP = 14

# Specify that default relation for location should be used
LDEFAULT = 0

class _LocationPosition(object):
	"""An abstract class with methods Location objects assume every position has
	"""
	def __init__(self, location, spec, refLocation = None):
		"""Initialization from given parameters

		Args:
			location (Location): this position is a part of it (must me specified)
			spec (tuple or int): position specification as (value, <relation>, <reference location>), at least value should be specified.
			refLocation (Location, optional): default reference Location.
		"""
		pass

	def value(self):
		"""Calculate and return current position value.
		Returns:
			int: the value
		"""
		raise NotImplementedError

	def originalValue(self):
		"""Return the value in upper Location with original size
		Returns:
		   int: the value for 
		"""
		raise NotImplementedError

	
class _LocationDynamicPosition(_LocationPosition):
	"""Base class for dynamic Location coordinates.
	The position is calculated using reference C{Location} (_ref), relation to the reference location (_rel)
	and _par. The interpretation of _par is relation type specific (see relation types)

	The position can be defined using "original coordinates", so absolute position inside upper level location.
	When original dimentions of that upper level location are known, required parameter for relative position
	inside location tree can be automatically calculated, see Relation types above.
	"""
	def __init__(self, location, spec, refLocation = None):
		"""Initialization from given parameters

		Args:
			location (Location): this position is a part of it (must me specified)
			spec (tuple or int): position specification as (value, <relation>, <reference location>), at least value should be specified.
			refLocation (Location, optional): default reference Location.
		"""
		if isinstance(spec, int):
			spec = (spec,)
		elif isinstance(spec, _LocationDynamicPosition):
			self._par = vtuple._par
			self._ref = vtuple._ref
			self._rel = vtuple._rel
			return
		slen = len(spec)
		self._par = spec[0] if slen > 0 else 0
		self._ref = spec[2] if slen > 2 else refLocation
		self._rel = spec[1] if slen > 1 else LDEFAULT
		if self._rel == LDEFAULT:
			self._rel = location._defRelation
			
		if self._ref is None:
			self._rel = LBEGIN # all other are not going to work
			
		if self._rel >= LABEGIN:
			self._par = self.__parFromAbsolute()

		if self._ref:
			self._ref._addRef(location)

	def _value(self, b, e):
		"""Calculate position value from given parameters and own specification.
		Args:
			b (int): reference segment begin
			e (int): reference segment end
		Returns:
			int: position value
		"""
		if self._ref is None:
			return self._par
		if self._rel == LBEGIN:
			return b + self._par
		if self._rel == LEND:
			return e + self._par
		if self._rel == LCENTER:
			return int((b + e)/2 + self._par)
		if self._rel == LPROP:
			return int(b + (e - b)*self._par/100.)
		return 0

	def _originalRefSegment(self):
		"""Ruturns segment of original sized reference (for parameter calcluation from absolute position)
		Returns:
			tuple: segment
		"""
		raise NotImplementedError
	
	def __parFromAbsolute(self):
		"""Return _par value calculated from absolute position, assuming it is currently in _par.
		_rel is adjusted.

		Returs:
		   int: usable _par value for final _rel
		"""
		self._rel -= 10
		b, e = self._originalRefSegment()
		if self._ref is None:
			return self._ref
		if self._rel == LBEGIN:
			return self._par - b
		if self._rel == LEND:
			return self._par - e
		if self._rel == LCENTER:
			return int(self._par - (b + e)/2)
		if self._rel == LPROP:
			return (self._par - b)*100./(e - b)
		return 0        

	def originalValue(self):
		"""Return the value in upper Location with original size
		Returns:
		   int: the value for 
		"""
		b, e = self._originalRefSegment()
		return self._value(b, e)

	
class _LocationX(_LocationDynamicPosition):
	"""
	Horisontal position
	"""
	def value(self):
		"""Calculate and return current position value.
		Returns:
			int: the value
		"""
		return self._value(self._ref.x, self._ref.r) if self._ref else self._par
			
	def _originalRefSegment(self):
		"""Ruturns segment of original sized reference (for parameter calcluation from absolute position)
		Returns:
			tuple: segment
		"""
		return self._ref.originalHSegment() if self._ref else (0, 0)


class _LocationY(_LocationDynamicPosition):
	"""
	Vertical position
	"""
	def value(self):
		"""Calculate and return current position value.
		Returns:
			int: the value
		"""
		return self._value(self._ref.y, self._ref.b) if self._ref else self._par
			
	def _originalRefSegment(self):
		"""Ruturns segment of original sized reference (for parameter calcluation from absolute position)
		Returns:
			tuple: segment
		"""
		return self._ref.originalVSegment() if self._ref else (0, 0)

		
class _LocationRef(object):
	"""
	(Weak) referene hierarchy
	That about memory leak by keeping all ever created locations, while informing
	about changes every location which currently exist.
	"""
	def __init__(self):
		self._refs = weakref.WeakValueDictionary()

	def _changed(self):
		"""Should be called to inform dependent locations
		"""
		for key in self._refs.keys():
			p = self._refs[key]
			if p is not None:
				p._changed()

	def _addRef(self, location):
		"""Add gived location to dependency list
		Args:
			location (Location): location to add
		"""
		pid = id(location)
		if pid not in self._refs:
			self._refs[pid] = location

class Location(_LocationRef):
	"""
	Location rectangle. When it covers just one pixel it is essentially a point.
	"""
	def __init__(self, refLocation, x, y, r = None, b = None, defRelation = None):
		"""
		Args:
			refLocation (Location): default reference location to use for coordinates
			x,y,r,b (tuple): left, top, right, bottom positions specification, see _LocationDynamicPosition.
						when r and/or b are not specified, corrsponding x/y is used
			defRelation (int, optional): can be set here, refLocation defRalation or LBEGIN used otherwise
		"""
		super(Location, self).__init__()
		self._defRelation = defRelation if defRelation is not None else (refLocation._defRelation if refLocation else LBEGIN) 
		self._x = _LocationX(self, x, refLocation)
		self._y = _LocationY(self, y, refLocation)
		self._r = _LocationX(self, r, refLocation) if r is not None else self._x
		self._b = _LocationY(self, b, refLocation) if b is not None else self._y
		self._changed() # calculate location the first time
		
	def _changed(self):
		self.x = self._x.value()
		self.y = self._y.value()
		self.r = self._r.value()
		self.b = self._b.value()
		super(Location, self)._changed() # important to call it when our location is re-calculated

	def originalHSegment(self):
		"""Ruturns own original horisontal position.
		That is not always required, so not pre-calculated at initialization.
		But once calculated, it is constant by definition.

		Returns:
			tuple: segment
		"""
		try:
			return self._originalHSegment
		except:
			# not calculated yet
			self._originalHSegment = (self._x.originalValue(), self._r.originalValue())
		return self._originalHSegment

	def originalVSegment(self):
		"""Ruturns own original vertical position.
		That is not always required, so not pre-calculated at initialization.
		But once calculated, it is constant by definition.

		Returns:
			tuple: segment
		"""
		try:
			return self._originalVSegment
		except:
			# not calculated yet
			self._originalVSegment = (self._y.originalValue(), self._b.originalValue())
		return self._originalVSegment


	@property
	def w(self):
		return self.r - self.x + 1

	@property
	def h(self):
		return self.b - self.y + 1


	def __str__(self):
		if self.x == self.r and self.y == self.b:
			return "LPoint(x=%d,y=%d)" % (self.x, self.y)
		return "LRect(x=%d,y=%d,r=%d,b=%d)" % (self.x, self.y, self.r, self.b)

	__repr__ = __str__

# Short form for Location Point
# LP = Location

#### example
#box = Location(None, 0, 0, 99, 99)
#section = Location(box, 20, 20, (0,LEND), (0,LEND))
#point = LP(section, (50., LPROP), (50., LPROP))

#box._x = _LocationX(box, 30)
#box._changed()
