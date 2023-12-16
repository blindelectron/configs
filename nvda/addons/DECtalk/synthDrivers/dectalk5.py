from cStringIO import StringIO
import synthDriverHandler
from synthDriverHandler import SynthDriver,VoiceInfo, NumericSynthSetting
from collections import OrderedDict
from logHandler import log
from ctypes import *
from ctypes.wintypes import *
import config
import speech
from winUser import WNDCLASSEXW, WNDPROC
import tones
import logging
import os
from nvwave import outputDeviceNameToID
import nvwave
import threading
from Queue import Queue
audioQueue = Queue() #for audio and indexes
audio_data = StringIO()

OWN_AUDIO_DEVICE=1
DO_NOT_USE_AUDIO_DEVICE=0x80000000
samples=8192
buffer=create_string_buffer(samples*2)
#11025khz mono
format=0x00000004
player = None
speaking = False

def bgPlay(s):
	player.feed(s)

def clear_queue(q):
	try:
		while True:
			q.get_nowait()
	except:
		pass

class BgThread(threading.Thread):
	def __init__(self, q):
		threading.Thread.__init__(self)
		self.setDaemon(True)
		self.q = q

	def run(self):
		try:
			while True:
				func, args, kwargs = self.q.get()
				if not func:
					break
				func(*args, **kwargs)
				self.q.task_done()
		except:
			logging.error("bgThread.run", exc_info=True)

def _bgExec(q, func, *args, **kwargs):
	q.put((func, args, kwargs))

def setLast(index):
	global lastIndex
	lastIndex = index
class TTS_CAPS_T(Structure):
	_fields_ = [
		('dwNumberOfLanguages', DWORD),
		('lpLanguageParamsArray', c_void_p), #don't care
		('dwSampleRate', DWORD),
		('dwMinimumSpeakingRate', DWORD),
		('dwMaximumSpeakingRate', DWORD),
		('dwNumberOfPredefinedSpeakers', DWORD),
		('dwCharacterSet', DWORD),
		('version', DWORD),
	]

class TTS_INDEX_T(Structure):
	_fields_ = [
		('dwIndexValue', DWORD),
		('dwIndexSampleNumber', DWORD),
		('dwReserved', DWORD),
	]
index_array_size=3
class TTS_BUFFER_T(Structure):
	_fields_ = [
	('lpData', POINTER(c_char*(samples*2))),
	('lpPhonemeArray', c_void_p),
	('lpIndexArray', POINTER(index_array_size*TTS_INDEX_T)),
	('dwMaximumBufferLength', DWORD),
	('dwMaximumNumberOfPhonemeChanges', DWORD),
	('dwMaximumNumberOfIndexMarks', DWORD),
	('dwBufferLength', DWORD),
	('dwNumberOfPhonemeChanges', DWORD),
	('dwNumberOfIndexMarks', DWORD),
	('dwReserved', DWORD),
	]

def errcheck(res, func, args):
	if res != 0:
		raise RuntimeError("%s: code %d" % (func.__name__, res))
	return res

dtPath=os.path.abspath(os.path.join(os.path.dirname(__file__), r"dectalk5.dll"))
dt = cdll.LoadLibrary(dtPath)
#dt = cdll.dectalk
for i in ('TextToSpeechStartup', 'TextToSpeechSpeak', 'TextToSpeechAddBuffer', 'TextToSpeechOpenInMemory', 'TextToSpeechGetCaps'):
	getattr(dt, i).errcheck = errcheck

#set later
WM_INDEX = 0
WM_BUFFER=0
lastIndex = None

appInstance = windll.kernel32.GetModuleHandleW(None)
nvdaDtSoftWndCls = WNDCLASSEXW()
nvdaDtSoftWndCls.cbSize = sizeof(nvdaDtSoftWndCls)
nvdaDtSoftWndCls.hInstance = appInstance
nvdaDtSoftWndCls.lpszClassName = u"nvdaDtSoftWndCls"
class SynthDriver(synthDriverHandler.SynthDriver):
	name="dectalk5"
	description=_("dectalk5")
	language=None
	supportedSettings = (SynthDriver.VoiceSetting(), SynthDriver.RateSetting(), SynthDriver.PitchSetting(), SynthDriver.InflectionSetting(),
#	NumericSynthSetting("sonic", "sonic") 
)
	minSonic = 1.0
	maxSonic = 4.0
	minRate = 75
	maxRate = 650
	minPitch=50
	maxPitch=350
	minInflection=0
	maxInflection=250

	@classmethod
	def check(cls):
		return True #fixme

	def __init__(self):
		global WM_INDEX, WM_BUFFER, player, g_stream
		self.mapping = windll.kernel32.CreateFileMappingA(0xFFFFFFFF, 0, 4, 0, 512, "a32DECtalkDllFileMap")
		buf = windll.kernel32.MapViewOfFile(self.mapping, 2, 0, 0, 0)
		array = (c_char*512).from_address(buf)
		array.value = '\0\0\0\0r250hRm2no9fmP75YwvRhnRB81Uv6vZOTb7SdKWKae8k3BXL8U6r??3B0P91'
		self.dt_rate = 350
		self.dt_pitch=100
		self.dt_inflection=100
		self._voice = 'Paul'
		self.setup_wndproc()
		self._messageWindowClassAtom = windll.user32.RegisterClassExW(byref(nvdaDtSoftWndCls))
		self._messageWindow = windll.user32.CreateWindowExW(0, self._messageWindowClassAtom, u"nvdaDtSoftWndCls window", 0, 0, 0, 0, 0, None, None, appInstance, None)
		self.handle = c_void_p()
		cwd = os.getcwd()
		os.chdir(os.path.dirname(__file__))
		dt.TextToSpeechStartup(self._messageWindow, byref(self.handle), 0, DO_NOT_USE_AUDIO_DEVICE)
		caps = TTS_CAPS_T()
		dt.TextToSpeechGetCaps(byref(caps))
		self.maxRate = caps.dwMaximumSpeakingRate
		self.minRate = caps.dwMinimumSpeakingRate
		os.chdir(cwd)
		WM_INDEX = windll.user32.RegisterWindowMessageW(u"DECtalkIndexMessage")
		WM_BUFFER = windll.user32.RegisterWindowMessageW(u"DECtalkBufferMessage")
		self.mem_buffer = TTS_BUFFER_T()
		self.index_array = (TTS_INDEX_T*index_array_size)()
		self.mem_buffer.lpData = pointer(buffer)
		self.mem_buffer.lpIndexArray = pointer(self.index_array)
		self.mem_buffer.dwMaximumBufferLength = samples*2
		self.mem_buffer.dwMaximumNumberOfIndexMarks=index_array_size
		dt.TextToSpeechOpenInMemory(self.handle, format)
		dt.TextToSpeechAddBuffer(self.handle, byref(self.mem_buffer))
		player = nvwave.WavePlayer(1, 11025, 16, outputDevice=config.conf["speech"]["outputDevice"])
		self.audio_thread = BgThread(audioQueue)
		self.audio_thread.start()

	def speak(self,speechSequence):
		global speaking
		text_list = []
		for item in speechSequence:
			if isinstance(item, basestring):
		#prevent control strings from going into our text from input
				item = item.replace('[:', ' ')
				text_list.append(item)
			elif isinstance(item, speech.IndexCommand):
				text_list.append(u"[:i m %d] " % item.index)
		text = " ".join(text_list).encode('iso8859-1', 'ignore')
		if text:
			#skip rule kills some abbrevs
			text = "[:skip rule :ra %d :dv ap %d] %s" % (self.dt_rate, self.dt_pitch, text)
			voice = self._voice[0].lower()
			text = "[:n%s :dv pr %d] %s [:i r 32000]" % (voice, self.dt_inflection, text)
			speaking = True
			dt.TextToSpeechSpeak(self.handle, text, 1)

	def cancel(self):
		global speaking
		speaking = False
		dt.TextToSpeechReset(self.handle, 0)
		clear_queue(audioQueue)
		audio_data.truncate(0)
		player.stop()

	def _set_rate(self, rate):
		val = self._percentToParam(rate, self.minRate, self.maxRate)
		self.dt_rate = val

	def _get_rate(self):
		return self._paramToPercent(self.dt_rate, self.minRate, self.maxRate)

	def _set_pitch(self, pitch):
		val = self._percentToParam(pitch, self.minPitch, self.maxPitch)
		self.dt_pitch = val

	def _get_pitch(self):
		return self._paramToPercent(self.dt_pitch, self.minPitch, self.maxPitch)

	def _set_inflection(self, inflection):
		val = self._percentToParam(inflection, self.minInflection, self.maxInflection)
		self.dt_inflection = val

	def _get_inflection(self):
		return self._paramToPercent(self.dt_inflection, self.minInflection, self.maxInflection)

	def _get_lastIndex(self):
		return lastIndex

	def terminate(self):
		self.cancel()
		dt.TextToSpeechShutdown(self.handle)
		windll.user32.DestroyWindow(self._messageWindow)
		windll.user32.UnregisterClassW(self._messageWindowClassAtom,appInstance)
		audio_data.truncate(0)
		audioQueue.put((None, None, None))

	def setup_wndproc(self):
		@WNDPROC
		def nvdaDtSoftWndProc(hwnd, msg, wParam, lParam):
			global lastIndex, speaking
			if msg == WM_INDEX:
				lastIndex = lParam
			elif msg == WM_BUFFER:
				lpBuffer = cast(lParam, POINTER(TTS_BUFFER_T))
				self.handle_buffer(lpBuffer)
			return windll.user32.DefWindowProcW(hwnd, msg, wParam, lParam)
		nvdaDtSoftWndCls.lpfnWndProc = nvdaDtSoftWndProc

	def handle_buffer(self, lpBuffer):
		lastmark=None
		indexes=None
		end = False

		data = string_at(lpBuffer.contents.lpData, lpBuffer.contents.dwBufferLength)
		audio_data.write(data)
		if lpBuffer.contents.dwBufferLength == 0:
			end = True
		lpBuffer.contents.dwBufferLength=0
		if lpBuffer.contents.dwNumberOfIndexMarks:
			indexes = lpBuffer.contents.lpIndexArray.contents
			marks = lpBuffer.contents.dwNumberOfIndexMarks
			if indexes[marks-1].dwIndexValue == 32000:
				end = True
				if marks > 1:
					lastmark = indexes[marks-2].dwIndexValue
				else:
					lastmark = 32000
			else: #index below 32k
				lastmark = indexes[marks-1].dwIndexValue

		if audio_data.tell() >= 8192 or end:
			_bgExec(audioQueue, bgPlay, audio_data.getvalue())
			audio_data.truncate(0)

		if lastmark is not None and lastmark is not 32000:
			_bgExec(audioQueue, setLast, lastmark)

		dt.TextToSpeechAddBuffer(self.handle, lpBuffer)

	def _get_availableVoices(self):
		voices = OrderedDict()
		for i in 'Paul Betty Harry Frank Dennis Kit Ursula Rita Wendy'.split(' '):
			voices[i] = VoiceInfo(i, i)
		return voices

	def _get_voice(self):
		return self._voice

	def _set_voice(self, val):
		self._voice = str(val)
