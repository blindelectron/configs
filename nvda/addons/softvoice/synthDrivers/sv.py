import tones
from collections import deque, OrderedDict
import os
from io import StringIO
import synthDriverHandler
from autoSettingsUtils.driverSetting import DriverSetting, NumericDriverSetting
from synthDriverHandler import SynthDriver, VoiceInfo
from logHandler import log
from ctypes import *
from ctypes.wintypes import *
import config
import speech
from winUser import WNDCLASSEXW, WNDPROC
import threading
import queue
from speech.commands import IndexCommand

bgQueue = queue.Queue()
sv_EVENT_SPEECH_STARTED = 1000
sv_EVENT_SPEECH_DONE = 1001
MSG_SV = 49937
MSG_SILENCE = 0x440
MSG_SPEAK = 0x441
variants = OrderedDict()
def v(id, x):
	variants[str(id)] = VoiceInfo(str(id), x)
v(0, 'Male')
v(1, 'Female')
v(2, 'Large Male')
v(3, 'Child')
v(4, 'Giant Male')
v(5, 'Mellow Female')
v(6, 'Mellow Male')
v(7, 'Crisp Male')
v(8, 'The Fly')
v(9, 'Robotoid')
v(10, 'Martian')
v(11, 'Colossus')
v(12, 'Fast Fred')
v(13, 'Old Woman')
v(14, 'Munchkin')
v(15, 'Troll')
v(16, 'Nerd')
v(17, 'Milktoast')
v(18, 'Tipsy')
v(19, 'Choirboy')

intstyles = OrderedDict()
def i(id, x):
	intstyles[str(id)] = VoiceInfo(str(id), x)
i(0, 'normal1')
i(1, 'normal2')
i(2, 'monotone')
i(3, 'sung')
i(4, 'random')

vmodes = OrderedDict()
def m(id, x):
	vmodes[str(id)] = VoiceInfo(str(id), x)
m(0, 'normal')
m(1, 'breathy')
m(2, 'whispered')

genders = OrderedDict()
def g(id, x):
	genders[str(id)] = VoiceInfo(str(id), x)
g(0, 'male')
g(1, 'female')
g(2, 'child')
g(3, 'giant')

glots = OrderedDict()
def t(id, x):
	glots[str(id)] = VoiceInfo(str(id), x)
t(0, 'default')
t(1, 'male')
t(2, 'female')
t(3, 'child')
t(4, 'high')
t(5, 'mellow')
t(6, 'impulse')
t(7, 'odd')
t(8, 'colossus')

smodes = OrderedDict()
def k(id, x):
	smodes[str(id)] = VoiceInfo(str(id), x)
k(0, 'natural')
k(1, 'word-at-a-time')
k(2, 'spelled')

def errcheck(res, func, args):
	if res != 0:
		raise RuntimeError("%s: code %d" % (func.__name__, res))
	return res

path = os.path.abspath(os.path.join(os.path.dirname(__file__), r"tibase32.dll"))
path2 = os.path.abspath(os.path.join(os.path.dirname(__file__), r"tieng32.dll"))
path3 = os.path.abspath(os.path.join(os.path.dirname(__file__), r"tispan32.dll"))

dll = windll[path]
dll2 = windll[path2]
dll3 = windll[path3]
dll.SVOpenSpeech.errcheck = errcheck

appInstance = windll.kernel32.GetModuleHandleW(None)
nvdaSvWndCls = WNDCLASSEXW()
nvdaSvWndCls.cbSize = sizeof(nvdaSvWndCls)
nvdaSvWndCls.hInstance = appInstance
nvdaSvWndCls.lpszClassName = u"nvdaSvWndCls"

speaking = False
lastindex = None
event = threading.Event()

class SynthDriver(synthDriverHandler.SynthDriver):
	name="sv"
	description = _("Softvoice")
	supportedSettings = (SynthDriver.RateSetting(), SynthDriver.VariantSetting(), SynthDriver.VoiceSetting(),
	SynthDriver.PitchSetting(), SynthDriver.InflectionSetting(), NumericDriverSetting("perturb", "Perturbation"), NumericDriverSetting("vfactor", "Vowel Factor"), NumericDriverSetting("avbias", "Voicing Gain"), NumericDriverSetting("afbias", "Frication Gain"), NumericDriverSetting("ahbias", "Aspiration Gain"), DriverSetting("intstyle", "Intonation Style"), DriverSetting("vmode", "Voicing Mode"), DriverSetting("gender", "Gender"), DriverSetting("glot", "Glottal Source"), DriverSetting("smode", "Speaking Mode")
	)
	speech = []
	lock = threading.Lock()
	availableVoices = OrderedDict((str(index+1), VoiceInfo(str(index+1),name)) for index,name in enumerate(("English", "Spanish")))
	@classmethod
	def check(cls):
		return True

	def __init__(self):
		super(synthDriverHandler.SynthDriver,self).__init__()
		global unit
		self.setup_wndproc()
		self._messageWindowClassAtom = windll.user32.RegisterClassExW(byref(nvdaSvWndCls))
		print("atom = %r" % self._messageWindowClassAtom)
		self._messageWindow = windll.user32.CreateWindowExW(0, self._messageWindowClassAtom, u"nvdaSvWndCls window", 0, 0, 0, 0, 0, None, None, appInstance, None)
		self.handle = c_int()
		self.speaking = False
		self.speech_list = deque()
		self.rate = 50
		self.pitch = 4
		self.inflection = 25
		self.perturb = 0
		self.vfactor = 20
		self.avbias = 45
		self.afbias = 45
		self.ahbias = 45
		self.voice = "1"
		self.variant = "0"
		self.intstyle = "0"
		self.vmode = "0"
		self.gender = "0"
		self.glot = "0"
		self.smode = "0"

	def open_synth(self, voice):
		if self.handle:
			dll.SVCloseSpeech(self.handle)
			self.handle = c_int()
		dll.SVOpenSpeech(byref(self.handle), self._messageWindow, 0, int(voice), 0)

	def speak(self, speechSequence):
		if len(speechSequence) == 1 and speechSequence[0].strip() == '':
			return
		textList = []
		index = None
		for item in speechSequence:
			if isinstance(item, str):
				textList.append(item)
			elif isinstance(item, IndexCommand):
				index = item.index
		text = u"".join(textList)
		text = text.encode('windows-1252', 'ignore')
		self.speech_list.append((text, index))
		windll.user32.PostMessageW(self._messageWindow, MSG_SPEAK, 0, 0)

	def cancel(self):
		windll.user32.SendMessageW(self._messageWindow, MSG_SILENCE, 0, 0)

	def _set_rate(self, rate):
		val = self._percentToParam(rate, 20, 500)
		self.synthrate = val
		dll.SVSetRate(self.handle, val)

	def _get_rate(self):
		return self._paramToPercent(self.synthrate, 20, 500)

	def _get_lastIndex(self):
		global lastindex
		return lastindex

	def terminate(self):
		dll.SVCloseSpeech(self.handle)
		windll.user32.DestroyWindow(self._messageWindow)
		windll.user32.UnregisterClassW(self._messageWindowClassAtom, None)

	def _get_voice(self):
		return self.curvoice

	def _set_voice(self, voice):
		self.curvoice = voice
		self.open_synth(int(voice))
		self.rate = self.rate

	def setup_wndproc(self):
		@WNDPROC
		def nvdaSvWndProc(hwnd, msg, wParam, lParam):
			global speaking
			if wParam in (1000, 1001, 1002):
				return self.handle_sv_message(wParam)
			elif msg == MSG_SILENCE:
				speaking = False
				dll.SVAbort(self.handle)
				self.speech_list.clear()
			elif msg == MSG_SPEAK and self.speech_list and not speaking:
				print("msg_speak")
				t = self.speech_list.popleft()
				print("speaking %r" % (t,))
				speaking = True
				self.sv_speak(t[0], t[1])
			return windll.user32.DefWindowProcW(hwnd,msg,wParam,lParam)
		nvdaSvWndCls.lpfnWndProc = nvdaSvWndProc

	def handle_sv_message(self, msg):
		global speaking
		print("svmsg %d" % msg)
		if msg == 1001 and speaking:
			speaking = False
			windll.user32.PostMessageW(self._messageWindow, MSG_SPEAK, 0, 0)
		return 1

	def sv_speak(self, text, index):
		global lastindex, speaking
		print("sv_speak %s" % text)
		print(dll.SVTTS(self.handle, text.strip(), 0, 0, self._messageWindow, 0, 0, 0))
		lastindex = index

	def _get_availableVariants(self):
		return variants

	def _get_variant(self):
		return self._variant

	def _set_variant(self, id):
		self._variant = id
		dll.SVSetPersonality(self.handle, int(id))
		self.rate = self.rate

	def _set_pitch(self, pitch):
		val = self._percentToParam(pitch, 10, 2000)
		self.synthpitch = val
		dll.SVSetPitch(self.handle, val)

	def _get_pitch(self):
		return self._paramToPercent(self.synthpitch, 10, 2000)

	def _set_inflection(self, inflection):
		val = self._percentToParam(inflection, 0, 500)
		self.synthinflection = val
		dll.SVSetF0Range(self.handle, val)

	def _get_inflection(self):
		return self._paramToPercent(self.synthinflection, 0, 500)

	def _set_perturb(self, perturb):
		val = self._percentToParam(perturb, 0, 500)
		self.synthperturb = val
		dll.SVSetF0Perturb(self.handle, val)

	def _get_perturb(self):
		return self._paramToPercent(self.synthperturb, 0, 500)

	def _set_vfactor(self, vfactor):
		val = self._percentToParam(vfactor, 0, 500)
		self.synthvfactor = val
		dll.SVSetVowelFactor(self.handle, val)

	def _get_vfactor(self):
		return self._paramToPercent(self.synthvfactor, 0, 500)

	def _set_avbias(self, avbias):
		val = self._percentToParam(avbias, -50, 50)
		self.synthavbias = val
		dll.SVSetAVBias(self.handle, val)

	def _get_avbias(self):
		return self._paramToPercent(self.synthavbias, -50, 50)

	def _set_afbias(self, afbias):
		val = self._percentToParam(afbias, -50, 50)
		self.synthafbias = val
		dll.SVSetAFBias(self.handle, val)

	def _get_afbias(self):
		return self._paramToPercent(self.synthafbias, -50, 50)

	def _set_ahbias(self, ahbias):
		val = self._percentToParam(ahbias, -50, 50)
		self.synthahbias = val
		dll.SVSetAHBias(self.handle, val)

	def _get_ahbias(self):
		return self._paramToPercent(self.synthahbias, -50, 50)

	def _get_availableIntstyles(self):
		return intstyles

	def _get_intstyle(self):
		return self._intstyle

	def _set_intstyle(self, id):
		self._intstyle = id
		dll.SVSetF0Style(self.handle, int(id))

	def _get_availableVmodes(self):
		return vmodes

	def _get_vmode(self):
		return self._vmode

	def _set_vmode(self, id):
		self._vmode = id
		dll.SVSetVoicingMode(self.handle, int(id))

	def _get_availableGenders(self):
		return genders

	def _get_gender(self):
		return self._gender

	def _set_gender(self, id):
		self._gender = id
		dll.SVSetGender(self.handle, int(id))

	def _get_availableGlots(self):
		return glots

	def _get_glot(self):
		return self._glot

	def _set_glot(self, id):
		self._glot = id
		dll.SVSetGlottalSource(self.handle, int(id))

	def _get_availableSmodes(self):
		return smodes

	def _get_smode(self):
		return self._smode

	def _set_smode(self, id):
		self._smode = id
		dll.SVSetSpeakingMode(self.handle, int(id))
