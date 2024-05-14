# A part of NonVisual Desktop Access (NVDA)
# This file is covered by the GNU General Public License.
# See the file COPYING for more details.
# Copyright (C) 2018-2023 NV Access Limited, Babbage B.V., Takuya Nishimoto

"""Default highlighterPlus based on GDI Plus."""
from typing import Optional, Tuple

from autoSettingsUtils.autoSettings import SupportedSettingType
from autoSettingsUtils.driverSetting import BooleanDriverSetting
import vision
from vision.constants import Context
from vision.util import getContextRect
from vision.visionHandlerExtensionPoints import EventExtensionPoints
from vision import providerBase
from windowUtils import CustomWindow
import wx
from gui.settingsDialogs import (
	AutoSettingsMixin,
	SettingsPanel,
	VisionProviderStateControl,
)
import api
from ctypes import byref, WinError
from ctypes.wintypes import COLORREF, MSG
import winUser
from logHandler import log
from mouseHandler import getTotalWidthAndHeightAndMinimumPosition
from locationHelper import RectLTWH
from collections import namedtuple
import threading
from winAPI.messageWindow import WindowMessage
import winGDI
import weakref
from colors import RGB
import core
import time

class HighlightStyle(
		namedtuple("HighlightStyle", ("color", "width", "style", "margin"))
):
	"""Represents the style of a highlight for a particular context.
	@ivar color: The color to use for the style
	@type color: L{RGB}
	@ivar width: The width of the lines to be drawn, in pixels.
		A higher width reduces the inner dimensions of the rectangle.
		Therefore, if you need to increase the outer dimensions of the rectangle,
		you need to increase the margin as well.
	@type width: int
	@ivar style: The style of the lines to be drawn;
		One of the C{winGDI.DashStyle*} enumeration constants.
	@type style: int
	@ivar margin: The number of pixels between the highlight's rectangle
		and the rectangle of the object to be highlighted.
		A higher margin stretches the highlight's rectangle.
		This value may also be negative.
	@type margin: int
	"""


BLUE = RGB(0x03, 0x36, 0xFF)
PINK = RGB(0xFF, 0x02, 0x66)
YELLOW = RGB(0xFF, 0xDE, 0x03)
thickness = 10
DASH_BLUE = HighlightStyle(BLUE, 5, winGDI.DashStyleDash, thickness)
SOLID_PINK = HighlightStyle(PINK, 5, winGDI.DashStyleSolid, thickness )
SOLID_BLUE = HighlightStyle(BLUE, 5, winGDI.DashStyleSolid, thickness)
SOLID_YELLOW = HighlightStyle(YELLOW, 2, winGDI.DashStyleSolid, thickness)


class HighlightWindow(CustomWindow):
	transparency = 0xff
	className = u"NVDAhighlighterPlus"
	windowName = u"NVDA highlighterPlus Window"
	windowStyle = winUser.WS_POPUP | winUser.WS_DISABLED
	extendedWindowStyle = (
		# Ensure that the window is on top of all other windows
		winUser.WS_EX_TOPMOST
		# A layered window ensures that L{transparentColor} will be considered transparent, when painted
		| winUser.WS_EX_LAYERED
		# Ensure that the window can't be activated when pressing alt+tab
		| winUser.WS_EX_NOACTIVATE
		# Make this a transparent window,
		# primarily for accessibility APIs to ignore this window when getting a window from a screen point
		| winUser.WS_EX_TRANSPARENT
	)
	transparentColor = 0  # Black

	@classmethod
	def _get__wClass(cls):
		wClass = super()._wClass
		wClass.style = winUser.CS_HREDRAW | winUser.CS_VREDRAW
		wClass.hbrBackground = winGDI.gdi32.CreateSolidBrush(COLORREF(cls.transparentColor))
		return wClass

	def updateLocationForDisplays(self):
		if vision._isDebug():
			log.debug("Updating NVDAhighlighterPlus window location for displays")
		displays = [wx.Display(i).GetGeometry() for i in range(wx.Display.GetCount())]
		screenWidth, screenHeight, minPos = getTotalWidthAndHeightAndMinimumPosition(displays)
		# Hack: Windows has a "feature" that will stop desktop shortcut hotkeys from working
		# when a window is full screen.
		# Removing one line of pixels from the bottom of the screen will fix this.
		left = minPos.x
		top = minPos.y
		width = screenWidth
		height = screenHeight - 1
		self.location = RectLTWH(left, top, width, height)
		winUser.user32.ShowWindow(self.handle, winUser.SW_HIDE)
		if not winUser.user32.SetWindowPos(
			self.handle,
			winUser.HWND_TOPMOST,
			left, top, width, height,
			winUser.SWP_NOACTIVATE
		):
			raise WinError()
		winUser.user32.ShowWindow(self.handle, winUser.SW_SHOWNA)

	def __init__(self, highlighterPlus):
		if vision._isDebug():
			log.debug("initializing NVDAhighlighterPlus window")
		super().__init__(
			windowName=self.windowName,
			windowStyle=self.windowStyle,
			extendedWindowStyle=self.extendedWindowStyle
		)
		self.location = None
		self.highlighterPlusRef = weakref.ref(highlighterPlus)
		winUser.SetLayeredWindowAttributes(
			self.handle,
			self.transparentColor,
			self.transparency,
			winUser.LWA_ALPHA | winUser.LWA_COLORKEY)
		self.updateLocationForDisplays()
		if not winUser.user32.UpdateWindow(self.handle):
			raise WinError()

	def windowProc(self, hwnd, msg, wParam, lParam):
		if msg == winUser.WM_PAINT:
			self._paint()
			# Ensure the window is top most
			# winUser.user32.SetWindowPos(
			# 	self.handle,
			# 	winUser.HWND_TOPMOST,
			# 	0, 0, 0, 0,
			# 	winUser.SWP_NOACTIVATE | winUser.SWP_NOMOVE | winUser.SWP_NOSIZE
			# )
		elif msg == winUser.WM_DESTROY:
			winUser.user32.PostQuitMessage(0)
		elif msg == winUser.WM_TIMER:
			self.refresh()
		elif msg == WindowMessage.DISPLAY_CHANGE:
			# wx might not be aware of the display change at this point
			core.callLater(100, self.updateLocationForDisplays)

	def _paint(self):
		highlighterPlus = self.highlighterPlusRef()
		if not highlighterPlus:
			# The highlighterPlus instance died unexpectedly, kill the window as well
			winUser.user32.PostQuitMessage(0)
			return
		contextRects = {}
		for context in highlighterPlus.enabledContexts:
			rect = highlighterPlus.contextToRectMap.get(context)
			if not rect:
				continue
			elif context == Context.NAVIGATOR and contextRects.get(Context.FOCUS) == rect:
				# When the focus overlaps the navigator object, which is usually the case,
				# show a different highlight style.
				# Focus is in contextRects, do not show the standalone focus highlight.
				contextRects.pop(Context.FOCUS)
				# Navigator object might be in contextRects as well
				contextRects.pop(Context.NAVIGATOR, None)
				context = Context.FOCUS_NAVIGATOR
			contextRects[context] = rect
		if not contextRects:
			return
		with winUser.paint(self.handle) as hdc:
			with winGDI.GDIPlusGraphicsContext(hdc) as graphicsContext:
				for context, rect in contextRects.items():
					HighlightStyle = highlighterPlus._ContextStyles[context]
					# Before calculating logical coordinates,
					# make sure the rectangle falls within the highlighterPlus window
					rect = rect.intersection(self.location)
					try:
						rect = rect.toLogical(self.handle)
					except RuntimeError:
						log.debugWarning("", exc_info=True)
					rect = rect.toClient(self.handle)
					try:
						rect = rect.expandOrShrink(HighlightStyle.margin)
					except RuntimeError:
						pass
					with winGDI.GDIPlusPen(
						HighlightStyle.color.toGDIPlusARGB(),
						HighlightStyle.width,
						HighlightStyle.style
					) as pen:
						winGDI.gdiPlusDrawRectangle(graphicsContext, pen, *rect.toLTWH())

	def refresh(self):
		winUser.user32.InvalidateRect(self.handle, None, True)


_contextOptionLabelsWithAccelerators = {
	# Translators: shown for a highlighterPlus setting that toggles
	# highlighting the system focus.
	Context.FOCUS: _("Highlight system fo&cus"),
	# Translators: shown for a highlighterPlus setting that toggles
	# highlighting the browse mode cursor.
	Context.BROWSEMODE: _("Highlight browse &mode cursor"),
	# Translators: shown for a highlighterPlus setting that toggles
	# highlighting the navigator object.
	Context.NAVIGATOR: _("Highlight navigator &object"),
}

_supportedContexts = (Context.FOCUS, Context.NAVIGATOR, Context.BROWSEMODE)


class NVDAhighlighterPlusSettings(providerBase.VisionEnhancementProviderSettings):
	# Default settings for parameters
	highlightPlusFocus = False
	highlightPlusNavigator = False
	highlightPlusBrowseMode = False

	@classmethod
	def getId(cls) -> str:
		return "NVDAhighlighterPlus"

	@classmethod
	def getDisplayName(cls) -> str:
		# Translators: Description for NVDA's built-in screen highlighterPlus.
		return _("Visual Highlight Plus")

	def _get_supportedSettings(self) -> SupportedSettingType:
		return [
			BooleanDriverSetting(
				'highlightPlus%s' % (context[0].upper() + context[1:]),
				_contextOptionLabelsWithAccelerators[context],
				defaultVal=True
			)
			for context in _supportedContexts
		]


class NVDAhighlighterPlusGuiPanel(
		AutoSettingsMixin,
		SettingsPanel
):

	_enableCheckSizer: wx.BoxSizer
	_enabledCheckbox: wx.CheckBox
	
	helpId = "VisionSettingsFocusHighlight"

	def __init__(
			self,
			parent: wx.Window,
			providerControl: VisionProviderStateControl
	):
		self._providerControl = providerControl
		initiallyEnabledInConfig = NVDAhighlighterPlus.isEnabledInConfig()
		if not initiallyEnabledInConfig:
			settingsStorage = self._getSettingsStorage()
			settingsToCheck = [
				settingsStorage.highlightPlusBrowseMode,
				settingsStorage.highlightPlusFocus,
				settingsStorage.highlightPlusNavigator,
			]
			if any(settingsToCheck):
				log.debugWarning(
					"highlighterPlus disabled in config while some of its settings are enabled. "
					"This will be corrected"
				)
				settingsStorage.highlightPlusBrowseMode = False
				settingsStorage.highlightPlusFocus = False
				settingsStorage.highlightPlusNavigator = False
		super().__init__(parent)

	def _buildGui(self):
		self.mainSizer = wx.BoxSizer(wx.VERTICAL)

		self._enabledCheckbox = wx.CheckBox(
			self,
			#  Translators: The label for a checkbox that enables / disables focus highlighting
			#  in the NVDA highlighterPlus vision settings panel.
			label=_("&Enable Highlighting"),
			style=wx.CHK_3STATE
		)

		self.mainSizer.Add(self._enabledCheckbox)
		self.mainSizer.AddSpacer(size=self.scaleSize(10))
		# this options separator is done with text rather than a group box because a groupbox is too verbose,
		# but visually some separation is helpful, since the rest of the options are really sub-settings.
		self.optionsText = wx.StaticText(
			self,
			# Translators: The label for a group box containing the NVDA highlighterPlus options.
			label=_("Options:")
		)
		self.mainSizer.Add(self.optionsText)

		self.lastControl = self.optionsText
		self.settingsSizer = wx.BoxSizer(wx.VERTICAL)
		self.makeSettings(self.settingsSizer)
		self.mainSizer.Add(self.settingsSizer, border=self.scaleSize(15), flag=wx.LEFT | wx.EXPAND)
		self.mainSizer.Fit(self)
		self.SetSizer(self.mainSizer)

	def getSettings(self) -> NVDAhighlighterPlusSettings:
		# AutoSettingsMixin uses the getSettings method (via getSettingsStorage) to get the instance which is
		# used to get / set attributes. The attributes must match the id's of the settings.
		# We want them set on our settings instance.
		return VisionEnhancementProvider.getSettings()

	def makeSettings(self, sizer: wx.BoxSizer):
		self.updateDriverSettings()
		# bind to all check box events
		self.Bind(wx.EVT_CHECKBOX, self._onCheckEvent)
		self._updateEnabledState()

	def onPanelActivated(self):
		self.lastControl = self.optionsText

	def _updateEnabledState(self):
		settingsStorage = self._getSettingsStorage()
		settingsToTriggerActivation = [
			settingsStorage.highlightPlusBrowseMode,
			settingsStorage.highlightPlusFocus,
			settingsStorage.highlightPlusNavigator,
		]
		isAnyEnabled = any(settingsToTriggerActivation)
		if all(settingsToTriggerActivation):
			self._enabledCheckbox.Set3StateValue(wx.CHK_CHECKED)
		elif isAnyEnabled:
			self._enabledCheckbox.Set3StateValue(wx.CHK_UNDETERMINED)
		else:   
			self._enabledCheckbox.Set3StateValue(wx.CHK_UNCHECKED)

		if not self._ensureEnableState(isAnyEnabled) and isAnyEnabled:
			self._onEnableFailure()

	def _onEnableFailure(self):
		""" Initialization of highlighterPlus failed. Reset settings / GUI
		"""
		settingsStorage = self._getSettingsStorage()
		settingsStorage.highlightPlusBrowseMode = False
		settingsStorage.highlightPlusFocus = False
		settingsStorage.highlightPlusNavigator = False
		self.updateDriverSettings()
		self._updateEnabledState()

	def _ensureEnableState(self, shouldBeEnabled: bool) -> bool:
		currentlyEnabled = bool(self._providerControl.getProviderInstance())
		if shouldBeEnabled and not currentlyEnabled:
			return self._providerControl.startProvider()
		elif not shouldBeEnabled and currentlyEnabled:
			return self._providerControl.terminateProvider()
		return True

	def _onCheckEvent(self, evt: wx.CommandEvent):
		settingsStorage = self._getSettingsStorage()
		if evt.GetEventObject() is self._enabledCheckbox:
			isEnableAllChecked = evt.IsChecked()
			settingsStorage.highlightPlusBrowseMode = isEnableAllChecked
			settingsStorage.highlightPlusFocus = isEnableAllChecked
			settingsStorage.highlightNavigator = isEnableAllChecked
			if not self._ensureEnableState(isEnableAllChecked) and isEnableAllChecked:
				self._onEnableFailure()
			self.updateDriverSettings()
		else:
			self._updateEnabledState()

		providerInst: Optional[NVDAhighlighterPlus] = self._providerControl.getProviderInstance()
		if providerInst:
			providerInst.refresh()


class NVDAhighlighterPlus(providerBase.VisionEnhancementProvider):
	_ContextStyles = {
		Context.FOCUS: DASH_BLUE,
		Context.NAVIGATOR: SOLID_PINK,
		Context.FOCUS_NAVIGATOR: SOLID_BLUE,
		Context.BROWSEMODE: SOLID_YELLOW,
	}
	_refreshInterval = 100
	customWindowClass = HighlightWindow
	_settings = NVDAhighlighterPlusSettings()
	_window: Optional[customWindowClass] = None
	enabledContexts: Tuple[Context]  # type info for autoprop: L{_get_enableContexts}

	@classmethod  # override
	def getSettings(cls) -> NVDAhighlighterPlusSettings:
		return cls._settings

	@classmethod  # override
	def getSettingsPanelClass(cls):
		"""Returns the class to be used in order to construct a settings panel for the provider.
		@return: Optional[SettingsPanel]
		@remarks: When None is returned, L{gui.settingsDialogs.VisionProviderSubPanel_Wrapper} is used.
		"""
		return NVDAhighlighterPlusGuiPanel

	@classmethod  # override
	def canStart(cls) -> bool:
		return True

	def registerEventExtensionPoints(  # override
			self,
			extensionPoints: EventExtensionPoints
	) -> None:
		extensionPoints.post_focusChange.register(self.handleFocusChange)
		extensionPoints.post_reviewMove.register(self.handleReviewMove)
		extensionPoints.post_browseModeMove.register(self.handleBrowseModeMove)
		extensionPoints.post_coreCycle.register(self.handleUpdateThread)
		extensionPoints.post_objectUpdate.register(self.handleUpdateThread)
		extensionPoints.post_coreCycleSecond.register(self.handleUpdateAllContexts)

	def __init__(self):
		super().__init__()
		log.debug("Starting NVDAhighlighterPlus")
		self.contextToRectMap = {}
		winGDI.gdiPlusInitialize()
		self._highlighterPlusThread = threading.Thread(
			name=f"{self.__class__.__module__}.{self.__class__.__qualname__}",
			target=self._run
		)
		self._highlighterPlusRunningEvent = threading.Event()
		self._highlighterPlusThread.daemon = True
		self._highlighterPlusThread.start()

		# Update
		# Make sure the highlighterPlus thread doesn't exit early.
		waitResult = self._highlighterPlusRunningEvent.wait(0.2)
		if waitResult is False or not self._highlighterPlusThread.is_alive():
			raise RuntimeError("highlighterPlus thread wasn't able to initialize correctly")

	def terminate(self):
		log.debug("Terminating NVDAhighlighterPlus")
		if self._highlighterPlusThread and self._window and self._window.handle:
			if not winUser.user32.PostThreadMessageW(self._highlighterPlusThread.ident, winUser.WM_QUIT, 0, 0):
				raise WinError()
			else:
				self._highlighterPlusThread.join()
			self._highlighterPlusThread = None
		winGDI.gdiPlusTerminate()
		self.contextToRectMap.clear()
		super().terminate()

	def _run(self):
		try:
			if vision._isDebug():
				log.debug("Starting NVDAhighlighterPlus thread")

			window = self._window = self.customWindowClass(self)
			timer = winUser.WinTimer(window.handle, 0, self._refreshInterval, None)
			self._highlighterPlusRunningEvent.set()  # notify main thread that initialisation was successful
			msg = MSG()
			while (res := winUser.getMessage(byref(msg), None, 0, 0)) > 0:
				winUser.user32.TranslateMessage(byref(msg))
				winUser.user32.DispatchMessageW(byref(msg))
			if res == -1:
				# See the return value section of
				# https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getmessage
				raise WinError()
			if vision._isDebug():
				log.debug("Quit message received on NVDAhighlighterPlus thread")
			timer.terminate()
			window.destroy()
		except Exception:
			log.exception("Exception in NVDA highlighterPlus thread")

	def handleUpdateThread(self):
		log.info("handle update thread corecycle")
		self.handleBrowseModeMove()
		self.handleReviewMove()

	def handleUpdateAllContexts(self):
		time.sleep(0.1)
		log.info("handle update thread second core cycle")
		self.handleBrowseModeMove()
		self.handleReviewMove()
   
	def updateContextRect(self, context, rect=None, obj=None):
		"""Updates the position rectangle of the highlight for the specified context.
		If rect is specified, the method directly writes the rectangle to the contextToRectMap.
		Otherwise, it will call L{getContextRect}
		"""
		if context not in self.enabledContexts:
			return
		if rect is None:
			try:
				rect = getContextRect(context, obj=obj)
			except (LookupError, NotImplementedError, RuntimeError, TypeError):
				rect = None
		self.contextToRectMap[context] = rect

	def handleFocusChange(self, obj):
		self.updateContextRect(context=Context.FOCUS, obj=obj)
		if not api.isObjectInActiveTreeInterceptor(obj):
			self.contextToRectMap.pop(Context.BROWSEMODE, None)
		else:
			self.handleBrowseModeMove()

	def handleReviewMove(self):
		self.updateContextRect(context=Context.NAVIGATOR)

	def handleBrowseModeMove(self):
		self.updateContextRect(context=Context.BROWSEMODE)

	def refresh(self):
		"""Refreshes the screen positions of the enabled highlights.
		"""
		if self._window and self._window.handle:
			self._window.refresh()

	def _get_enabledContexts(self):
		"""Gets the contexts for which the highlighterPlus is enabled.
		"""
		return tuple(
			context for context in _supportedContexts
			if getattr(self.getSettings(), 'highlightPlus%s' % (context[0].upper() + context[1:]))
		)

VisionEnhancementProvider = NVDAhighlighterPlus
