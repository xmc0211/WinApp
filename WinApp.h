// MIT License
//
// Copyright (c) 2025 WinApp - xmc0211 <xmc0211@qq.com>
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.
// 
// References:
// - https://github.com/mohabouje/WinToast 
//     MIT License, Copyright (C) 2016-2023 WinToast v1.3.0 - Mohammed Boujemaoui <mohabouje@gmail.com>
// Thanks for them.

#ifndef WINAPP_H
#define WINAPP_H

#include <Windows.h>
#include <string>
#include <vector>
#include <map>
#include <winstring.h>
#include <wrl.h>
#include <ShlObj.h>
#include <ShObjIdl.h>
#include <wrl/event.h>
#include <wrl/implements.h>
#include <windows.ui.Notifications.h>

#if defined(UNICODE) 
#define _tstring wstring
#else
#define _tstring string
#endif

// Errors
#define E_NOT_SUPPORT (MAKE_HRESULT(1, FACILITY_ITF, 0x1000))
#define S_SKIPPED (MAKE_HRESULT(0, FACILITY_ITF, 0x1000))

// Namespaces
using namespace Microsoft::WRL;
using namespace ABI::Windows::Data::Xml::Dom;
using namespace ABI::Windows::UI::Notifications;
using namespace ABI::Windows::Foundation;
using namespace Windows::Foundation;

// Declarations of class
class WinApp;
class WinToast;
class WinTaskbarIcon;
class TaskbarButton;
class JumpListShelllink;
class JumpListCollection;

/* ==================================== WinApp Begin ==================================== */

#define WINAPP_SYSTEM_TIMEOUT (3500)
#define WINAPP_READY_WAIT_TIMEOUT (50)

enum WinAppRegisterShortcutPolicy {
	// Don't check, create, or modify a shortcut.
	ShortcutPolicy_Ignore = 0,
	// Require a shortcut with matching AUMI, don't create or modify an existing one.
	ShortcutPolicy_RequireNoCreate = 1,
	// Require a shortcut with matching AUMI, create if missing, modify if not matching. This is the default.
	ShortcutPolicy_RequireCreate = 2,
};

// WinApp class
class WinApp {
private:
	std::_tstring _AppName;
	std::_tstring _AppID;
	BOOL _bRegistered, _bIsReady;
	WinAppRegisterShortcutPolicy _ShortcutPolicy;

protected:

public:
	WinApp();
	~WinApp();

	// Set the name of your application.
	HRESULT SetAppName(std::_tstring AppName);
	// Set the ID of your application. Recommended format: "Company.Product.SubProduct.Version".
	HRESULT SetAppID(std::_tstring AppID);
	HRESULT SetAppID(std::_tstring CompanyName, std::_tstring ProductName, std::_tstring SubProductName, std::_tstring VersionInfo);
	// Set the policy for creating start menu shortcuts.
	HRESULT SetShorcutPolicy(WinAppRegisterShortcutPolicy Policy);

	// Register your application for this process.
	// WARNING: When a new app is registered in the system, it cannot immediately send a Toast notification, 
	// you should leaving the system with about 3.5 seconds of processing time. Otherwise, the notification 
	// will not be displayed or there may be cause error.
	// Set parameter 'bWaitForSystemIfAppNotExist' to 'FALSE' to disable 'wait for 3.5 seconds automatically'.
	HRESULT RegisterApp(BOOL bWaitForSystemIfAppNotExist = TRUE);

	// Get Value
	std::_tstring GetAppName();
	std::_tstring GetAppID();
	WinAppRegisterShortcutPolicy GetShortcutPolicy();
	BOOL IsRegistered();
	BOOL IsReady();

	// Helper
	HRESULT ValidateShellLink(BOOL& bWasChanged);
	HRESULT CreateShellLink();
};

/* ===================================== WinApp End ===================================== */





/* =================================== WinToast Begin =================================== */

// Dismissal Reasons
enum WinToastDismissalReason {
	// The user moved the notification to the notification center.
	DismissReason_UserCanceled = ToastDismissalReason::ToastDismissalReason_UserCanceled,
	// The user has set to hide all notifications for your registered application, so the notifications cannot be displayed.
	DismissReason_ApplicationHidden = ToastDismissalReason::ToastDismissalReason_ApplicationHidden,
	// Notification has not been operated within the time frame you have set.
	DismissReason_Timedout = ToastDismissalReason::ToastDismissalReason_TimedOut,
};

// Audio Options
enum WinToastAudioOption {
	// Default (play once)
	AudioOption_Default = 0,
	// No sound is emitted, including the default notification sound of the system.
	AudioOption_Silent,
	// Repeatedly play the sound you specified.
	AudioOption_Loop,
};

// Types
enum WinToastType {
	// A large image and a single string wrapped across thResee lines of text.
	ToastType_ImageAndText01 = ToastTemplateType::ToastTemplateType_ToastImageAndText01,
	// A large image, one string of bold text on the first line, one string of regular text wrapped across the second and third lines.
	ToastType_ImageAndText02 = ToastTemplateType::ToastTemplateType_ToastImageAndText02,
	// A large image, one string of bold text wrapped across the first two lines, one string of regular text on the third line.
	ToastType_ImageAndText03 = ToastTemplateType::ToastTemplateType_ToastImageAndText03,
	// A large image, one string of bold text on the first line, one string of regular text on the second line, one string of regular text on the third line.
	ToastType_ImageAndText04 = ToastTemplateType::ToastTemplateType_ToastImageAndText04,
	// Single string wrapped across thResee lines of text.
	ToastType_Text01 = ToastTemplateType::ToastTemplateType_ToastText01,
	// One string of bold text on the first line, one string of regular text wrapped across the second and third lines.
	ToastType_Text02 = ToastTemplateType::ToastTemplateType_ToastText02,
	// One string of bold text wrapped across the first two lines, one string of regular text on the third line.
	ToastType_Text03 = ToastTemplateType::ToastTemplateType_ToastText03,
	// One string of bold text on the first line, one string of regular text on the second line, one string of regular text on the third line.
	ToastType_Text04 = ToastTemplateType::ToastTemplateType_ToastText04
};

// Duration Options
enum WinToastDuration {
	// System default
	Duration_System,
	// Short time
	Duration_Short,
	// Long time
	Duration_Long
};

// Small Image Shapes
enum WinToastCropHint {
	// Square
	CropHint_Square,
	// Circle
	CropHint_Circle,
};

// Options used by SetTextField
enum WinToastTextField {
	// Set the first line.
	TextField_FirstLine = 0,
	// Set the second line.
	TextField_SecondLine,
	// Set the third line.
	TextField_ThirdLine
};

// System Audio Files (Located at %WINDIR%\\Media\\)
enum WinToastAudioSystemFile {
	DefaultSound,
	IM,
	Mail,
	Reminder,
	SMS,
	Alarm,
	Alarm2,
	Alarm3,
	Alarm4,
	Alarm5,
	Alarm6,
	Alarm7,
	Alarm8,
	Alarm9,
	Alarm10,
	Call,
	Call1,
	Call2,
	Call3,
	Call4,
	Call5,
	Call6,
	Call7,
	Call8,
	Call9,
	Call10,
};

// Scenario Options
enum WinToastScenario {
	// Default
	Scenario_Default,
	// Alarm
	Scenario_Alarm,
	// IncomingCall
	Scenario_IncomingCall,
	// Reminder
	Scenario_Reminder
};

// Callback function types
typedef void(CALLBACK* WinToast_Clicked)();
typedef void(CALLBACK* WinToast_ActionSelected)(INT ActionIndex);
typedef void(CALLBACK* WinToast_Replied)(std::_tstring Response);
typedef void(CALLBACK* WinToast_Dismissed)(WinToastDismissalReason State);
typedef void(CALLBACK* WinToast_Failed)();

// WinToast class
// You can destroy the WinToast instance you are using before notifying the callback function call.
class WinToast {
private:
	WinApp* _App;

	BOOL _bHasInput;
	std::vector<std::_tstring> _TextFields;
	std::vector<std::_tstring> _Actions;
	std::_tstring _ImagePath;
	std::_tstring _HeroImagePath;
	BOOL _InlineHeroImage;
	std::_tstring _AudioPath;
	std::_tstring _AttributionText;
	std::_tstring _InputReplyText;
	std::_tstring _InputBackgroundText;
	std::_tstring _Scenario;
	INT64 _Expiration;
	WinToastAudioOption _AudioOption;
	WinToastType _Type;
	WinToastDuration _Duration;
	WinToastCropHint _CropHint;

protected:

	void MarkAsReadyForDeletion(INT64 id);

	struct NotifyData {
	public:
		NotifyData();
		NotifyData(ComPtr<IToastNotification> Notify, EventRegistrationToken ActivatedToken,
			EventRegistrationToken DismissedToken, EventRegistrationToken FailedToken);
		~NotifyData();
		void RemoveTokens();
		void MarkAsReadyForDeletion();
		BOOL IsReadyForDeletion() const;
		IToastNotification* Notification();

	private:
		ComPtr<IToastNotification> Notify;
		EventRegistrationToken ActivatedToken;
		EventRegistrationToken DismissedToken;
		EventRegistrationToken FailedToken;
		BOOL ReadyForDeletion;
		BOOL PreviouslyTokenRemoved;
	};

	std::map<INT64, NotifyData> _Buffer;

public:
	// Due to functional requirements, you MUST specify a valid WinApp pointer.
	WinToast(WinApp* App, WinToastType Type = WinToastType::ToastType_ImageAndText02);
	~WinToast();

	// Get parameters
	std::size_t TextFieldsCount() const;
	std::size_t ActionsCount() const;
	BOOL HasImage() const;
	BOOL HasHeroImage() const;
	std::vector<std::_tstring> const& GetTextFields() const;
	std::_tstring const& GetTextField(WinToastTextField Pos) const;
	std::_tstring const& GetActionLabel(std::size_t Pos) const;
	std::_tstring const& GetImagePath() const;
	std::_tstring const& GetHeroImagePath() const;
	std::_tstring const& GetAudioPath() const;
	std::_tstring const& GetAttributionText() const;
	std::_tstring const& GetScenario() const;
	std::_tstring const& GetInputBackgroundText() const;
	std::_tstring const& GetInputReplyText() const;
	INT64 GetExpiration() const;
	WinToastType GetType() const;
	WinToastAudioOption GetAudioOption() const;
	WinToastDuration GetDuration() const;
	BOOL IsToastGeneric() const;
	BOOL IsInlineHeroImage() const;
	BOOL IsCropHintCircle() const;
	BOOL IsInput() const;

	// Set parameters related to notifications
	void SetFirstLine(std::_tstring const& Text);
	void SetSecondLine(std::_tstring const& Text);
	void SetThirdLine(std::_tstring const& Text);
	void SetTextField(std::_tstring const& Text, WinToastTextField Pos);
	void SetAttributionText(std::_tstring const& AttributionText);
	void SetImagePath(std::_tstring const& ImagePath, WinToastCropHint CropHint = WinToastCropHint::CropHint_Square);
	void SetHeroImagePath(std::_tstring const& ImagePath, BOOL InlineImage = FALSE);
	void SetAudioPath(WinToastAudioSystemFile Audio);
	void SetAudioPath(std::_tstring const& AudioPath);
	void SetAudioOption(WinToastAudioOption AudioOption);
	void SetDuration(WinToastDuration Duration);
	void SetExpiration(INT64 MillisecondsFromNow);
	void SetScenario(WinToastScenario Scenario);
	void AddAction(std::_tstring const& Label);
	void AddInput(std::_tstring const& BackgroundText = TEXT("..."), std::_tstring const& ReplyText = TEXT("Reply"));

	void SetCallbackClicked(WinToast_Clicked Func);
	void SetCallbackActionSelected(WinToast_ActionSelected Func);
	void SetCallbackReplied(WinToast_Replied Func);
	void SetCallbackDismissed(WinToast_Dismissed Func);
	void SetCallbackFailed(WinToast_Failed Func);

	// Show this toast
	INT64 Show(INT64* pToastID = NULL);
	
};

/* ==================================== WinToast End ==================================== */





/* ================================ WinTaskbarIcon Begin ================================ */

#define TASKBAR_FLASH_INFINITY (0xFFFFFFFF)
#define INVALID_TOOLBARBUTTON_INDEX (-1)
#define MAX_TOOLBARBUTTON (7)

// The status of taskbar icons
enum WinTaskbarIcon_State {
	// No effect at all
	TaskbarState_Normal, 
	// Error. Display the progress you have set on the taskbar with a red background.
	TaskbarState_Error,
	// Paused. Display the progress you have set on the taskbar with a yellow background.
	TaskbarState_Paused,
	// Processing. Display the progress you have set on the taskbar with a green background.
	TaskbarState_Progress,
	// Processing. Display the indeterminate progress on the taskbar with a green background.
	TaskbarState_Indeterminate,
	// Flash. Flash the taskbar icon of your window with an orange background to attract user attention.
	TaskbarState_Flash, 
};

// Flash Flags
enum WinTaskbarIcon_FlashFlag : DWORD {
	// When the user hovers the mouse over a taskbar icon, a border appears displaying a thumbnail of the window inside the border. This border is referred to as Window Taskbar-border.

	// Only Taskbar-border flashes
	TaskbarFlash_Caption = 0x01,
	// Only Taskbar Icon flashes
	TaskbarFlash_Tray = 0x02,
	// Taskbar-border and Taskbar Icon are both flashing
	TaskbarFlash_All = TaskbarFlash_Caption | TaskbarFlash_Tray,
};

// Taskbar Button Flags (Toolbar Buttons)
enum WinTaskbarButton_Flag : UINT {
	// Enable the button. (Default)
	TaskbarButton_Enabled = THBF_ENABLED,
	// Disable the button. The button will turn gray and cannot be clicked.
	TaskbarButton_Disabled = THBF_DISABLED,
	// After clicking the settings button, the entire preview screen disappears. The button will not be affected.
	TaskbarButton_DismissonClick = THBF_DISMISSONCLICK,
	// Do not draw borders when setting display buttons.
	TaskbarButton_NoBackground = THBF_NOBACKGROUND,
	// Hide this button. (If not re displayed, the effect is equivalent to deletion)
	TaskbarButton_Hidden = THBF_HIDDEN,
	// Disable button. The button will not turn gray, but it cannot be clicked.
	TaskbarButton_NoActive = THBF_NONINTERACTIVE,
};

// WinTaskbarIcon class
class WinTaskbarIcon {

private:
	HWND _hWnd;
	ComPtr<ITaskbarList3> _pTaskbar;
	ComPtr<ICustomDestinationList> _pJumpList;

	ULONGLONG _nStateProgressValue;
	ULONGLONG _nStateProgressTotal;
	WinTaskbarIcon_State _StateNow;

	UINT _nFlashCount;
	DWORD _dwFlashFlags;
	DWORD _dwFlashTimeout;

	UINT _nJumpListMaxSlots;
	BOOL _bJumpListCanAdd;

protected:
	BOOL IsProgressState();

public:
	// You can specify a valid WinApp pointer at the InApp parameter or fill in NULL.
	WinTaskbarIcon(WinApp* pApp, HWND hWnd);
	~WinTaskbarIcon();

	/* ----------- Taskbar Icon State ----------- */

	// Set the current progress value and total progress value.
	// If the overall progress value is ignored or filled in as -1, the value set for the first time will remain unchanged. 
	// The default value is 100.
	HRESULT SetProgressValue(ULONGLONG nProgressValue, ULONGLONG nProgressTotal = -1);
	// Set the time interval and number of times the taskbar icon flashes. 
	// After flashing for more than the specified number of times, the orange light will remain on continuously. 
	// If TASKBAR_FLASH_INFINITY is ignored or filled in, it will flash continuously.
	HRESULT SetFlashValue(DWORD dwTimeout = 500, UINT nCount = TASKBAR_FLASH_INFINITY);
	// Set Flash Flags. 
	// The result of the previous setting will be overwritten.
	HRESULT SetFlashFlag(DWORD dwFlags);
	// Set taskbar icon state. 
	// The result of the last setting will be saved.
	HRESULT SetState(WinTaskbarIcon_State State);

	/* ------------- Taskbar Button ------------- */

	// TaskbarButton class (Taskbar button)
	class TaskbarButton {

	private:
		THUMBBUTTON _Button;

	protected:

	public:
		TaskbarButton();
		TaskbarButton(UINT nButtonID);
		TaskbarButton(const TaskbarButton& other);
		~TaskbarButton();

		// Set Button Flags. 
		// The result of the last setting will be saved.
		HRESULT AddFlag(WinTaskbarButton_Flag Flag);
		// Remove Button Flags. 
		// The result of the last setting will be saved.
		HRESULT RemoveFlag(WinTaskbarButton_Flag Flag);
		// Set button icon. 
		// You can destroy the IOCON passed to the function before adding the button to the window after calling the function.
		HRESULT SetImage(HICON hIcon);
		// Set the prompt text displayed when the mouse hovers over the button. 
		// If filling in the blank string, there will be no prompt text.
		HRESULT SetTip(std::_tstring Tip);
		// Set the ID of the button. 
		// You need to treat it as a button in the window and process the message when the button is pressed in the WM_COMMAND message.
		HRESULT SetButtonID(UINT nID);

		// Get the button object
		THUMBBUTTON* GetButtonObject();

	};

	// Add 1 to 7 buttons. 
	// pButton is the first address of the TaskbarButton type array. 
	// The default nCount is 1.
	// You must add all the buttons at once.
	HRESULT AddButton(TaskbarButton* pButton, UINT nCount = 1);
	// Update 1 to 7 buttons. 
	// If the given button ID is already displayed, it will overwrite the previous setting.
	HRESULT UpdateButton(TaskbarButton* pButton, UINT nCount = 1);

	/* ------------ Taskbar JumpList ------------ */

	// JumpListShelllink class
	class JumpListShelllink {

	private:
		ComPtr<IShellLink> _pShelllink;
		std::_tstring _Name;

	protected:

	public:
		JumpListShelllink();
		~JumpListShelllink();

		// Set Arguments
		HRESULT SetName(std::_tstring Name);
		HRESULT SetFilePath(std::_tstring FilePath);
		HRESULT SetArguments(std::_tstring Arguments);
		HRESULT SetWorkingDirectory(std::_tstring WorkingDirectory);
		HRESULT SetIconPath(std::_tstring IconPath, INT nIndex = 0);
		HRESULT SetDescription(std::_tstring Description);
		HRESULT SetShowCmd(UINT nShowCmd);

		ComPtr<IShellLink> GetShelllinkObject();
		
	};

	// JumpListCollection class
	class JumpListCollection {

	private:
		ComPtr<IObjectCollection> _pCollection;
		std::_tstring _Name;

	protected:

	public:
		JumpListCollection();
		~JumpListCollection();

		// Set collection name
		// To avoid conflicts with the system collection, please avoid using the following names:
		// Tasks, Frequent, Recent, Pin, Properties.
		HRESULT SetName(std::_tstring Name);

		// Provide an JumpListShelllink pointer and add it to this collection.
		HRESULT AddShelllink(JumpListShelllink* pShelllink);

		std::_tstring GetName();
		ComPtr<IObjectCollection> GetCollectionObject();
	};

	// How many Jump List Collections does the system support at most
	UINT GetMaxSlots();

	// NOTE: You MUST first call ResetJumpList, then call AddJumpList several times
	// and finally call CommitJumpList apply Jump List. Calling AddJumpList in other locations 
	// is invalid.

	// Reset and delete all Jump List items (Except for System Jump Lists)
	HRESULT ResetJumpList();

	// Specify JumpListCollection pointer and add the collection to Jump List
	HRESULT AddJumpList(JumpListCollection* pCollection);
	HRESULT SetUserTaskJumpList(JumpListCollection* pCollection); // Add the collection into default user task. (Ignore the name of the Collection)

	// Commit Jump List
	HRESULT CommitJumpList();

	/* ----------------- Other ------------------ */

	// Show or hide the icon of the window on the taskbar.
	HRESULT ShowTaskbarIcon(BOOL bShow);

	// Set the overlay icon and hint text.
	// If hIcon is NULL, overlay icon will not be displayed. If InHint is empty, it will not be displayed.
	HRESULT SetOverlayIcon(HICON hIcon, std::_tstring Hint);

};

/* ================================= WinTaskbarIcon End ================================= */

#endif