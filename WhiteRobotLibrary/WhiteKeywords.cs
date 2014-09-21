using System;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Diagnostics;
using System.Windows.Automation;
using System.Windows.Automation.Peers;
using System.Collections.Generic;
using TestStack.White;
using TestStack.White.Configuration;
using TestStack.White.UIItems;
using TestStack.White.UIItems.TabItems;
using TestStack.White.UIItems.ListBoxItems;
using TestStack.White.UIItems.TreeItems;
using TestStack.White.UIItems.Finders;
using TestStack.White.AutomationElementSearch;
using TestStack.White.UIItems.WindowItems;
using TestStack.White.UIItems.WPFUIItems;
using TestStack.White.WindowsAPI;
using TestStack.White.InputDevices;

namespace WhiteRobotLibrary
{
    /// <summary>
    /// Robot Framework library implementation to drive the TestStack.White native Windows automation tool.
    /// The concept of xpath is extended here to be used as a reasonably generic locator for different screen objects.
    /// The complete list of Windows native Screen Objects (controlTypes) can be found from: http://msdn.microsoft.com/en-us/library/ms749005%28v=vs.110%29.aspx
    /// This Library requires using the NRobotRemote server, which is officially approved by the Robot Framework development team.
    /// Commadline startup via NRobotRemote is done as (all dlls in the same directory):
    /// | NRobotRemoteConsole.exe -p 8271 -k WhiteRobotLibrary.dll:WhiteRobotLibrary.WhiteKeywords |  
    /// Inclusion of the library in the test scripts is done as:
    /// | Library | Remote | http://127.0.0.1:8271/WhiteRobotLibrary/WhiteKeywords | WITH NAME | White |
    /// </summary>
    public class WhiteKeywords
    {
        private int waitAfterAction = 0;
        private Application testApp = null;
        private Window testAppWin = null;
        private List<String> locators = new List<String>() {"automationId", "className", "controlType", "text", "xpath"};
        private const char SEPARATOR = '=';
        private char MENUPATHSEPARATOR = ',';

        /// <summary>
        /// Returns a string containing a space-separated string
        /// of each process id of the given process name.
        /// | GetProcessIDsByName | iexplore |
        /// </summary>
        /// <param name="processName"></param>
        public string GetProcessIDsByName(String processName) {
            string prcsIds = ""; 
            var processes = Process.GetProcessesByName(processName);
            foreach (var process in processes) {
                prcsIds += process.Id.ToString() + " ";
            }
            return prcsIds.Trim();
        }

        /// <summary>
        /// Controls the execution speed.
        /// The given amount of time is waited between executing actions on objects.
        /// Value can be given in decimal format, such as 0.2 or 1.5 for optimum performance.
        /// | SetWaitAfterAction | 0.3 |
        /// </summary>
        /// <param name="waitSecs"></param>
        public void SetWaitAfterAction(String waitSecs) {
            this.waitAfterAction = 
                Convert.ToInt32(1000.0 * (Convert.ToDouble(waitSecs, System.Globalization.CultureInfo.InvariantCulture)));
        }

        /// <summary>
        /// Controls the dynamic waiting time for objects to appear on the screen. Default value is 5 seconds.
        /// | SetFindObjectTimeout | 15 |
        /// </summary>
        /// <param name="timeout"></param>
        public void SetFindObjectTimeout(String timeout) {
            CoreAppXmlConfiguration.Instance.BusyTimeout =
                Convert.ToInt32(1000.0 * (Convert.ToDouble(timeout, System.Globalization.CultureInfo.InvariantCulture)));
        }

        /// <summary>
        /// Launches the given application using the executable, e.g. iexplore.exe.
        /// If the executable is not within PATH, then it is necessary to include complete path as: C:\\temp\\iexplore.exe
        /// | LaunchApp | calc.exe               |
        /// | LaunchApp | C:\\temp\\iexplore.exe |
        /// </summary>
        /// <param name="appExecutable"></param>
        public void LaunchApp(String appExecutable) {
            testApp = Application.Launch(appExecutable);
        }

        /// <summary>
        /// Closes the application launched by LaunchApp keyword.
        /// The process is killed by force if tidy closing fails.
        /// | CloseApp |
        /// </summary>
        public void CloseApp() {
            testApp.Close();
        }

        /// <summary>
        /// Assuming the application is already launched,
        /// this keyword can be used for connecting to the already running application.
        /// The application can be identified using its process name or process ID.
        /// Note that the appName parameter is now given without the .exe type definition.
        /// To attach to Internet Explorer, use iexplore as the name.
        /// | AttachToApp | iexplore |
        /// | AttachToApp | 12345    |
        /// </summary>
        /// <param name="nameOrPSId"></param>
        public void AttachToApp(String nameOrPSId) {
            short psId;
            if (Int16.TryParse(nameOrPSId, out psId)) {
                testApp = Application.Attach((int)psId);
            } else {
                testApp = Application.Attach(nameOrPSId);
            }
        }

        /// <summary>
        /// After launching or attaching to the application, the tested window needs to be defined.
        /// With this keyword, the window can be accessed using the title of the window, e.g. Blank Page - Windows Internet Explorer
        /// | GetWindowByTitle | Blank Page - Windows Internet Explorer |
        /// </summary>
        /// <param name="windowTitle"></param>
        public void GetWindowByTitle(String windowTitle) {
            testAppWin = testApp.GetWindow(windowTitle);
        }

        /// <summary>
        /// A modal window is a child element to some other window. Get the main window first and then the modal window.
        /// With this keyword, the window can be accessed using the title of the window, e.g. Blank Page - Windows Internet Explorer
        /// | SelectModalWindowByTitle | Calculator |
        /// </summary>
        /// <param name="modalWindowTitle"></param>
        public void SelectModalWindowByTitle(String modalWindowTitle) {
            testAppWin = testAppWin.ModalWindow(modalWindowTitle);
        }

        /// <summary>
        /// After launching or attaching to the application, the tested window needs to be defined.
        /// With this keyword, the window can be accessed using the (partial) title of the window, e.g. Windows Internet Explorer or using index numbering.
        /// | SelectWindow | Windows Internet Explorer |
        /// | SelectWindow | 1                         |
        /// </summary>
        /// <param name="TitleOrIndex"></param>
        public void SelectWindow(String TitleOrIndex) {
            short indx;
            if (Int16.TryParse(TitleOrIndex, out indx)) {
                List<Window> windowList = testApp.GetWindows();
                testAppWin = windowList[indx - 1];
            } else {
                testAppWin = Desktop.Instance.Windows().Find(obj => obj.Title.Contains(TitleOrIndex));
            }        
        }

        /// <summary>
        /// Invokes an internal click event on the given UI-item. NOTE! Internal click events are not supported by all UI-items!
        /// Click is executed via internal event, screen can be locked because mouse is not used.
        /// The simple element locators are automationId, className, controlType, text and those are used as:
        /// automationId=myId, className=myClassName, controlType=MenuItem, text=someText.
        /// The xpath is much more versatile locator, which combines the controlType, simple element locator and index, for example
        /// | ClickObject | automationId=myId                                          |
        /// | ClickObject | text=Some Text                                             |
        /// | ClickObject | xpath=//Button[@text='Some text']                          |
        /// | ClickObject | xpath=//MenuItem[@className=File][@text='Some Item']       |
        /// | ClickObject | xpath=//MenuItem[1]                                        |
        /// | ClickObject | xpath=//Edit[@automationId='001234'][@text='Some Text'][2] |
        /// | ClickObject | xpath=//Pane[@text='x']/descendant::Button[@text='Ok'][3]  |
        /// </summary>
        /// <param name="elementId"></param>
        public void ClickObject(String elementId) {
            (findUIObject(elementId) as UIItem).RaiseClickEvent();
        }

        /// <summary>
        /// Clicks the given screen object using the left mouse button. The click is executed using a 'true' mouse event.
        /// The simple element locators are automationId, className, controlType, text and those are used as:
        /// automationId=myId, className=myClassName, controlType=MenuItem, text=someText.
        /// The xpath is much more versatile locator, which combines the controlType, simple element locator and index, for example
        /// | ClickObjectWithMouse | automationId=myId                                          |
        /// | ClickObjectWithMouse | text=Some Text                                             |
        /// | ClickObjectWithMouse | xpath=//Button[@text='Some text']                          |
        /// | ClickObjectWithMouse | xpath=//MenuItem[@className=File][@text='Some Item']       |
        /// | ClickObjectWithMouse | xpath=//MenuItem[1]                                        |
        /// | ClickObjectWithMouse | xpath=//Edit[@automationId='001234'][@text='Some Text'][2] |
        /// | ClickObjectWithMouse | xpath=//Pane[@text='x']/descendant::Button[@text='Ok'][3]  |
        /// </summary>
        /// <param name="elementId"></param>
        public void ClickObjectWithMouse(String elementId) {
            findUIObject(elementId).Click();
        }

        /// <summary>
        /// Clicks away from the given screen object by the given offset using the left mouse button.
        /// The click is executed using a 'true' mouse event. Horizontal positive direction is left,
        /// vertical positive direction is down.
        /// The simple element locators are automationId, className, controlType, text and those are used as:
        /// automationId=myId, className=myClassName, controlType=MenuItem, text=someText.
        /// The xpath is much more versatile locator, which combines the controlType, simple element locator and index, for example
        /// | ClickWithOffset | automationId=myId                                          | 100 | 100 |
        /// | ClickWithOffset | text=Some Text                                             | 10  | -10 |
        /// | ClickWithOffset | xpath=//Button[@text='Some text']                          |  0  | 100 |
        /// | ClickWithOffset | xpath=//MenuItem[@className=File][@text='Some Item']       |-20  | -20 |
        /// | ClickWithOffset | xpath=//MenuItem[1]                                        | 24  |  0  |
        /// | ClickWithOffset | xpath=//Edit[@automationId='001234'][@text='Some Text'][2] |  4  |  4  |
        /// | ClickWithOffset | xpath=//Pane[@text='x']/descendant::Button[@text='Ok'][3]  |  0  |  0  |
        /// </summary>
        /// <param name="elementId"></param>
        /// <param name="xOff"></param>
        /// <param name="yOff"></param>
        public void ClickWithOffset(String elementId, String xOff, String yOff) {
            System.Windows.Point po = findUIObject(elementId).Bounds.BottomRight;
            po.Offset(1.0 * Convert.ToInt32(xOff), 1.0 * Convert.ToInt32(yOff));
            testAppWin.Mouse.Click(po);
        }

        /// <summary>
        /// Double clicks the given screen object using the left mouse button. The click is executed using a 'true' mouse event.
        /// The simple element locators are automationId, className, controlType, text and those are used as:
        /// automationId=myId, className=myClassName, controlType=MenuItem, text=someText.
        /// The xpath is much more versatile locator, which combines the controlType, simple element locator and index, for example
        /// | DoubleClickObject | automationId=myId                                          |
        /// | DoubleClickObject | text=Some Text                                             |
        /// | DoubleClickObject | xpath=//Button[@text='Some text']                          |
        /// | DoubleClickObject | xpath=//MenuItem[@className=File][@text='Some Item']       |
        /// | DoubleClickObject | xpath=//MenuItem[1]                                        |
        /// | DoubleClickObject | xpath=//Edit[@automationId='001234'][@text='Some Text'][2] |
        /// | DoubleClickObject | xpath=//Pane[@text='x']/descendant::Button[@text='Ok'][3]  |
        /// </summary>
        /// <param name="elementId"></param>
        public void DoubleClickObject(String elementId) {
            findUIObject(elementId).DoubleClick();
        }

        /// <summary>
        /// Clicks the given screen object using the right mouse button. The click is executed using a 'true' mouse event.
        /// The simple element locators are automationId, className, controlType, text and those are used as:
        /// automationId=myId, className=myClassName, controlType=MenuItem, text=someText.
        /// The xpath is much more versatile locator, which combines the controlType, simple element locator and index, for example
        /// | RightClickObject | automationId=myId                                          |
        /// | RightClickObject | text=Some Text                                             |
        /// | RightClickObject | xpath=//Button[@text='Some text']                          |
        /// | RightClickObject | xpath=//MenuItem[@className=File][@text='Some Item']       |
        /// | RightClickObject | xpath=//MenuItem[1]                                        |
        /// | RightClickObject | xpath=//Edit[@automationId='001234'][@text='Some Text'][2] |
        /// | RightClickObject | xpath=//Pane[@text='x']/descendant::Button[@text='Ok'][3]  |
        /// </summary>
        /// <param name="elementId"></param>
        public void RightClickObject(String elementId) {
            findUIObject(elementId).RightClick();
        }

        /// <summary>
        /// Places a focus on the given screen object.
        /// The simple element locators are automationId, className, controlType, text and those are used as:
        /// automationId=myId, className=myClassName, controlType=MenuItem, text=someText.
        /// The xpath is much more versatile locator, which combines the controlType, simple element locator and index, for example
        /// | FocusOnObject | automationId=myId                                          |
        /// | FocusOnObject | text=Some Text                                             |
        /// | FocusOnObject | xpath=//Button[@text='Some text']                          |
        /// | FocusOnObject | xpath=//MenuItem[@className=File][@text='Some Item']       |
        /// | FocusOnObject | xpath=//MenuItem[1]                                        |
        /// | FocusOnObject | xpath=//Edit[@automationId='001234'][@text='Some Text'][2] |
        /// | FocusOnObject | xpath=//Pane[@text='x']/descendant::Button[@text='Ok'][3]  |
        /// </summary>
        /// <param name="elementId"></param>
        public void FocusOnObject(String elementId) {
            findUIObject(elementId).Focus();
        }

        /// <summary>
        /// Inserts text on on the given screen object, which should be of the control type Edit or Document.
        /// Any existing text from the text element is cleared before new text is inserted.
        /// Inserting is done via internal event, screen can be locked because keyboard is not used.
        /// The simple element locators are automationId, className, controlType, text and those are used as:
        /// automationId=myId, className=myClassName, controlType=MenuItem, text=someText.
        /// The xpath is much more versatile locator, which combines the controlType, simple element locator and index, for example
        /// | InsertTextToTextBox | automationId=myId                                          | Some text |
        /// | InsertTextToTextBox | controlType=Edit                                           | Some text |
        /// | InsertTextToTextBox | xpath=//Document[@automationId='someId']                   | Some text |
        /// | InsertTextToTextBox | xpath=//Edit[1]                                            | Some text |
        /// | InsertTextToTextBox | xpath=//Edit[@automationId='001234'][@text='Some Text'][2] | Some text |
        /// </summary>
        /// <param name="elementId"></param>
        /// <param name="text"></param>
        public void InsertTextToTextBox(String elementId, String text) {
            testAppWin.Get<TextBox>(GetBy(elementId)).BulkText = text;
        }

        /// <summary>
        /// Returns the visible text on on the given screen object, which should be of the control type Edit or Document.
        /// The simple element locators are automationId, className, controlType, text and those are used as:
        /// automationId=myId, className=myClassName, controlType=MenuItem, text=someText.
        /// The xpath is much more versatile locator, which combines the controlType, simple element locator and index, for example
        /// | ${text}= | ReadTextFromTextBox | automationId=myId                                          |
        /// | ${text}= | ReadTextFromTextBox | controlType=Edit                                           |
        /// | ${text}= | ReadTextFromTextBox | xpath=//Document[@automationId='someId']                   |
        /// | ${text}= | ReadTextFromTextBox | xpath=//Edit[1]                                            |
        /// | ${text}= | ReadTextFromTextBox | xpath=//Edit[@automationId='001234'][@text='Some Text'][2] |
        /// </summary>
        /// <param name="elementId"></param>
        /// <returns></returns>
        public String ReadTextFromTextBox(String elementId) {
            return testAppWin.Get<TextBox>(GetBy(elementId)).BulkText;
        }

        /// <summary>
        /// Writes text on the given screen object. The typing is executed using a 'true' keyboard event.
        /// The simple element locators are automationId, className, controlType, text and those are used as:
        /// automationId=myId, className=myClassName, controlType=MenuItem, text=someText.
        /// The xpath is much more versatile locator, which combines the controlType, simple element locator and index, for example
        /// | WriteTextOnObject | automationId=myId                                          | Some text |
        /// | WriteTextOnObject | text=Some Text                                             | Some text |
        /// | WriteTextOnObject | xpath=//Button[@text='Some text']                          | Some text |
        /// | WriteTextOnObject | xpath=//MenuItem[@className=File][@text='Some Item']       | Some text |
        /// | WriteTextOnObject | xpath=//MenuItem[1]                                        | Some text |
        /// | WriteTextOnObject | xpath=//Edit[@automationId='001234'][@text='Some Text'][2] | Some text |
        /// | WriteTextOnObject | xpath=//Pane[@text='x']/descendant::Button[@text='Ok'][3]  | Some text |
        /// </summary>
        /// <param name="elementId"></param>
        /// <param name="text"></param>
        public void WriteTextOnObject(String elementId, String text) {
            findUIObject(elementId).Enter(text);
        }

        /// <summary>
        /// Writes text on the current carret position. The typing is executed using a 'true' keyboard event.
        /// The keyword is used as
        /// | WriteText | Some text |
        /// </summary>
        /// <param name="text"></param>
        public void WriteText(String text) {
            testAppWin.Keyboard.Enter(text);
        }

        /// <summary>
        /// Presses the selected special key from the keyboard. The typing is executed using a 'true' keyboard event.
        /// The currently supported keys are: ENTER, BACKSPACE, TAB, ESC, ARROW_UP, ARROW_DOWN, 
		/// ARROW_RIGHT, ARROW_LEFT, PAGE_UP, PAGE_DOWN, DELETE, END, HOME, INSERT, SHIFT, CTRL, ALT and F1 - F12.
        /// The keyword is used as
        /// | PressKey | ENTER |
        /// </summary>
        /// <param name="key"></param>
        public void PressKey(String key) { 
            switch (key) { 
                case "ENTER":       testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.RETURN);    break;
                case "BACKSPACE":   testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.BACKSPACE); break;
                case "TAB":         testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.TAB);       break;
                case "ESC":         testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.ESCAPE);    break;
                case "ARROW_UP":    testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.UP);        break;
                case "ARROW_DOWN":  testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.DOWN);      break;
                case "ARROW_RIGHT": testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.RIGHT);     break;
                case "ARROW_LEFT":  testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.LEFT);      break;
                case "PAGE_UP":     testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.PAGEUP);    break;
                case "PAGE_DOWN":   testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.PAGEDOWN);  break;
                case "DELETE":      testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.DELETE);    break;
                case "END":         testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.END);       break;
                case "HOME":        testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.HOME);      break;
                case "INSERT":      testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.INSERT);    break;
                case "SHIFT":       testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.SHIFT);     break;
                case "CTRL":        testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.CONTROL);   break;
                case "ALT":         testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.ALT);       break;
                case "F1":          testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.F1);        break;
                case "F2":          testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.F2);        break;
                case "F3":          testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.F3);        break;
                case "F4":          testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.F4);        break;
                case "F5":          testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.F5);        break;
                case "F6":          testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.F6);        break;
                case "F7":          testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.F7);        break;
                case "F8":          testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.F8);        break;
                case "F9":          testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.F9);        break;
                case "F10":         testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.F10);       break;
                case "F11":         testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.F11);       break;
                case "F12":         testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.F12);       break;
                case "NUMLOCK":     testAppWin.Keyboard.PressSpecialKey(KeyboardInput.SpecialKeys.NUMLOCK);   break;
                default: break;
            }            
        }

        /// <summary>
        /// Presses the selected special key from the keyboard plus a normal key.
        /// The currently supported modifier keys are: CTRL, ALT, SHIFT and CTRL+ALT.
        /// The keyword is used as
        /// | PressKeyCombination | CTRL      | a |
        /// | PressKeyCombination | CTRL      | v |
        /// | PressKeyCombination | ALT       | u |
        /// | PressKeyCombination | SHIFT     | x |
        /// | PressKeyCombination | CTRL+ALT  | a |
        /// </summary>
        /// <param name="modifier"></param>
        /// <param name="keys"></param>
        public void PressKeyCombination(String modifier, String keys) {
            switch (modifier) {
                case "CTRL":   
                    testAppWin.Keyboard.HoldKey(KeyboardInput.SpecialKeys.CONTROL);
                    testAppWin.Keyboard.Enter(keys);
                    testAppWin.Keyboard.LeaveKey(KeyboardInput.SpecialKeys.CONTROL);
                    break;
                case "ALT":
                    testAppWin.Keyboard.HoldKey(KeyboardInput.SpecialKeys.ALT);
                    testAppWin.Keyboard.Enter(keys);
                    testAppWin.Keyboard.LeaveKey(KeyboardInput.SpecialKeys.ALT);
                    break;
                case "SHIFT":
                    testAppWin.Keyboard.HoldKey(KeyboardInput.SpecialKeys.SHIFT);
                    testAppWin.Keyboard.Enter(keys);
                    testAppWin.Keyboard.LeaveKey(KeyboardInput.SpecialKeys.SHIFT);
                    break;
                case "CTRL+ALT":
                    testAppWin.Keyboard.HoldKey(KeyboardInput.SpecialKeys.CONTROL);
                    testAppWin.Keyboard.HoldKey(KeyboardInput.SpecialKeys.ALT);
                    testAppWin.Keyboard.Enter(keys);
                    testAppWin.Keyboard.LeaveKey(KeyboardInput.SpecialKeys.ALT);
                    testAppWin.Keyboard.LeaveKey(KeyboardInput.SpecialKeys.CONTROL);
                    break;
                default: break;
            }
        }

        /// <summary>
        /// Typically the Name property of the object is equal to the visible text of the object.
        /// This keyword returns the value of the Name property of the given object.
        /// The simple element locators are automationId, className, controlType, text and those are used as:
        /// automationId=myId, className=myClassName, controlType=MenuItem, text=someText.
        /// The xpath is much more versatile locator, which combines the controlType, simple element locator and index, for example
        /// | ${text}= | ReadTextFromObject | automationId=myId                                          |
        /// | ${text}= | ReadTextFromObject | text=Some Text                                             |
        /// | ${text}= | ReadTextFromObject | xpath=//Button[@text='Some text']                          |
        /// | ${text}= | ReadTextFromObject | xpath=//MenuItem[@className=File][@text='Some Item']       |
        /// | ${text}= | ReadTextFromObject | xpath=//MenuItem[1]                                        |
        /// | ${text}= | ReadTextFromObject | xpath=//Edit[@automationId='001234'][@text='Some Text'][2] |
        /// | ${text}= | ReadTextFromObject | xpath=//Pane[@text='x']/descendant::Button[@text='Ok'][3]  |
        /// </summary>
        /// <param name="elementId"></param>
        /// <returns></returns>
        public String ReadTextFromObject(String elementId) {
            return findUIObject(elementId).Name;
        }

        /// <summary>
        /// Selects a tab page with the given title or index from the given tab element
        /// The simple element locators are automationId, className, controlType, text and those are used as:
        /// automationId=myId, className=myClassName, controlType=MenuItem, text=someText.
        /// The xpath is much more versatile locator, which combines the controlType, simple element locator and index, for example
        /// | SelectTabPage | automationId=myTab                     | TabTitle  |
        /// | SelectTabPage | controlType=Tab                        | 1         |
        /// | SelectTabPage | xpath=//Tab[1]                         | 3         |
        /// | SelectTabPage | xpath=//Tab[@automationId='001234'][2] | TabTitle  |
        /// </summary>
        /// <param name="elementId"></param>
        /// <param name="TitleOrIndex"></param>
        /// <returns></returns>
        public void SelectTabPage(String elementId, String TitleOrIndex) {            
            short indx;
            Tab tabi = testAppWin.Get<Tab>(GetBy(elementId));
            if (Int16.TryParse(TitleOrIndex, out indx)) {
                tabi.SelectTabPage(indx);
            } else {
                tabi.SelectTabPage(TitleOrIndex);
            }
        }

        /// <summary>
        /// Places a check mark in a CheckBox object, which is identified by the given object locator.
        /// Uses an internal toggle event, no mouse interaction needed.
        /// The simple element locators are automationId, className, controlType, text and those are used as:
        /// automationId=myId, className=myClassName, controlType=MenuItem, text=someText.
        /// The xpath is much more versatile locator, which combines the controlType, simple element locator and index, for example
        /// | SelectCheckBox | automationId=myCheckBox                     |
        /// | SelectCheckBox | controlType=CheckBox                        |
        /// | SelectCheckBox | xpath=//CheckBox[1]                         |
        /// | SelectCheckBox | xpath=//CheckBox[@automationId='001234'][2] |
        /// </summary>
        /// <param name="elementId"></param>
        /// <returns></returns>
        public void SelectCheckBox(String elementId) {
            ToggleableItem togg = new ToggleableItem(findUIObject(elementId) as UIItem);
            while (ToggleState.On != togg.State)
                togg.Toggle();
        }

        /// <summary>
        /// Removes a check mark from a CheckBox object, which is identified by the given object locator.
        /// Uses an internal toggle event, no mouse interaction needed.
        /// The simple element locators are automationId, className, controlType, text and those are used as:
        /// automationId=myId, className=myClassName, controlType=MenuItem, text=someText.
        /// The xpath is much more versatile locator, which combines the controlType, simple element locator and index, for example
        /// | UnSelectCheckBox | automationId=myCheckBox                     |
        /// | UnSelectCheckBox | controlType=CheckBox                        |
        /// | UnSelectCheckBox | xpath=//CheckBox[1]                         |
        /// | UnSelectCheckBox | xpath=//CheckBox[@automationId='001234'][2] |
        /// </summary>
        /// <param name="elementId"></param>
        /// <returns></returns>
        public void UnSelectCheckBox(String elementId) {
            ToggleableItem togg = new ToggleableItem(findUIObject(elementId) as UIItem);
            while (ToggleState.Off != togg.State)
                togg.Toggle();
        }

        /// <summary>
        /// Selects a RadioButton item, which is identified by the given object locator.
        /// Selection is done by invoking an internal selection event. No mouse interaction is needed.
        /// The simple element locators are automationId, className, controlType, text and those are used as:
        /// automationId=myId, className=myClassName, controlType=MenuItem, text=someText.
        /// The xpath is much more versatile locator, which combines the controlType, simple element locator and index, for example
        /// | SelectRadioButton | automationId=myRadioButton                     |
        /// | SelectRadioButton | controlType=RadioButton                        |
        /// | SelectRadioButton | xpath=//RadioButton[1]                         |
        /// | SelectRadioButton | xpath=//RadioButton[@automationId='001234'][2] |
        /// </summary>
        /// <param name="elementId"></param>
        /// <returns></returns>
        public void SelectRadioButton(String elementId) {
            testAppWin.Get<RadioButton>(GetBy(elementId)).Select();       
        }

        /// <summary>
        /// Selects a MenuItem from the main menubar.
        /// Selection is done by invoking an true mouse event.
        /// The menu item is identified by giving the path from the main menu item and the items should be separated using a comma as shown in examples below:
        /// | SelectMenuItem | File                     |
        /// | SelectMenuItem | Help,About               |
        /// </summary>
        /// <param name="menuPath"></param>
        /// <returns></returns>
        public void SelectMenuItem(String menuPath) {
            (testAppWin.MenuBar.MenuItem(menuPath.Split(MENUPATHSEPARATOR)) as UIItem).RaiseClickEvent();
        }

        /// <summary>
        /// Selects an item from a tree element, which is identified by the given object locator.
        /// Selection is done by invoking an internal selection event, but if that does not work, then mouse interaction is used.
        /// The tree items are identified using the path from the root element, the items should be separated by a comma as shown in the examples.
        /// The simple element locators are automationId, className, controlType, text and those are used as:
        /// automationId=myId, className=myClassName, controlType=MenuItem, text=someText.
        /// The xpath is much more versatile locator, which combines the controlType, simple element locator and index, for example
        /// | SelectTreeItem | myTree                       | Root,Level1        |
        /// | SelectTreeItem | xpath=//Tree[@text='myTree'] | Root,Level1,Level2 |
        /// </summary>
        /// <param name="elementId"></param>
        /// <param name="treePath"></param>
        /// <returns></returns>
        public void SelectTreeItem(String elementId, String treePath) {
            testAppWin.Get<Tree>(GetBy(elementId)).Node(treePath.Split(MENUPATHSEPARATOR)).Select();
        }

        /// <summary>
        /// Selects an item from a drop down menu, which is identified by the given object locator.
        /// The item can be identified by using its visible text or index in the drop down menu.
        /// Selection is done by invoking an internal selection event. No mouse interaction is needed.
        /// The simple element locators are automationId, className, controlType, text and those are used as:
        /// automationId=myId, className=myClassName, controlType=MenuItem, text=someText.
        /// The xpath is much more versatile locator, which combines the controlType, simple element locator and index, for example
        /// | SelectDropDownMenuItem | myComboBox                      | Pickme |
        /// | SelectDropDownMenuItem | xpath=//ComboBox[@text='combo'] | 2      |
        /// </summary>
        /// <param name="elementId"></param>
        /// <param name="nameOrIndex"></param>
        /// <returns></returns>
        public void SelectDropDownMenuItem(String elementId, String nameOrIndex) {
            short indx;
            ComboBox cbox = testAppWin.Get<ComboBox>(GetBy(elementId));
            if (Int16.TryParse(nameOrIndex, out indx)) {
                cbox.Select(indx);
            } else {
                cbox.Select(nameOrIndex);
            }
        }

        /// <summary>
        /// Saves a screenshot of the current screen.
        /// Creates a directory with name "scrshots" relative to the .dll path and saves the screenshot there.
        /// | GetScreenshot |
        /// </summary>
        public void GetScreenshot() {
            string fname = DateTime.Now.ToString("MMddHHmmss") + ".png";
            Directory.CreateDirectory(".\\scrshots");
            Desktop.TakeScreenshot(".\\scrshots\\"+fname, ImageFormat.Png);
            // APPENDING TO LOG FILE DOES NOT WORK
            //Trace.TraceInformation("*HTML* <img src=./srcshots/'" + fname + "'></img>");
        }

        /// <summary>
        /// Keyword for testing purposes to see what patterns are supported by the given screen object
        /// </summary>
        /// <param name="elementId"></param>
        public void LogSupportedPatterns(String elementId)
        {
            AutomationElement ae = testAppWin.GetElement(GetBy(elementId));
            AutomationPattern[] patterns = ae.GetSupportedPatterns();
            foreach (AutomationPattern patt in patterns)
            {
                Trace.TraceInformation(patt.ProgrammaticName);
                Trace.TraceInformation(Automation.PatternName(patt));
                Console.WriteLine(patt.ProgrammaticName);
                Console.WriteLine(Automation.PatternName(patt));
            }
        }

        /// <summary>
        /// Keyword for testing purposes to see what elements are found under some top-level element.
        /// </summary>
        /// <param name="elementId"></param>
        public void LogObject(String elementId) {
            testAppWin.Get(GetBy(elementId)).LogStructure();
        }

        /// <summary>
        /// Test method for quering if the given element is in Enabled state.
        /// If not Enabled, an exception is thrown.
        /// </summary>
        /// <param name="elementId"></param>
        public void IsObjectEnabled(String elementId) {
            if ( !(findUIObject(elementId).AutomationElement.Current.IsEnabled) )
                throw new ElementNotEnabledException("Given object is not enabled");
        }

        /// <summary>
        /// Test method for quering if the given element is visible.
        /// If not visible, an exception is thrown.
        /// </summary>
        /// <param name="elementId"></param>
        public void IsObjectVisible(String elementId) {
            if ( !(findUIObject(elementId).AutomationElement.Current.IsOffscreen) )
                throw new ElementNotAvailableException("Given object is not visible");
        }


        /* Class internal helper methods */
        private SearchCriteria GetBy(String id) {
            System.Threading.Thread.Sleep(waitAfterAction);
            SearchCriteria sCrit = null;
            String by = "";
            int x = id.IndexOf(SEPARATOR);
            if (x > 0 && locators.Contains(id.Substring(0, x))) {
                by = id.Substring(0, x);
                id = id.Substring(x + 1);
            }
            if (by.Equals("automationId"))
                sCrit = SearchCriteria.ByAutomationId(id);
            else if (by.Equals("className"))
                sCrit = SearchCriteria.ByClassName(id);
            else if (by.Equals("controlType"))
                sCrit = SearchCriteria.ByControlType(GetControlTypes(id));
            else if (by.Equals("text"))
                sCrit = SearchCriteria.ByText(id);
            else if (by.Equals("xpath")) {
                int startIndex = 2; int indexx = 0;
                String element = id.Substring(startIndex, (id.IndexOf('[') - startIndex)).Trim();
                sCrit = SearchCriteria.ByControlType(GetControlTypes(element));
                while ((startIndex = id.IndexOf("[@")) >= 0) {
                    id = id.Substring(startIndex+2);
                    startIndex = id.IndexOf("='");
                    String attribute = id.Substring(0, startIndex).Trim();
                    String param = id.Substring(startIndex+2, (id.IndexOf("']") - (startIndex+2)));
                    if (attribute.Equals("automationId"))
                        sCrit = sCrit.AndAutomationId(param);
                    else if (attribute.Equals("className"))
                        sCrit = sCrit.AndByClassName(param);
                    else if (attribute.Equals("text"))
                        sCrit = sCrit.AndByText(param);    
                }

                if ( (startIndex = id.IndexOf('[')) >= 0 ) {
                    indexx = Convert.ToInt16( id.Substring(startIndex+1, id.LastIndexOf(']') - (startIndex+1)).Trim() );
                }
                
                if (indexx > 0) sCrit = sCrit.AndIndex(indexx-1);
            }
            else
                sCrit = SearchCriteria.ByText(id);

            return sCrit;
        }

        private ControlType GetControlTypes(String controlTypeString) {
            ControlType.Button.ToString();
            AutomationControlType typeEnum;
            bool result = Enum.TryParse(controlTypeString, true, out typeEnum);
            if (result) typeEnum = (AutomationControlType)Enum.Parse(typeof(AutomationControlType), controlTypeString);
            int enumId = (int)typeEnum + 50000;
            return ControlType.LookupById(enumId);
        }


        private IUIItem findUIObject(String elementId) {
            string[] elements = new string[] { elementId, "", "" ,"" };
            // applies only to special case xpath=//Pane[@text='ss']>>>Edit>>>Settings>>>1
            if (elementId.Contains(">>>")) {
                elements = elementId.Split(new string[] { ">>>" }, StringSplitOptions.RemoveEmptyEntries);
                AutomationElement parent = (testAppWin.Get(GetBy(elements[0])) as UIItem).AutomationElement;
                return simpleDescendantSearch(parent, elements[1], elements[2], Convert.ToInt32(elements[3]));
            }
            else if (elementId.StartsWith("xpath=") && !elementId.Contains(">>>")) {
                string elementString = elementId.Substring(8);
                string[] els = elementString.Split(new string[] { "/descendant::" }, StringSplitOptions.RemoveEmptyEntries);
                short i = 0;
                foreach (string s in els) {
                    elements[i] = "xpath=//" + s;
                    //Trace.TraceInformation(elements[i]);
                    i++;
                }
            }
            IUIItem ui = testAppWin.Get(GetBy(elements[0]));
            Trace.TraceInformation("ELEM 0: " + elements[0]);
            for (short i = 1; i < elements.Length; i++) {
                if (elements[i].Length > 0) { ui = ui.Get(GetBy(elements[i])); }
            }
            return ui;
        }

        // This was created as a hack in a need for partial match for UIobject name property
        // usage: xpath=//Pane[@text='ss']>>>Edit>>>Settings>>>1  , where Edit is ctrlType, Settings is the name and 1 is the index
        private IUIItem simpleDescendantSearch(AutomationElement parent, String ctrlType, String text, int indx)
        {
            AutomationElementCollection descendants = parent.FindAll(TreeScope.Descendants, Condition.TrueCondition);
            UIItem ui = null;
            AutomationElement ae = null;
            int i = 0;
            foreach (AutomationElement de in descendants) {
                Trace.TraceInformation("simpleDescendantSearch: " + de.Current.ControlType.ProgrammaticName + " - " + de.Current.Name + " - " + de.Current.IsOffscreen);
                if (de.Current.ControlType.ProgrammaticName.Substring(12).Equals(ctrlType) && de.Current.Name.Contains(text)) {
                    i++;
                }
                if (i == indx) { ae = de; i++; }
            }
            if (ae != null) { ui = new UIItem(ae, testAppWin.ActionListener); }
            return ui as IUIItem;
        }

    }
}
