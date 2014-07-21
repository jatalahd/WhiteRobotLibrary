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
        /// </summary>
        /// <param name="elementId"></param>
        public void ClickObject(String elementId) {
            (testAppWin.Get(GetBy(elementId)) as UIItem).RaiseClickEvent();
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
        /// </summary>
        /// <param name="elementId"></param>
        public void ClickObjectWithMouse(String elementId) {
            testAppWin.Get(GetBy(elementId)).Click();
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
        /// </summary>
        /// <param name="elementId"></param>
        public void DoubleClickObject(String elementId) {
            testAppWin.Get(GetBy(elementId)).DoubleClick();
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
        /// </summary>
        /// <param name="elementId"></param>
        public void RightClickObject(String elementId) {
            testAppWin.Get(GetBy(elementId)).RightClick();
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
        /// </summary>
        /// <param name="elementId"></param>
        public void FocusOnObject(String elementId) {
            testAppWin.Get(GetBy(elementId)).Focus();
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
        /// Writes text on on the given screen object. The typing is executed using a 'true' keyboard event.
        /// The simple element locators are automationId, className, controlType, text and those are used as:
        /// automationId=myId, className=myClassName, controlType=MenuItem, text=someText.
        /// The xpath is much more versatile locator, which combines the controlType, simple element locator and index, for example
        /// | WriteTextOnObject | automationId=myId                                          | Some text |
        /// | WriteTextOnObject | text=Some Text                                             | Some text |
        /// | WriteTextOnObject | xpath=//Button[@text='Some text']                          | Some text |
        /// | WriteTextOnObject | xpath=//MenuItem[@className=File][@text='Some Item']       | Some text |
        /// | WriteTextOnObject | xpath=//MenuItem[1]                                        | Some text |
        /// | WriteTextOnObject | xpath=//Edit[@automationId='001234'][@text='Some Text'][2] | Some text |
        /// </summary>
        /// <param name="elementId"></param>
        /// <param name="text"></param>
        public void WriteTextOnObject(String elementId, String text) {
            testAppWin.Get(GetBy(elementId)).Enter(text);
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
        /// </summary>
        /// <param name="elementId"></param>
        /// <returns></returns>
        public String ReadTextFromObject(String elementId) {
            return testAppWin.Get(GetBy(elementId)).Name;
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
            UIItem item = (testAppWin.Get(GetBy(elementId)) as UIItem);
            ToggleableItem togg = new ToggleableItem(item);
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
            UIItem item = (testAppWin.Get(GetBy(elementId)) as UIItem);
            ToggleableItem togg = new ToggleableItem(item);
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
        public void LogObject(String elementId)
        {
            (testAppWin.Get(GetBy(elementId)) as UIItem).LogStructure();
        }

        /// <summary>
        /// Test method for quering if the given element is in Enabled state.
        /// If not Enabled, an exception is thrown.
        /// </summary>
        /// <param name="elementId"></param>
        public void IsObjectEnabled(String elementId) {
            if ( !(testAppWin.GetElement(GetBy(elementId)).Current.IsEnabled) )
                throw new ElementNotEnabledException("Given object is not enabled");
        }

        /// <summary>
        /// Test method for quering if the given element is visible.
        /// If not visible, an exception is thrown.
        /// </summary>
        /// <param name="elementId"></param>
        public void IsObjectVisible(String elementId) {
            if (!(testAppWin.GetElement(GetBy(elementId)).Current.IsOffscreen))
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

    }
}
