using System;
using System.Diagnostics;
using System.Windows.Automation;
using System.Windows.Automation.Peers;
using System.Collections.Generic;
using TestStack.White;
using TestStack.White.Configuration;
using TestStack.White.UIItems;
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

        /// <summary>
        /// Controls the execution speed.
        /// The given amount of time is waited between executing actions on objects.
        /// Value can be given in decimal format, such as 0.2 or 1.5 for optimum performance.
        /// | SetWaitAfterActions | 0.3 |
        /// </summary>
        /// <param name="waitSecs"></param>
        public void SetWaitAfterActions(String waitSecs) {
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
        /// Note that the appName parameter is now given without the .exe type definition.
        /// To attach to the Internet Explorer, use iexplore as the appName.
        /// | AttachToApp | iexplore |
        /// </summary>
        /// <param name="appName"></param>
        public void AttachToApp(String appName) {
            testApp = Application.Attach(appName);
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
        /// | GetModalWindowByTitle | Calculator |
        /// </summary>
        /// <param name="modalWindowTitle"></param>
        public void GetModalWindowByTitle(String modalWindowTitle) {
            testAppWin = testAppWin.ModalWindow(modalWindowTitle);
        }

        /// <summary>
        /// Same as GetWindowByTitle, but accepts partial title text.
        /// | GetWindowByTitle | Internet Explorer |
        /// </summary>
        /// <param name="partialTitle"></param>
        public void GetWindowByPartialTitle(String partialTitle) {
            testAppWin = Desktop.Instance.Windows().Find(obj => obj.Title.Contains(partialTitle));
        }
        
        /// <summary>
        /// After launching or attaching to the application, the tested window needs to be defined.
        /// With this keyword, the windows inside the application can be accessed using an index.
        /// This is useful when the window does not have a title. Values 1,2,3, ... are accepted as index.
        /// | GetWindowByIndex | 2 |
        /// </summary>
        /// <param name="index"></param>
        public void GetWindowByIndex(String index) {
            int indVal = Convert.ToInt16(index);
            List<Window> windowList = testApp.GetWindows();
            testAppWin = windowList[indVal - 1];
        }

        /// <summary>
        /// Clicks the given screen object using the left mouse button. The click is executed using a 'true' mouse event.
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
        /// Writes text on on the given screen object.
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
            //testAppWin.Get<TextBox>(GetBy(elementId)).BulkText = text;
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
        /// Test method for quering if the given element is in Enabled state.
        /// If not Enabled, an exception is thrown.
        /// </summary>
        /// <param name="elementId"></param>
        public void IsObjectEnabled(String elementId) {
            if ( !(testAppWin.Get(GetBy(elementId)).Enabled) )
                throw new ElementNotEnabledException("Given object is not enabled");
        }

        /// <summary>
        /// Test method for quering if the given element is Visible.
        /// If not Visible, an exception is thrown.
        /// </summary>
        /// <param name="elementId"></param>
        public void IsObjectVisible(String elementId) {
            if (!(testAppWin.Get(GetBy(elementId)).Visible))
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
