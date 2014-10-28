using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Windows.Input;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;

namespace GorillaDocs.CodedUITests
{
    [CodedUITest]
    public class RibbonTests
    {
        [TestMethod]
        public void CodedUITestMethod1()
        {
            ApplicationUnderTest.Launch(@"C:\Program Files (x86)\Microsoft Office\Office15\WinWord.exe");
            OpenDocument("c:\\Repos\\GorillaDocs\\GorillaDocs.CodedUITests\\Sample data\\Blank.docx");

            var fileNoteButton = GetFileNoteButton();
            Assert.IsTrue(fileNoteButton.Exists);
        }

        static WinButton GetFileNoteButton()
        {
            var window = GetWordWindow();
            var ribbonWindow = GetWordRibbon(window);

            var templatesToolbar = new WinToolBar(ribbonWindow);
            templatesToolbar.SearchProperties[WinToolBar.PropertyNames.Name] = "Templates";
            templatesToolbar.WindowTitles.Add("Blank.docx - Word");

            var fileNoteButton = new WinButton(templatesToolbar);
            fileNoteButton.SearchProperties[WinButton.PropertyNames.Name] = "BCC FileNotes";
            fileNoteButton.WindowTitles.Add("Blank.docx - Word");
            return fileNoteButton;
        }
        static void OpenDocument(string filename)
        {
            var window = GetWordWindow();
            var ribbonWindow = GetWordRibbon(window);

            var ribbonControl = new WinControl(ribbonWindow);
            ribbonControl.SearchProperties[UITestControl.PropertyNames.Name] = "Ribbon";
            ribbonControl.SearchProperties[UITestControl.PropertyNames.ControlType] = "PropertyPage";
            ribbonControl.WindowTitles.Add("Document1 - Word");

            var fileButton = new WinButton(ribbonControl);
            fileButton.SearchProperties[WinButton.PropertyNames.Name] = "File Tab";
            fileButton.WindowTitles.Add("Document1 - Word");

            Mouse.Click(fileButton);

            var fileMenu = new WinMenuBar(ribbonWindow);
            fileMenu.SearchProperties[WinMenu.PropertyNames.Name] = "File";
            fileMenu.WindowTitles.Add("Document1 - Word");

            var openTab = new WinTabPage(fileMenu);
            openTab.SearchProperties[WinTabPage.PropertyNames.Name] = "Open";
            openTab.WindowTitles.Add("Document1 - Word");

            Mouse.Click(openTab);

            var itemGroup = new WinGroup(window);
            itemGroup.WindowTitles.Add("Document1 - Word");

            var computerTab = new WinTabPage(itemGroup);
            computerTab.SearchProperties[WinTabPage.PropertyNames.Name] = "Computer";
            Mouse.Click(computerTab);

            var pickaFolderGroup = new WinGroup(window);
            pickaFolderGroup.SearchProperties[WinControl.PropertyNames.Name] = "Pick a Folder";

            var browseButton = new WinButton(pickaFolderGroup);
            browseButton.SearchProperties[WinButton.PropertyNames.Name] = "Browse";
            Mouse.Click(browseButton);

            var openDlg = new WinWindow();
            openDlg.SearchProperties[WinWindow.PropertyNames.ClassName] = "#32770";
            openDlg.WindowTitles.Add("Open");

            var openDlg1 = new WinWindow(openDlg);
            openDlg1.SearchProperties[WinWindow.PropertyNames.ControlId] = "1148";
            openDlg1.WindowTitles.Add("Open");

            var fileNameCombo = new WinComboBox(openDlg1);
            fileNameCombo.SearchProperties[WinComboBox.PropertyNames.Name] = "File name:";
            fileNameCombo.EditableItem = filename;

            var fileNameEdit = new WinEdit(openDlg1);
            fileNameEdit.SearchProperties[WinEdit.PropertyNames.Name] = "File name:";

            Keyboard.SendKeys(fileNameEdit, "{Enter}", ModifierKeys.None);
        }
        static WinWindow GetWordWindow()
        {
            var window = new WinWindow();
            window.SearchProperties[WinWindow.PropertyNames.ClassName] = "OpusApp";
            window.WindowTitles.Add("Document1 - Word");
            return window;
        }
        static WinWindow GetWordRibbon(WinWindow window)
        {
            var ribbonWindow = new WinWindow(window);
            ribbonWindow.SearchProperties[WinWindow.PropertyNames.AccessibleName] = "Ribbon";
            ribbonWindow.SearchProperties[WinWindow.PropertyNames.ClassName] = "NetUIHWND";
            ribbonWindow.WindowTitles.Add("Document1 - Word");
            return ribbonWindow;
        }

        #region Stuff I don't think I need

        #region Additional test attributes

        // You can use the following additional attributes as you write your tests:

        ////Use TestInitialize to run code before running each test 
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{        
        //    // To generate code for this test, select "Generate Code for Coded UI Test" from the shortcut menu and select one of the menu items.
        //}

        ////Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{        
        //    // To generate code for this test, select "Generate Code for Coded UI Test" from the shortcut menu and select one of the menu items.
        //}

        #endregion

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext { get; set; }

        public UIMap UIMap
        {
            get
            {
                if ((this.map == null))
                {
                    this.map = new UIMap();
                }

                return this.map;
            }
        }
        private UIMap map;
        #endregion
    }
}
