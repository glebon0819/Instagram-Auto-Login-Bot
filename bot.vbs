Set browser = CreateObject("InternetExplorer.application")
browser.Statusbar = false
browser.Toolbar = false
browser.Menubar = false
browser.Visible = true

browser.Navigate("https://www.instagram.com/accounts/login/")

wscript.Sleep(5000)

browser.Document.All.Item("username").Value = "username"
browser.Document.All.Item("password").Value = "password"

wscript.Sleep(100)

Set buttons = browser.Document.GetElementsByTagName("button")

buttons.Item(0).Click