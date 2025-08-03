import win32com.client

mtb = win32com.client.Dispatch("Mtb.Application")
mtb.UserInterface.Visible = True
proj = mtb.ActiveProject
proj.ExecuteCommand('Print "Hello from Python via COM!"')
proj.SaveAs("d:/auto_project.mpj")
mtb.Quit()
