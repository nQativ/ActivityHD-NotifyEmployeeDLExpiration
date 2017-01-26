Set activity = CreateObject("ActivityOAS.Activity")
activity.ServerAddress = "appserver3"
activity.Connect
Set company = activity.Companies("Demo Fair")
company.Connect
company.RunAction "Payroll", "Notes", "Macro", _
 "<p>" & _
 "<Macro Name='Notify DL renewal needed'/>" & _
 "</p>"