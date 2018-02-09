Set Service = GetObject("winmgmts:{impersonationLevel=impersonate}!root/Microsoft/HomeNet")
Set enumset = Service.InstancesOf("HNet_ConnectionProperties") 

for each Instance in enumSet

if Instance.IsIcsPrivate = true then
Instance.IsIcsPrivate = false
Instance.put_()
end if

Next
