Set Service = GetObject("winmgmts:{impersonationLevel=impersonate}!root/Microsoft/HomeNet")
Set enumset = Service.InstancesOf("HNet_ConnectionProperties") 

for each Instance in enumSet

if Instance.IsIcsPublic = true then
Instance.IsIcsPublic = false
Instance.put_()
end if

Next
