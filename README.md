# Windows Privacy

Since there are a lot of users out there who deny switching to Windows 10 because of privacy concerns, Microsoft decided that also those users data needs to be collected through updates like Telemetry also for Windows 7 and 8.

I've created this one to be used as simple as possible.

#### What it does

###### Startup handling (Simplicity)

- If started from an network location, it will copy itself as a random name to `%TEMP%`, execute, and upon completion remove itself from `%TEMP%`
- Checks if administrative permissions are available and if needed, restarts itself elevated to ask for permission
- If run by an doubleclick relaunches itself for use with cscript
- Can be used as part of an Domain GPO (check `boolUserInteraction`)
 
###### System related (Tasks/Updates/Services)
- Removes updates as defined in INI file and sets them as hidden for Windows Update
- Deactivation of tasks as defined in INI file
- Removal or disabling of services as defined in INI file

#### Execution

Proper execution is always:
> cscript.exe windowsprivacy.vbs

For the end user, a simple doubleclick should be enough

For any upcoming "spying" updates, they can be easily added in INI file

#### INI file structure

###### Updates

- Linewise, where **kb** is optional
  > **kb**123456=Some Description
  
###### Tasks

- Linewise, path to task
  > You can get a list of tasks from commandline by executing `schtasks` 

###### Services

- Linewise, only `disable` or `delete` are processed
  > servicename=`action`

Maybe more to come (registry)...
