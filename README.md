# Windows Privacy

Since there are a lot of users out there who deny switching to Windows 10 because of privacy concerns, Microsoft decided that also those users data needs to be collected through updates like Telemetry also for Windows 7 and 8.

I've created this one to be used as simple as possible.

#### What it does
- Removes updates as defined in INI file and sets them as hidden for Windows Update
- Deactivation of tasks as defined in INI file
- Removal or disabling of services as defined in INI file

#### Startup handling (simplicity)
- If started from an network location, it will copy itself as a random name to `%TEMP%`, execute, and upon completion remove itself from `%TEMP%`. This will also happen when the drive executed from is a network location.
- Checks if administrative permissions are available and if needed, restarts itself elevated to ask for permission
- If run by an doubleclick relaunches itself for use with cscript
- Can be used as part of an Domain GPO (check `boolUserInteraction`)

#### Usage
Proper execution from windows command is always:
> cscript.exe windowsprivacy.vbs

For the end user, a simple doubleclick should be enough

For any upcoming "spying" updates, they can be easily added in INI file

#### INI Structure (Tasks/Updates/Services)

##### Updates

- Linewise, where **kb** is optional

	> **kb**123456=Some Description
  
##### Tasks

- Linewise, path to task

	> You can get a list of tasks from commandline by executing `schtasks` 

##### Services

- Linewise, only `disable` or `delete` are processed

	> servicename=`action`

##### Registry

Can add or delete registry values

Format:

`KEY`||`SUBKEY`||`TYPE` or `ACTION`||`VALUE`

Supported Keys:

- `HKEY_LOCAL_MACHINE`
- `HKEY_CURRENT_USER`
- `HKEY_CLASSES_ROOT`
- `HKEY_USERS`

Accepted Types:

- `REG_DWORD`
- `REG_SZ`
- `REG_EXPAND_SZ`

Action is one of the TYPE or `DELETE`

