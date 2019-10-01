-- Project Redirection for Clarion 10.0 just to get the Clarion10\Bin\Debug DLLs

-- This RED has a [Copy] *.dll = %BIN%\debug
{include clarion_DebugDllCopy.red}
                                     
-- this is the normal RED --                                     
{include %REDDIR%\%REDNAME%}



