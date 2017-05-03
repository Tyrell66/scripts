@ECHO OFF
FOR %%D IN ("DC=oa,DC=pnrad,DC=net" Forestroot "DC=dwain,DC=infra") DO (
  FOR /F "delims=; usebackq" %%Q IN (`dsquery computer %%D -limit 0 -name *%1*`) DO (
    dsquery computer %%Q -o rdn
    echo %%Q
	dsget computer %%Q -desc | FIND /V "  desc" | FIND /V "dsget succeeded"
	echo.
  )
)
