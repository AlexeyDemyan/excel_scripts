echo Clearing EVERYTHING...
set _my_datetime=%date%_%time%
set _my_datetime=%_my_datetime: =_%
set _my_datetime=%_my_datetime::=%
set _my_datetime=%_my_datetime:/=_%
set _my_datetime=%_my_datetime:.=_%
echo %_my_datetime%
mkdir orders_backup\%_my_datetime%
move ../Documents/*.lock orders_backup\%_my_datetime%
move ../Documents/*.hive orders_backup\%_my_datetime%
pause