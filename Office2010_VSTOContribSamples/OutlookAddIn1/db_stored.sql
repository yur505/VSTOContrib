CREATE PROCEDURE `new_procedure` (OUT s int)

BEGIN
	select count(*) into s from city
END