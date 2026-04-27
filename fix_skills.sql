UPDATE your_table_name SET general_skill = TRIM(REGEXP_REPLACE(general_skill, ' \|\| \|\| .*', '')) WHERE general_skill LIKE '% || || %';
