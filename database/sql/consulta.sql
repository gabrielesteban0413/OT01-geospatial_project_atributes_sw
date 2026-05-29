SELECT * FROM staging.merge_corporativo_sync
WHERE cable_asphia IN (
    SELECT regexp_split_to_table(
        '
		

70235-34
71239-68


',
        '\n'
    )
);