# Excel tables union

The program is designed for excel files union by built in template. Originally it was developed for prices files union, you can use it as you like, if you properly fill in the template.

The template is a built in code python dictionary, looks like this:
```json
{
	'Main_column': ['probable_value_1', 'probable_value_2', ... , 'probable_value_n'],
	'Other_column_1': ['probable_value_1_1', 'probable_value_1_2', ... , 'probable_value_1_n'],
	...
	'Other_column_m': ['probable_value_m_1', 'probable_value_m_2', ... , 'probable_value_m_n'],
}
```

List of probable values is used for find matches in excel tables and specify the names of columns. If one of your excel tables doesnt contain a row with column names - you will have one single column (Main_column) as a result. Columns that are not specified in this dictionary will not be counted in result table.

The script also can proccess multi-sheet files and lots of tables in one sheet. For each founded table there will be unique column name specified. 

There is some sort of garbage eraser algorithm: all rows and columns that have more than 85% and 95% of empty cells in certain range will not be counted in results.

# How to use?

